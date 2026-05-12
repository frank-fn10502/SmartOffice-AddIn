using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        // ────────────────────────────────────────────────────────────────────────────────
        // fetch_folder_mails_slice
        // Directly enumerates a single Outlook folder using Items / Items.Restrict.
        // MUST NOT call Application.AdvancedSearch or any Outlook content search.
        // ────────────────────────────────────────────────────────────────────────────────
        private async Task HandleFolderMailsSliceAsync(OutlookCommand cmd)
        {
            var req = cmd.FolderMailsSliceRequest;

            if (req == null
                || string.IsNullOrEmpty(req.StoreId)
                || string.IsNullOrEmpty(req.FolderEntryId)
                || string.IsNullOrEmpty(req.FolderPath))
            {
                await _signalRClient.CompleteFolderMailsSliceAsync(new FolderMailsCompleteDto
                {
                    FolderMailsId = req?.FolderMailsId ?? "",
                    CommandId = cmd.Id,
                    ParentCommandId = req?.ParentCommandId ?? "",
                    Success = false,
                    Message = "fetch_folder_mails_slice failed: storeId, folderEntryId and folderPath are required"
                });
                return;
            }

            string folderMailsId = req.FolderMailsId ?? "";

            // Signal start of slice (only on first slice)
            if (req.ResetResults)
            {
                await _signalRClient.BeginFolderMailsAsync(new FolderMailsSliceResultDto
                {
                    FolderMailsId = folderMailsId,
                    CommandId = cmd.Id,
                    ParentCommandId = req.ParentCommandId ?? "",
                    Sequence = 0,
                    SliceIndex = req.SliceIndex,
                    SliceCount = req.SliceCount,
                    Reset = true,
                    IsFinal = false,
                    IsSliceComplete = false,
                    Mails = new List<MailItemDto>(),
                    Message = "slice started"
                });
            }

            try
            {
                // Clamp resultBatchSize to contract range [3, 5]; default 5.
                int batchSize = req.ResultBatchSize > 0
                    ? Math.Max(3, Math.Min(5, req.ResultBatchSize))
                    : 5;

                int total = await PushFolderMailsSliceFromOutlookAsync(cmd, req, batchSize);

                if (req.CompleteOnSlice)
                {
                    await _signalRClient.CompleteFolderMailsSliceAsync(new FolderMailsCompleteDto
                    {
                        FolderMailsId = folderMailsId,
                        CommandId = cmd.Id,
                        ParentCommandId = req.ParentCommandId ?? "",
                        Success = true,
                        Message = "fetch_folder_mails_slice completed"
                    });
                }

                await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                    $"fetch_folder_mails_slice completed. Items: {total}");
            }
            catch (Exception ex)
            {
                await _signalRClient.CompleteFolderMailsSliceAsync(new FolderMailsCompleteDto
                {
                    FolderMailsId = folderMailsId,
                    CommandId = cmd.Id,
                    ParentCommandId = req.ParentCommandId ?? "",
                    Success = false,
                    Message = "fetch_folder_mails_slice error: " + SanitizeExceptionForLog(ex)
                });
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_folder_mails_slice error: " + SanitizeExceptionForLog(ex));
            }
        }

        /// <summary>
        /// Enumerates a single Outlook folder using Items / Items.Restrict.
        /// Never calls AdvancedSearch or any Outlook content search.
        /// Returns metadata-only MailItemDtos; body is never read.
        /// Must be called on the UI (STA) thread.
        /// </summary>
        private async Task<int> PushFolderMailsSliceFromOutlookAsync(
            OutlookCommand cmd,
            OutlookCommandFolderMailsSliceRequest req,
            int batchSize)
        {
            string folderMailsId = req.FolderMailsId ?? "";
            int total = 0;
            int sequence = 1;
            var batch = new List<MailItemDto>(batchSize);

            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            Outlook.Items filtered = null;
            try
            {
                if (!string.IsNullOrEmpty(req.FolderEntryId))
                    folder = GetFolderByEntryIdInStore(req.StoreId, req.FolderEntryId);

                if (folder == null && !string.IsNullOrEmpty(req.FolderPath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "PushFolderMailsSliceFromOutlookAsync: folderEntryId could not be resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null)
                    return total;

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                items = folder.Items;

                string filterExpr = BuildFolderMailsFilter(req.ReceivedFrom, req.ReceivedTo);
                filtered = filterExpr != null ? items.Restrict(filterExpr) : items;

                foreach (var obj in filtered)
                {
                    var mail = obj as Outlook.MailItem;
                    MailItemDto dto = null;
                    if (mail == null)
                    {
                        if (obj != null) try { Marshal.ReleaseComObject(obj); } catch { }
                        continue;
                    }

                    try { dto = ReadMailListMetadataDto(mail, currentFolderPath); }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(mail); } catch { } }

                    if (dto == null)
                        continue;

                    batch.Add(dto);
                    total++;
                    if (batch.Count >= batchSize)
                    {
                        await PushFolderMailsBatchAsync(cmd, req, folderMailsId, sequence++, batch, false, false);
                        batch = new List<MailItemDto>(batchSize);
                        System.Windows.Forms.Application.DoEvents();
                    }
                }
            }
            finally
            {
                if (filtered != null && !ReferenceEquals(filtered, items))
                    try { Marshal.ReleaseComObject(filtered); } catch { }
                if (items != null) try { Marshal.ReleaseComObject(items); } catch { }
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }

            await PushFolderMailsBatchAsync(
                cmd,
                req,
                folderMailsId,
                sequence,
                batch,
                true,
                req.CompleteOnSlice);

            return total;
        }

        private async Task PushFolderMailsBatchAsync(
            OutlookCommand cmd,
            OutlookCommandFolderMailsSliceRequest req,
            string folderMailsId,
            int sequence,
            List<MailItemDto> mails,
            bool isSliceComplete,
            bool isFinal)
        {
            await _signalRClient.PushFolderMailsSliceResultAsync(new FolderMailsSliceResultDto
            {
                FolderMailsId = folderMailsId,
                CommandId = cmd.Id,
                ParentCommandId = req.ParentCommandId ?? "",
                Sequence = sequence,
                SliceIndex = req.SliceIndex,
                SliceCount = req.SliceCount,
                Reset = req.ResetResults && sequence == 1,
                IsFinal = isFinal,
                IsSliceComplete = isSliceComplete,
                Mails = mails ?? new List<MailItemDto>(),
                Message = ""
            });
        }

        private static string BuildFolderMailsFilter(DateTime? receivedFrom, DateTime? receivedTo)
        {
            if (!receivedFrom.HasValue && !receivedTo.HasValue) return null;

            if (receivedFrom.HasValue && receivedTo.HasValue)
                return string.Format("[ReceivedTime] >= '{0}' AND [ReceivedTime] <= '{1}'",
                    receivedFrom.Value.ToString("MM/dd/yyyy HH:mm"),
                    receivedTo.Value.ToString("MM/dd/yyyy HH:mm"));

            if (receivedFrom.HasValue)
                return string.Format("[ReceivedTime] >= '{0}'",
                    receivedFrom.Value.ToString("MM/dd/yyyy HH:mm"));

            return string.Format("[ReceivedTime] <= '{0}'",
                receivedTo.Value.ToString("MM/dd/yyyy HH:mm"));
        }
    }
}
