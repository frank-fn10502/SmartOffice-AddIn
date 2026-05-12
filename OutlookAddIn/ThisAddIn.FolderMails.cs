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
                List<MailItemDto> mails = null;
                Exception folderEx = null;

                _chatPane.Invoke((Action)(() =>
                {
                    try { mails = ReadFolderMailsSlice(req); }
                    catch (Exception ex) { folderEx = ex; }
                }));

                if (folderEx != null)
                    throw folderEx;

                // Clamp resultBatchSize to contract range [3, 5]; default 5.
                int batchSize = req.ResultBatchSize > 0
                    ? Math.Max(3, Math.Min(5, req.ResultBatchSize))
                    : 5;

                var allMails = mails ?? new List<MailItemDto>();
                int total = allMails.Count;
                int sequence = 1;

                if (total == 0)
                {
                    await _signalRClient.PushFolderMailsSliceResultAsync(new FolderMailsSliceResultDto
                    {
                        FolderMailsId = folderMailsId,
                        CommandId = cmd.Id,
                        ParentCommandId = req.ParentCommandId ?? "",
                        Sequence = sequence,
                        SliceIndex = req.SliceIndex,
                        SliceCount = req.SliceCount,
                        Reset = req.ResetResults,
                        IsFinal = req.CompleteOnSlice,
                        IsSliceComplete = true,
                        Mails = new List<MailItemDto>(),
                        Message = ""
                    });
                }
                else
                {
                    for (int offset = 0; offset < total; offset += batchSize)
                    {
                        int count = Math.Min(batchSize, total - offset);
                        var batch = allMails.GetRange(offset, count);
                        bool isLastBatch = (offset + count >= total);

                        await _signalRClient.PushFolderMailsSliceResultAsync(new FolderMailsSliceResultDto
                        {
                            FolderMailsId = folderMailsId,
                            CommandId = cmd.Id,
                            ParentCommandId = req.ParentCommandId ?? "",
                            Sequence = sequence++,
                            SliceIndex = req.SliceIndex,
                            SliceCount = req.SliceCount,
                            Reset = req.ResetResults && offset == 0,
                            IsFinal = isLastBatch && req.CompleteOnSlice,
                            IsSliceComplete = isLastBatch,
                            Mails = batch,
                            Message = ""
                        });
                    }
                }

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
        private List<MailItemDto> ReadFolderMailsSlice(OutlookCommandFolderMailsSliceRequest req)
        {
            var results = new List<MailItemDto>();
            Outlook.MAPIFolder folder = null;
            Outlook.Items items = null;
            Outlook.Items filtered = null;
            try
            {
                // Prefer storeId + folderEntryId; fall back to storeId + folderPath with log warning.
                if (!string.IsNullOrEmpty(req.FolderEntryId))
                    folder = GetFolderByEntryIdInStore(req.StoreId, req.FolderEntryId);

                if (folder == null && !string.IsNullOrEmpty(req.FolderPath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "ReadFolderMailsSlice: folderEntryId could not be resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null) return results;

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                items = folder.Items;

                // Apply Items.Restrict for received time if specified
                string filterExpr = BuildFolderMailsFilter(req.ReceivedFrom, req.ReceivedTo);
                filtered = filterExpr != null ? items.Restrict(filterExpr) : items;

                foreach (var obj in filtered)
                {
                    var mail = obj as Outlook.MailItem;
                    if (mail == null)
                    {
                        if (obj != null) try { Marshal.ReleaseComObject(obj); } catch { }
                        continue;
                    }
                    try
                    {
                        var dto = ReadSingleMailDto(mail, currentFolderPath, false);
                        if (dto != null) results.Add(dto);
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(mail); } catch { } }
                }
            }
            finally
            {
                if (filtered != null && !ReferenceEquals(filtered, items))
                    try { Marshal.ReleaseComObject(filtered); } catch { }
                if (items != null) try { Marshal.ReleaseComObject(items); } catch { }
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }
            return results;
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
