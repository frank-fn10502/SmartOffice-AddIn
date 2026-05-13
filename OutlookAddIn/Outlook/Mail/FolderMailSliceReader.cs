using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
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
        internal async Task HandleFolderMailsSliceAsync(OutlookCommand cmd)
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
                int maxCount = req.MaxCount > 0
                    ? Math.Max(1, Math.Min(500, req.MaxCount))
                    : 30;

                List<MailItemDto> mails = await _outlookThread.InvokeAsync(() => ReadFolderMailsSlice(req, maxCount));
                int total = await PushFolderMailsBatchesAsync(cmd, req, batchSize, mails ?? new List<MailItemDto>());

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
                    Message = "fetch_folder_mails_slice error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex)
                });
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_folder_mails_slice error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
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
            int batchSize,
            int maxCount)
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
                try { filtered.Sort("[ReceivedTime]", true); } catch { }

                foreach (var obj in filtered)
                {
                    if (total >= maxCount) break;

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

        private async Task<int> PushFolderMailsBatchesAsync(
            OutlookCommand cmd,
            OutlookCommandFolderMailsSliceRequest req,
            int batchSize,
            List<MailItemDto> allMails)
        {
            string folderMailsId = req.FolderMailsId ?? "";
            int total = allMails?.Count ?? 0;
            int sequence = 1;

            if (total == 0)
            {
                await PushFolderMailsBatchAsync(
                    cmd,
                    req,
                    folderMailsId,
                    sequence,
                    new List<MailItemDto>(),
                    true,
                    req.CompleteOnSlice);
                return total;
            }

            for (int offset = 0; offset < total; offset += batchSize)
            {
                int count = Math.Min(batchSize, total - offset);
                var batch = allMails.GetRange(offset, count);
                bool isLastBatch = offset + count >= total;
                await PushFolderMailsBatchAsync(
                    cmd,
                    req,
                    folderMailsId,
                    sequence++,
                    batch,
                    isLastBatch,
                    isLastBatch && req.CompleteOnSlice);
            }

            return total;
        }

        private List<MailItemDto> ReadFolderMailsSlice(
            OutlookCommandFolderMailsSliceRequest req,
            int maxCount)
        {
            var results = new List<MailItemDto>();

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
                        "ReadFolderMailsSlice: folderEntryId could not be resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null)
                    return results;

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                items = folder.Items;

                string filterExpr = BuildFolderMailsFilter(req.ReceivedFrom, req.ReceivedTo);
                filtered = filterExpr != null ? items.Restrict(filterExpr) : items;
                try { filtered.Sort("[ReceivedTime]", true); } catch { }

                foreach (var obj in filtered)
                {
                    if (results.Count >= maxCount) break;

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

                    if (dto != null)
                        results.Add(dto);
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
                    OutlookDateFilter.FormatItemsDateTime(receivedFrom.Value),
                    OutlookDateFilter.FormatItemsDateTime(receivedTo.Value));

            if (receivedFrom.HasValue)
                return string.Format("[ReceivedTime] >= '{0}'",
                    OutlookDateFilter.FormatItemsDateTime(receivedFrom.Value));

            return string.Format("[ReceivedTime] <= '{0}'",
                OutlookDateFilter.FormatItemsDateTime(receivedTo.Value));
        }
    }
}
