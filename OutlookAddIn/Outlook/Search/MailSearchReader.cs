using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Handles the fetch_mail_search_slice command.
        /// Executes a DASL-based AdvancedSearch within a single Outlook folder as instructed by Hub.
        /// Hub is responsible for cross-folder scheduling; AddIn must not scan globally.
        /// Body is never included (metadata-only per contract).
        /// </summary>
        internal async Task HandleMailSearchSliceAsync(OutlookCommand cmd)
        {
            var req = cmd.MailSearchSliceRequest;
            if (req == null
                || string.IsNullOrEmpty(req.StoreId)
                || string.IsNullOrEmpty(req.FolderEntryId)
                || string.IsNullOrEmpty(req.FolderPath))
            {
                await _signalRClient.CompleteMailSearchSliceAsync(new MailSearchCompleteDto
                {
                    SearchId = req?.SearchId ?? "",
                    CommandId = cmd.Id,
                    ParentCommandId = req?.ParentCommandId ?? "",
                    Success = false,
                    Message = "fetch_mail_search_slice failed: storeId, folderEntryId and folderPath are required"
                });
                return;
            }

            string searchId = req.SearchId ?? "";

            // Signal start of slice (only on first slice)
            if (req.ResetSearchResults)
            {
                await _signalRClient.BeginMailSearchAsync(new MailSearchSliceResultDto
                {
                    SearchId = searchId,
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

                int total = 0;
                string mode = req.ExecutionMode?.ToLower() ?? "outlook_search";
                if (mode == "items_filter")
                {
                    List<MailItemDto> mails = await _outlookThread.InvokeAsync(() => ExecuteMailSearchSliceItemsFilter(req));
                    total = await PushMailSearchBatchesAsync(cmd, req, searchId, batchSize, mails ?? new List<MailItemDto>());
                }
                else
                {
                    List<MailItemDto> mails = await _outlookThread.InvokeAsync(() => ExecuteMailSearchSliceAdvancedSearch(req));
                    total = await PushMailSearchBatchesAsync(cmd, req, searchId, batchSize, mails ?? new List<MailItemDto>());
                }

                if (req.CompleteSearchOnSlice)
                {
                    await _signalRClient.CompleteMailSearchSliceAsync(new MailSearchCompleteDto
                    {
                        SearchId = searchId,
                        CommandId = cmd.Id,
                        ParentCommandId = req.ParentCommandId ?? "",
                        Success = true,
                        Message = "fetch_mail_search_slice completed"
                    });
                }

                await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                    $"fetch_mail_search_slice completed. Items: {total}");
            }
            catch (Exception ex)
            {
                await _signalRClient.CompleteMailSearchSliceAsync(new MailSearchCompleteDto
                {
                    SearchId = searchId,
                    CommandId = cmd.Id,
                    ParentCommandId = req.ParentCommandId ?? "",
                    Success = false,
                    Message = "fetch_mail_search_slice error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex)
                });
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_mail_search_slice error: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
        }

        /// <summary>
        /// Dispatches to items_filter (Items.Restrict) or outlook_search (AdvancedSearch)
        /// based on executionMode.
        /// Returns metadata-only MailItemDtos; body is never read here.
        /// </summary>
        private List<MailItemDto> ExecuteMailSearchSlice(OutlookCommandMailSearchSliceRequest req)
        {
            string mode = req.ExecutionMode?.ToLower() ?? "outlook_search";
            if (mode == "items_filter")
                return ExecuteMailSearchSliceItemsFilter(req);
            return ExecuteMailSearchSliceAdvancedSearch(req);
        }

        private async Task<int> PushMailSearchSliceItemsFilterFromOutlookAsync(
            OutlookCommand cmd,
            OutlookCommandMailSearchSliceRequest req,
            int batchSize)
        {
            string searchId = req.SearchId ?? "";
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
                        "ExecuteMailSearchSliceItemsFilter: folderEntryId not resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null)
                {
                    await PushMailSearchBatchAsync(cmd, req, searchId, sequence, batch, true, req.CompleteSearchOnSlice);
                    return total;
                }

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                items = folder.Items;
                string filterExpr = BuildItemsFilterExpr(req);
                filtered = filterExpr != null ? items.Restrict(filterExpr) : items;

                var textFields = (req.TextFields != null && req.TextFields.Count > 0)
                    ? req.TextFields
                    : new List<string> { "subject" };
                bool searchSubject = textFields.Contains("subject");
                bool searchSender = textFields.Contains("sender");

                foreach (var obj in filtered)
                {
                    var mail = obj as Outlook.MailItem;
                    MailItemDto dto = null;
                    if (mail == null)
                    {
                        if (obj != null) try { Marshal.ReleaseComObject(obj); } catch { }
                        continue;
                    }
                    try
                    {
                        if (!MatchesItemsFilter(mail, req, searchSubject, searchSender))
                            continue;

                        dto = ReadMailListMetadataDto(mail, currentFolderPath);
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(mail); } catch { } }

                    if (dto == null)
                        continue;

                    batch.Add(dto);
                    total++;
                    if (batch.Count >= batchSize)
                    {
                        await PushMailSearchBatchAsync(cmd, req, searchId, sequence++, batch, false, false);
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

            await PushMailSearchBatchAsync(cmd, req, searchId, sequence, batch, true, req.CompleteSearchOnSlice);
            return total;
        }

        private async Task PushMailSearchBatchAsync(
            OutlookCommand cmd,
            OutlookCommandMailSearchSliceRequest req,
            string searchId,
            int sequence,
            List<MailItemDto> mails,
            bool isSliceComplete,
            bool isFinal)
        {
            await _signalRClient.PushMailSearchSliceResultAsync(new MailSearchSliceResultDto
            {
                SearchId = searchId,
                CommandId = cmd.Id,
                ParentCommandId = req.ParentCommandId ?? "",
                Sequence = sequence,
                SliceIndex = req.SliceIndex,
                SliceCount = req.SliceCount,
                Reset = req.ResetSearchResults && sequence == 1,
                IsFinal = isFinal,
                IsSliceComplete = isSliceComplete,
                Mails = mails ?? new List<MailItemDto>(),
                Message = ""
            });
        }

        private async Task<int> PushMailSearchBatchesAsync(
            OutlookCommand cmd,
            OutlookCommandMailSearchSliceRequest req,
            string searchId,
            int batchSize,
            List<MailItemDto> allMails)
        {
            int total = allMails?.Count ?? 0;
            int sequence = 1;

            if (total == 0)
            {
                await PushMailSearchBatchAsync(
                    cmd,
                    req,
                    searchId,
                    sequence,
                    new List<MailItemDto>(),
                    true,
                    req.CompleteSearchOnSlice);
                return total;
            }

            for (int offset = 0; offset < total; offset += batchSize)
            {
                int count = Math.Min(batchSize, total - offset);
                var batch = allMails.GetRange(offset, count);
                bool isLastBatch = offset + count >= total;
                await PushMailSearchBatchAsync(
                    cmd,
                    req,
                    searchId,
                    sequence++,
                    batch,
                    isLastBatch,
                    isLastBatch && req.CompleteSearchOnSlice);
            }

            return total;
        }

        /// <summary>
        /// items_filter mode: uses folder Items / Items.Restrict for metadata-only filtering.
        /// Never calls AdvancedSearch. Handles subject/sender keyword, category, attachment,
        /// flag state, read state and received time filters.
        /// </summary>
        private List<MailItemDto> ExecuteMailSearchSliceItemsFilter(OutlookCommandMailSearchSliceRequest req)
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
                        "ExecuteMailSearchSliceItemsFilter: folderEntryId not resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null) return results;

                string currentFolderPath = "";
                try { currentFolderPath = folder.FolderPath ?? ""; } catch { }

                items = folder.Items;

                // Build Items.Restrict filter from time range only (DASL property filter)
                string filterExpr = BuildItemsFilterExpr(req);
                filtered = filterExpr != null ? items.Restrict(filterExpr) : items;

                // Determine fields to search keyword in
                var textFields = (req.TextFields != null && req.TextFields.Count > 0)
                    ? req.TextFields
                    : new List<string> { "subject" };
                bool searchSubject = textFields.Contains("subject");
                bool searchSender = textFields.Contains("sender");

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
                        if (!MatchesItemsFilter(mail, req, searchSubject, searchSender))
                            continue;
                        var dto = ReadMailListMetadataDto(mail, currentFolderPath);
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

        private static string BuildItemsFilterExpr(OutlookCommandMailSearchSliceRequest req)
        {
            var parts = new List<string>();

            if (req.ReceivedFrom.HasValue)
                parts.Add(string.Format("[ReceivedTime] >= '{0}'",
                    OutlookDateFilter.FormatItemsDateTime(req.ReceivedFrom.Value)));
            if (req.ReceivedTo.HasValue)
                parts.Add(string.Format("[ReceivedTime] <= '{0}'",
                    OutlookDateFilter.FormatItemsDateTime(req.ReceivedTo.Value)));

            return parts.Count > 0 ? string.Join(" AND ", parts) : null;
        }

        private static bool MatchesItemsFilter(
            Outlook.MailItem mail,
            OutlookCommandMailSearchSliceRequest req,
            bool searchSubject,
            bool searchSender)
        {
            // keyword filter
            if (!string.IsNullOrEmpty(req.Keyword))
            {
                bool keywordHit = false;
                if (searchSubject)
                {
                    string subj = ""; try { subj = mail.Subject ?? ""; } catch { }
                    if (subj.IndexOf(req.Keyword, StringComparison.OrdinalIgnoreCase) >= 0) keywordHit = true;
                }
                if (!keywordHit && searchSender)
                {
                    string sName = ""; try { sName = mail.SenderName ?? ""; } catch { }
                    string sAddr = ""; try { sAddr = mail.SenderEmailAddress ?? ""; } catch { }
                    if (sName.IndexOf(req.Keyword, StringComparison.OrdinalIgnoreCase) >= 0 ||
                        sAddr.IndexOf(req.Keyword, StringComparison.OrdinalIgnoreCase) >= 0) keywordHit = true;
                }
                if (!keywordHit) return false;
            }

            // category filter
            if (req.CategoryNames != null && req.CategoryNames.Count > 0)
            {
                string cats = ""; try { cats = mail.Categories ?? ""; } catch { }
                bool catHit = false;
                foreach (var cat in req.CategoryNames)
                {
                    if (cats.IndexOf(cat, StringComparison.OrdinalIgnoreCase) >= 0) { catHit = true; break; }
                }
                if (!catHit) return false;
            }

            // attachment filter
            if (req.HasAttachments.HasValue)
            {
                int attCount = 0;
                Outlook.Attachments atts = null;
                try
                {
                    atts = mail.Attachments;
                    if (atts != null) attCount = atts.Count;
                }
                catch { }
                finally { if (atts != null) try { Marshal.ReleaseComObject(atts); } catch { } }

                if (req.HasAttachments.Value && attCount == 0) return false;
                if (!req.HasAttachments.Value && attCount > 0) return false;
            }

            // flag state filter
            if (!string.IsNullOrEmpty(req.FlagState) && req.FlagState != "any")
            {
                Outlook.OlFlagStatus fs = Outlook.OlFlagStatus.olNoFlag;
                try { fs = mail.FlagStatus; } catch { }
                if (req.FlagState == "flagged" && fs != Outlook.OlFlagStatus.olFlagMarked) return false;
                if (req.FlagState == "unflagged" && fs != Outlook.OlFlagStatus.olNoFlag) return false;
            }

            // read state filter
            if (!string.IsNullOrEmpty(req.ReadState) && req.ReadState != "any")
            {
                bool unread = true;
                try { unread = mail.UnRead; } catch { }
                if (req.ReadState == "unread" && !unread) return false;
                if (req.ReadState == "read" && unread) return false;
            }

            return true;
        }

        /// <summary>
        /// outlook_search mode: Executes Outlook AdvancedSearch with a DASL filter.
        /// Returns metadata-only MailItemDtos; body is never read here.
        /// </summary>
        private List<MailItemDto> ExecuteMailSearchSliceAdvancedSearch(OutlookCommandMailSearchSliceRequest req)
        {
            var results = new List<MailItemDto>();
            Outlook.MAPIFolder folder = null;
            Outlook.Search search = null;
            try
            {
                // Prefer storeId + folderEntryId; fall back to storeId + folderPath with warning.
                if (!string.IsNullOrEmpty(req.FolderEntryId))
                    folder = GetFolderByEntryIdInStore(req.StoreId, req.FolderEntryId);

                if (folder == null && !string.IsNullOrEmpty(req.FolderPath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "ExecuteMailSearchSlice: folderEntryId could not be resolved; using folderPath fallback");
                    folder = GetFolderByPathInStore(req.StoreId, req.FolderPath);
                }

                if (folder == null) return results;

                // AdvancedSearch expects a DASL filter string directly (no @SQL= prefix).
                string dasl = BuildDaslFilter(req);
                // Outlook AdvancedSearch scope must wrap each folder path in single quotes.
                // Example: '\\Mailbox - User\\Inbox'
                string scope = BuildAdvancedSearchScope(folder.FolderPath);
                string tag = "SmartOfficeSlice_" + req.SearchId;

                Exception searchError = null;
                var mre = new ManualResetEventSlim(false);

                // Register event handler BEFORE calling AdvancedSearch to avoid race condition
                // where a small folder completes synchronously before we subscribe.
                void OnAdvancedSearchComplete(Outlook.Search searchObject)
                {
                    if (searchObject.Tag != tag) return;
                    this.Application.AdvancedSearchComplete -= OnAdvancedSearchComplete;
                    try
                    {
                        var resultSet = searchObject.Results;
                        int count = resultSet?.Count ?? 0;
                        for (int i = 1; i <= count; i++)
                        {
                            Outlook.MailItem mail = null;
                            try
                            {
                                mail = resultSet[i] as Outlook.MailItem;
                                if (mail == null) continue;
                                var dto = ReadMailListMetadataDto(mail, folder.FolderPath);
                                if (dto != null) results.Add(dto);
                            }
                            catch { }
                            finally { if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { } }
                        }
                        if (resultSet != null) try { Marshal.ReleaseComObject(resultSet); } catch { }
                    }
                    catch (Exception ex) { searchError = ex; }
                    finally { mre.Set(); }
                }

                void StartSearchWithFilter(string filter)
                {
                    this.Application.AdvancedSearchComplete += OnAdvancedSearchComplete;
                    try
                    {
                        search = this.Application.AdvancedSearch(scope, filter, false, tag);
                    }
                    catch
                    {
                        // If AdvancedSearch itself throws, unsubscribe before re-throwing
                        this.Application.AdvancedSearchComplete -= OnAdvancedSearchComplete;
                        throw;
                    }
                }

                try
                {
                    StartSearchWithFilter(dasl);
                }
                catch (COMException)
                {
                    // Some stores/providers reject complex DASL combinations.
                    // Retry with a minimal, broadly-compatible subject filter.
                    var fallbackDasl = BuildFallbackDaslFilter(req);
                    if (string.Equals(fallbackDasl, dasl, StringComparison.Ordinal))
                        throw;
                    StartSearchWithFilter(fallbackDasl);
                }

                // Wait up to 30 s; pump messages so AdvancedSearchComplete can fire on the STA thread.
                var deadline = DateTime.UtcNow.AddSeconds(30);
                while (!mre.IsSet && DateTime.UtcNow < deadline)
                {
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(50);
                }
                bool completed = mre.IsSet;
                if (!completed)
                {
                    this.Application.AdvancedSearchComplete -= OnAdvancedSearchComplete;
                    try { search.Stop(); } catch { }
                    throw new TimeoutException("AdvancedSearch timed out after 30 s");
                }
                if (searchError != null) throw searchError;
            }
            catch
            {
                throw;
            }
            finally
            {
                if (search != null) try { Marshal.ReleaseComObject(search); } catch { }
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }
            return results;
        }

        private static string BuildDaslFilter(OutlookCommandMailSearchSliceRequest req)
        {
            var parts = new List<string>();

            // --- keyword ---
            if (!string.IsNullOrWhiteSpace(req.Keyword))
            {
                var fields = req.TextFields != null && req.TextFields.Count > 0
                    ? req.TextFields
                    : new List<string> { "subject" };

                var kwParts = new List<string>();
                foreach (var field in fields)
                {
                    string daslProp = FieldToDaslProperty(field);
                    if (daslProp != null)
                        kwParts.Add($"{daslProp} like '%{EscapeDaslValue(req.Keyword)}%'");
                }
                if (kwParts.Count == 1)
                    parts.Add(kwParts[0]);
                else if (kwParts.Count > 1)
                    parts.Add("(" + string.Join(" OR ", kwParts) + ")");
            }

            // NOTE: Outlook AdvancedSearch requires the "@SQL=" prefix for all DASL filters.

            // --- categories ---
            if (req.CategoryNames != null && req.CategoryNames.Count > 0)
            {
                var catParts = new List<string>();
                foreach (var cat in req.CategoryNames)
                    catParts.Add($"\"urn:schemas-microsoft-com:office:office#Keywords\" like '%{EscapeDaslValue(cat)}%'");
                if (catParts.Count == 1)
                    parts.Add(catParts[0]);
                else
                    parts.Add("(" + string.Join(" OR ", catParts) + ")");
            }

            // --- attachments ---
            if (req.HasAttachments.HasValue)
            {
                parts.Add(req.HasAttachments.Value
                    ? "\"urn:schemas:httpmail:hasattachment\" = 1"
                    : "\"urn:schemas:httpmail:hasattachment\" = 0");
            }

            // --- flag state ---
            if (!string.IsNullOrEmpty(req.FlagState) && req.FlagState != "any")
            {
                // olFlagMarked = 2 (flagged), olNoFlag = 0 (unflagged)
                if (req.FlagState == "flagged")
                    parts.Add("\"urn:schemas:httpmail:messageflag\" = 2");
                else if (req.FlagState == "unflagged")
                    parts.Add("\"urn:schemas:httpmail:messageflag\" = 0");
            }

            // --- read state ---
            if (!string.IsNullOrEmpty(req.ReadState) && req.ReadState != "any")
            {
                if (req.ReadState == "unread")
                    parts.Add("\"urn:schemas:httpmail:read\" = 0");
                else if (req.ReadState == "read")
                    parts.Add("\"urn:schemas:httpmail:read\" = 1");
            }

            // --- time range ---
            if (req.ReceivedFrom.HasValue)
                parts.Add($"\"urn:schemas:httpmail:datereceived\" >= '{FormatDaslDateTime(req.ReceivedFrom.Value)}'");
            if (req.ReceivedTo.HasValue)
                parts.Add($"\"urn:schemas:httpmail:datereceived\" <= '{FormatDaslDateTime(req.ReceivedTo.Value)}'");

            // No conditions: return all mail-class items.
            if (parts.Count == 0)
                return "@SQL=\"urn:schemas:httpmail:subject\" like '%'";

            return "@SQL=" + string.Join(" AND ", parts);
        }

        private static string FieldToDaslProperty(string field)
        {
            switch (field?.ToLower())
            {
                case "subject": return "\"urn:schemas:httpmail:subject\"";
                case "sender":  return "\"urn:schemas:httpmail:fromname\"";
                case "body":    return "\"urn:schemas:httpmail:textdescription\"";
                default:        return null;
            }
        }

        private static string EscapeDaslValue(string value)
        {
            // Escape single quotes in DASL
            return value?.Replace("'", "''") ?? "";
        }

        private static string BuildFallbackDaslFilter(OutlookCommandMailSearchSliceRequest req)
        {
            if (!string.IsNullOrWhiteSpace(req?.Keyword))
                return $"@SQL=\"urn:schemas:httpmail:subject\" like '%{EscapeDaslValue(req.Keyword)}%'";

            return "@SQL=\"urn:schemas:httpmail:subject\" like '%'";
        }

        private static string FormatDaslDateTime(DateTime value)
        {
            return OutlookDateFilter.FormatDaslDateTime(value);
        }

        private static string BuildAdvancedSearchScope(string folderPath)
        {
            // Scope syntax requires a quoted folder path. Any single quote inside path must be escaped.
            var safePath = folderPath?.Replace("'", "''") ?? "";
            return $"'{safePath}'";
        }

        /// <summary>
        /// Locates a single folder within the specified store.
        /// </summary>
        private Outlook.MAPIFolder GetFolderByPathInStore(string storeId, string folderPath)
        {
            try
            {
                var stores = this.Application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    try
                    {
                        string sid = "";
                        try { sid = store.StoreID ?? ""; } catch { }
                        if (!string.Equals(sid, storeId, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var root = store.GetRootFolder();
                        var found = NavigateToFolder(root, folderPath);
                        if (found != null)
                        {
                            try { Marshal.ReleaseComObject(stores); } catch { }
                            return found;
                        }
                    }
                    catch { }
                    finally { try { Marshal.ReleaseComObject(store); } catch { } }
                }
                try { Marshal.ReleaseComObject(stores); } catch { }
            }
            catch { }
            return null;
        }
    }
}
