using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using OutlookAddIn.OutlookServices.Categories;
using OutlookAddIn.OutlookServices.Common;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        internal async Task HandleUpdateMailPropertiesAsync(OutlookCommand cmd)
        {
            var req = cmd.MailPropertiesRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "update_mail_properties failed: missing mail id");
                return;
            }

            var mail = FindMailByEntryId(req.MailId);
            if (mail == null)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "update_mail_properties failed: mail not found");
                return;
            }

            bool categoriesChanged = false;
            try
            {
                if (req.IsRead.HasValue)
                    mail.UnRead = !req.IsRead.Value;

                if (req.Categories != null)
                    mail.Categories = string.Join(", ", req.Categories);

                // Task flag handling
                if (!string.IsNullOrEmpty(req.FlagInterval) && req.FlagInterval != "none")
                {
                    if (req.FlagInterval == "complete")
                    {
                    mail.FlagRequest = string.IsNullOrEmpty(req.FlagRequest) ? "??" : req.FlagRequest;
                        mail.FlagStatus = Outlook.OlFlagStatus.olFlagComplete;
                        if (req.TaskCompletedDate.HasValue)
                            try { mail.TaskCompletedDate = OutlookDateFilter.ToOutlookLocalDateTime(req.TaskCompletedDate.Value); } catch { }
                        else
                            try { mail.TaskCompletedDate = DateTime.Today; } catch { }
                    }
                    else
                    {
                        mail.FlagRequest = string.IsNullOrEmpty(req.FlagRequest) ? "Follow up" : req.FlagRequest;
                        mail.FlagStatus = Outlook.OlFlagStatus.olFlagMarked;

                        DateTime? autoStart = null;
                        DateTime? autoDue = null;
                        switch (req.FlagInterval)
                        {
                            case "today":
                                autoStart = DateTime.Today;
                                autoDue = DateTime.Today;
                                break;
                            case "tomorrow":
                                autoStart = DateTime.Today.AddDays(1);
                                autoDue = DateTime.Today.AddDays(1);
                                break;
                            case "this_week":
                                autoStart = DateTime.Today;
                                // End of week = next Sunday
                                int daysUntilSunday = ((int)DayOfWeek.Sunday - (int)DateTime.Today.DayOfWeek + 7) % 7;
                                autoDue = DateTime.Today.AddDays(daysUntilSunday == 0 ? 0 : daysUntilSunday);
                                break;
                            case "next_week":
                                int daysUntilNextMonday = ((int)DayOfWeek.Monday - (int)DateTime.Today.DayOfWeek + 7) % 7;
                                if (daysUntilNextMonday == 0) daysUntilNextMonday = 7;
                                autoStart = DateTime.Today.AddDays(daysUntilNextMonday);
                                autoDue = autoStart.Value.AddDays(6);
                                break;
                            case "no_date":
                                // Mark as task without dates
                                break;
                        }

                        DateTime startDate = OutlookDateFilter.ToOutlookLocalDateTime(req.TaskStartDate ?? autoStart ?? DateTime.Today);
                        DateTime dueDate = OutlookDateFilter.ToOutlookLocalDateTime(req.TaskDueDate ?? autoDue ?? DateTime.Today);
                        if (req.FlagInterval != "no_date")
                        {
                            try { mail.TaskStartDate = startDate; } catch { }
                            try { mail.TaskDueDate = dueDate; } catch { }
                        }
                    }
                }
                else if (req.FlagInterval == "none")
                {
                    try { mail.ClearTaskFlag(); } catch { }
                }

                // Create new master categories if requested
                if (req.NewCategories != null && req.NewCategories.Count > 0)
                {
                    try
                    {
                        var ns = this.Application.Session;
                        var masterCategories = ns.Categories;
                        foreach (var nc in req.NewCategories)
                        {
                            if (string.IsNullOrWhiteSpace(nc.Name)) continue;
                            try
                            {
                                Outlook.OlCategoryColor color = Outlook.OlCategoryColor.olCategoryColorNone;
                                if (!string.IsNullOrEmpty(nc.Color))
                                {
                                    try { color = (Outlook.OlCategoryColor)Enum.Parse(typeof(Outlook.OlCategoryColor), nc.Color); }
                                    catch
                                    {
                                        if (nc.ColorValue > 0)
                                            color = (Outlook.OlCategoryColor)nc.ColorValue;
                                    }
                                }
                                else if (nc.ColorValue > 0)
                                {
                                    color = (Outlook.OlCategoryColor)nc.ColorValue;
                                }
                                masterCategories.Add(nc.Name, color);
                                categoriesChanged = true;
                            }
                            catch { /* category may already exist */ }
                        }
                        try { Marshal.ReleaseComObject(masterCategories); } catch { }
                    }
                    catch { }
                }

                mail.Save();
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "update_mail_properties completed.");

                // Push single mail update instead of re-reading entire folder
                var updatedDto = ReadSingleMailDto(mail, req.FolderPath);
                if (updatedDto != null)
                    await _signalRClient.PushMailAsync(updatedDto);

                // Re-push categories if new ones were added
                if (categoriesChanged)
                {
                    var cats = new OutlookCategoryReader(Application).ReadCategories();
                    await _signalRClient.PushCategoriesAsync(cats);
                }
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "update_mail_properties failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                try { Marshal.ReleaseComObject(mail); } catch { }
            }
        }

        internal async Task HandleMoveMailAsync(OutlookCommand cmd)
        {
            var req = cmd.MoveMailRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId) || string.IsNullOrEmpty(req.DestinationFolderPath))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "move_mail failed: missing mail id or destination");
                return;
            }

            Outlook.MailItem mail = null;
            Outlook.MAPIFolder dest = null;
            try
            {
                mail = FindMailByEntryId(req.MailId);
                if (mail == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "move_mail failed: mail not found");
                    return;
                }

                // Resolve destination: always use olFolderDeletedItems for "deleted items" to be locale-independent.
                dest = ResolveDestinationFolder(req.DestinationFolderPath);
                if (dest == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "move_mail failed: destination folder not found");
                    return;
                }

                // Always move ¡X never call MailItem.Delete()
                var moved = mail.Move(dest);
                try { if (moved != null) Marshal.ReleaseComObject(moved); } catch { }

                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "move_mail completed.");

                // Re-push source folder mails (mail removed)
                if (!string.IsNullOrEmpty(req.SourceFolderPath))
                {
                    var sourceMails = ReadMails(new FetchMailsRequest { FolderPath = req.SourceFolderPath, Range = "30d", MaxCount = 100 });
                    await _signalRClient.PushMailsAsync(sourceMails);
                }

                // Update folder counts (source + destination)
                await PushFolderSyncAsync(req.SourceFolderPath, req.DestinationFolderPath);
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "move_mail failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                if (dest != null) try { Marshal.ReleaseComObject(dest); } catch { }
                if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { }
            }
        }

        /// <summary>
        /// Resolves a destination folder path to a MAPIFolder.
        /// If the path matches the Deleted Items folder (by olFolderDeletedItems ¡X locale-independent),
        /// that folder is returned directly without relying on a locale-specific folder name string.
        /// </summary>
        private Outlook.MAPIFolder ResolveDestinationFolder(string destinationFolderPath)
        {
            // First try via olFolderDeletedItems so we are locale-independent
            try
            {
                var deletedItems = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                if (deletedItems != null && deletedItems.FolderPath == destinationFolderPath)
                    return deletedItems;
                if (deletedItems != null) try { Marshal.ReleaseComObject(deletedItems); } catch { }
            }
            catch { }

            // Fallback: walk the store tree by path
            return GetFolderByPath(destinationFolderPath);
        }
        internal async Task HandleCreateFolderAsync(OutlookCommand cmd)
        {
            var req = cmd.CreateFolderRequest;
            if (req == null || string.IsNullOrWhiteSpace(req.Name) || string.IsNullOrWhiteSpace(req.ParentFolderPath))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "create_folder failed: missing name or parent path");
                return;
            }

            if (Regex.IsMatch(req.Name, "[\\\\/:*?\"<>|]"))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "create_folder failed: name contains invalid characters");
                return;
            }

            Outlook.MAPIFolder parent = null;
            try
            {
                parent = GetFolderByPath(req.ParentFolderPath);
                if (parent == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "create_folder failed: parent folder not found");
                    return;
                }

                var newFolder = parent.Folders.Add(req.Name, Outlook.OlDefaultFolders.olFolderInbox);
                string newFolderPath = null;
                try { newFolderPath = newFolder.FolderPath; } catch { }
                try { Marshal.ReleaseComObject(newFolder); } catch { }

                // Push incremental sync for parent + new folder
                await PushFolderSyncAsync(req.ParentFolderPath, newFolderPath);
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "create_folder completed.");
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "create_folder failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { }
            }
        }

        internal async Task HandleDeleteFolderAsync(OutlookCommand cmd)
        {
            var req = cmd.DeleteFolderRequest;
            if (req == null || string.IsNullOrWhiteSpace(req.FolderPath))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_folder failed: missing folder path");
                return;
            }

            Outlook.MAPIFolder folder = null;
            Outlook.MAPIFolder deletedItemsDest = null;
            Outlook.Store store = null;
            try
            {
                folder = GetFolderByPath(req.FolderPath);
                if (folder == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_folder failed: folder not found");
                    return;
                }

                // Do not allow deleting store root, hidden, or system folders.
                bool isStoreRoot = false;
                try
                {
                    var parent = folder.Parent as Outlook.MAPIFolder;
                    isStoreRoot = parent == null;
                    if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { }
                }
                catch { isStoreRoot = true; }

                if (isStoreRoot)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_folder failed: cannot delete store root folder");
                    return;
                }

                // Resolve the store-specific Deleted Items folder (locale-independent).
                try { store = folder.Store; } catch { }
                try
                {
                    if (store != null)
                        deletedItemsDest = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                }
                catch { }
                if (deletedItemsDest == null)
                {
                    try { deletedItemsDest = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems); } catch { }
                }
                if (deletedItemsDest == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_folder failed: cannot resolve Deleted Items folder");
                    return;
                }

                // Guard: if the folder is already under Deleted Items, request manual permanent delete.
                string deletedItemsPath = deletedItemsDest.FolderPath ?? "";
                string folderPath = folder.FolderPath ?? "";
                if (!string.IsNullOrEmpty(deletedItemsPath) && !string.IsNullOrEmpty(folderPath) &&
                    (folderPath == deletedItemsPath || folderPath.StartsWith(deletedItemsPath + "\\")))
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                        "manual_delete_required: folder is already in Deleted Items; please permanently delete from Outlook");
                    return;
                }

                // Soft delete: move the folder to Deleted Items. Never call folder.Delete().
                folder.MoveTo(deletedItemsDest);

                // Push incremental sync for parent of deleted folder + Deleted Items folder
                string parentPath = null;
                if (req.FolderPath.Contains("\\"))
                    parentPath = req.FolderPath.Substring(0, req.FolderPath.LastIndexOf('\\'));
                await PushFolderSyncAsync(parentPath, deletedItemsPath);
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "delete_folder completed (moved to Deleted Items).");
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_folder failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                if (store != null) try { Marshal.ReleaseComObject(store); } catch { }
                if (deletedItemsDest != null) try { Marshal.ReleaseComObject(deletedItemsDest); } catch { }
                if (folder != null) try { Marshal.ReleaseComObject(folder); } catch { }
            }
        }

        internal async Task HandleUpsertCategoryAsync(OutlookCommand cmd)
        {
            var req = cmd.CategoryRequest;
            if (req == null || string.IsNullOrWhiteSpace(req.Name))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "upsert_category failed: missing name");
                return;
            }

            try
            {
                var ns = this.Application.Session;
                var masterCategories = ns.Categories;

                // Try to find existing category
                Outlook.Category existing = null;
                try
                {
                    existing = masterCategories[req.Name];
                }
                catch { existing = null; }

                Outlook.OlCategoryColor color = Outlook.OlCategoryColor.olCategoryColorNone;
                if (!string.IsNullOrEmpty(req.Color))
                {
                    try { color = (Outlook.OlCategoryColor)Enum.Parse(typeof(Outlook.OlCategoryColor), req.Color); }
                    catch
                    {
                        // Fallback to numeric ColorValue
                        if (req.ColorValue > 0)
                            color = (Outlook.OlCategoryColor)req.ColorValue;
                    }
                }
                else if (req.ColorValue > 0)
                {
                    color = (Outlook.OlCategoryColor)req.ColorValue;
                }

                if (existing != null)
                {
                    // Update existing
                    existing.Color = color;
                    if (!string.IsNullOrEmpty(req.ShortcutKey))
                    {
                        try
                        {
                            existing.ShortcutKey = (Outlook.OlCategoryShortcutKey)Enum.Parse(typeof(Outlook.OlCategoryShortcutKey), req.ShortcutKey);
                        }
                        catch { }
                    }
                    try { Marshal.ReleaseComObject(existing); } catch { }
                }
                else
                {
                    // Add new
                    masterCategories.Add(req.Name, color);
                }

                try { Marshal.ReleaseComObject(masterCategories); } catch { }

                // After upsert, push full category list
                var categories = new OutlookCategoryReader(Application).ReadCategories();
                await _signalRClient.PushCategoriesAsync(categories);
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "upsert_category completed.");
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "upsert_category failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
        }

        private Outlook.MailItem FindMailByEntryId(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return null;
            try
            {
                var item = this.Application.Session.GetItemFromID(entryId);
                var mail = item as Outlook.MailItem;
                if (mail != null) return mail;
                try { if (item != null) Marshal.ReleaseComObject(item); } catch { }
            }
            catch { }
            return null;
        }

        /// <summary>
        /// Moves a single mail to the Deleted Items folder (locale-independent).
        /// MailItem.Delete() must never be called.
        /// </summary>
        internal async Task HandleDeleteMailAsync(OutlookCommand cmd)
        {
            var req = cmd.DeleteMailRequest;
            if (req == null || string.IsNullOrEmpty(req.MailId))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_mail failed: missing mail id");
                return;
            }

            Outlook.MailItem mail = null;
            Outlook.MAPIFolder deletedItems = null;
            Outlook.MAPIFolder parentFolder = null;
            Outlook.Store store = null;
            try
            {
                mail = FindMailByEntryId(req.MailId);
                if (mail == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_mail failed: mail not found");
                    return;
                }

                // Use the mail's own store Deleted Items folder (not the default store)
                // so that mails in PST / secondary stores go to the correct Deleted Items.
                try
                {
                    parentFolder = (Outlook.MAPIFolder)mail.Parent;
                    store = parentFolder.Store;
                    deletedItems = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                }
                catch
                {
                    // Fallback to session default if store-level lookup fails
                    deletedItems = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDeletedItems);
                }

                if (deletedItems == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_mail failed: cannot resolve Deleted Items folder");
                    return;
                }

                // Guard: if mail is already in Deleted Items or a sub-folder of it,
                // do not move further – ask the user to permanently delete from Outlook.
                string _deletedItemsRootPath = "";
                try { _deletedItemsRootPath = deletedItems.FolderPath ?? ""; } catch { }
                string _mailCurrentFolderPath = "";
                try
                {
                    if (parentFolder != null)
                        _mailCurrentFolderPath = parentFolder.FolderPath ?? "";
                    else
                    {
                        var _parentTmp = mail.Parent as Outlook.MAPIFolder;
                        try { _mailCurrentFolderPath = _parentTmp?.FolderPath ?? ""; }
                        finally { if (_parentTmp != null) try { Marshal.ReleaseComObject(_parentTmp); } catch { } }
                    }
                }
                catch { }
                if (!string.IsNullOrEmpty(_deletedItemsRootPath) && !string.IsNullOrEmpty(_mailCurrentFolderPath) &&
                    (_mailCurrentFolderPath == _deletedItemsRootPath ||
                     _mailCurrentFolderPath.StartsWith(_deletedItemsRootPath + "\\")))
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                        "manual_delete_required: mail is already in Deleted Items; please permanently delete from Outlook");
                    return;
                }

                // Move to Deleted Items – never Delete()
                var moved = mail.Move(deletedItems);
                try { if (moved != null) Marshal.ReleaseComObject(moved); } catch { }

                await _signalRClient.ReportCommandResultAsync(cmd.Id, true, "delete_mail completed.");

                // Re-push source folder mail list
                if (!string.IsNullOrEmpty(req.FolderPath))
                {
                    var sourceMails = ReadMails(new FetchMailsRequest { FolderPath = req.FolderPath, Range = "30d", MaxCount = 100 });
                    await _signalRClient.PushMailsAsync(sourceMails);
                }

                // Update folder counts (source + deleted items)
                string deletedItemsPath = null;
                try { deletedItemsPath = deletedItems?.FolderPath; } catch { }
                await PushFolderSyncAsync(req.FolderPath, deletedItemsPath);
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false, "delete_mail failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                if (store != null) try { Marshal.ReleaseComObject(store); } catch { }
                if (parentFolder != null) try { Marshal.ReleaseComObject(parentFolder); } catch { }
                if (deletedItems != null) try { Marshal.ReleaseComObject(deletedItems); } catch { }
                if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { }
            }
        }

        /// <summary>
        /// Moves multiple mails to the destination folder.
        /// Hub dispatches at most 500 mailIds per call; callers must batch larger sets.
        /// MailItem.Delete() must never be called.
        /// </summary>
        internal async Task HandleMoveMailsAsync(OutlookCommand cmd)
        {
            var req = cmd.MoveMailsRequest;
            if (req == null || req.MailIds == null || req.MailIds.Count == 0 ||
                string.IsNullOrEmpty(req.DestinationFolderPath))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "move_mails failed: missing mailIds or destination");
                return;
            }

            Outlook.MAPIFolder dest = null;
            try
            {
                dest = ResolveDestinationFolder(req.DestinationFolderPath);
                if (dest == null)
                {
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                        "move_mails failed: destination folder not found");
                    return;
                }

                int successCount = 0;
                int failCount = 0;

                foreach (var mailId in req.MailIds)
                {
                    Outlook.MailItem mail = null;
                    try
                    {
                        mail = FindMailByEntryId(mailId);
                        if (mail == null)
                        {
                            failCount++;
                            if (!req.ContinueOnError)
                            {
                                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                                    $"move_mails failed: mail not found (succeeded: {successCount}, failed: {failCount})");
                                return;
                            }
                            continue;
                        }

                        // Always move – never call Delete()
                        var moved = mail.Move(dest);
                        try { if (moved != null) Marshal.ReleaseComObject(moved); } catch { }
                        successCount++;
                    }
                    catch
                    {
                        failCount++;
                        if (!req.ContinueOnError)
                        {
                            await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                                $"move_mails failed during move (succeeded: {successCount}, failed: {failCount})");
                            return;
                        }
                    }
                    finally
                    {
                        if (mail != null) try { Marshal.ReleaseComObject(mail); } catch { }
                    }
                }

                string summary = $"succeeded: {successCount}, failed: {failCount}";
                await _signalRClient.ReportCommandResultAsync(cmd.Id, failCount == 0, $"move_mails completed. {summary}");

                // Re-push source folder(s)
                var sourcePaths = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(req.SourceFolderPath))
                    sourcePaths.Add(req.SourceFolderPath);
                if (req.SourceFolderPaths != null)
                    foreach (var p in req.SourceFolderPaths)
                        if (!string.IsNullOrEmpty(p)) sourcePaths.Add(p);

                foreach (var sourcePath in sourcePaths)
                {
                    var sourceMails = ReadMails(new FetchMailsRequest
                        { FolderPath = sourcePath, Range = "30d", MaxCount = 100 });
                    await _signalRClient.PushMailsAsync(sourceMails);
                }

                // Update folder counts: each source folder paired with destination
                bool destPushed = false;
                foreach (var sourcePath in sourcePaths)
                {
                    await PushFolderSyncAsync(sourcePath, destPushed ? null : req.DestinationFolderPath);
                    destPushed = true;
                }
                if (!destPushed)
                    await PushFolderSyncAsync(req.DestinationFolderPath);
            }
            catch (Exception ex)
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "move_mails failed: " + OutlookAddIn.Infrastructure.Diagnostics.SensitiveLogSanitizer.Sanitize(ex));
            }
            finally
            {
                if (dest != null) try { Marshal.ReleaseComObject(dest); } catch { }
            }
        }
    }
}
