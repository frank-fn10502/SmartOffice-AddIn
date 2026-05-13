using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        // ????????????????????????????????????????????????????????????????????????????????
        // MAPI property tags for PR_ATTR_HIDDEN / PR_ATTR_SYSTEM.
        // These must be read from MAPI; never inferred from folder name or display text.
        // ????????????????????????????????????????????????????????????????????????????????
        private const string MapiPropAttrHidden = "http://schemas.microsoft.com/mapi/proptag/0x10F4000B";
        private const string MapiPropAttrSystem = "http://schemas.microsoft.com/mapi/proptag/0x10F5000B";

        private static bool ReadMapiBoolean(Outlook.MAPIFolder folder, string mapiPropTag)
        {
            try
            {
                var val = folder.PropertyAccessor.GetProperty(mapiPropTag);
                if (val is bool b) return b;
                if (val is int i) return i != 0;
            }
            catch { }
            return false;
        }

        // ????????????????????????????????????????????????????????????????????????????
        // Folder type helpers
        // Builds a map of entryId ? OutlookFolderType string by calling
        // Store.GetDefaultFolder once per well-known type. Must be called on UI thread.
        // ????????????????????????????????????????????????????????????????????????????
        // Only probe folder types that are universally supported across Exchange, PST, and IMAP stores.
        // Journal, RssFeeds, SyncIssues, Conflicts, LocalFailures, ServerFailures, Notes are absent on
        // PST/IMAP stores and throw COMException on every call — excluded to avoid exception noise.
        private static readonly Outlook.OlDefaultFolders[] s_knownDefaultFolders = new[]
        {
            Outlook.OlDefaultFolders.olFolderInbox,
            Outlook.OlDefaultFolders.olFolderSentMail,
            Outlook.OlDefaultFolders.olFolderDrafts,
            Outlook.OlDefaultFolders.olFolderDeletedItems,
            Outlook.OlDefaultFolders.olFolderJunk,
            Outlook.OlDefaultFolders.olFolderCalendar,
            Outlook.OlDefaultFolders.olFolderContacts,
            Outlook.OlDefaultFolders.olFolderTasks,
            Outlook.OlDefaultFolders.olFolderOutbox,
        };

        private static readonly OutlookFolderType[] s_knownDefaultFolderTypes = new[]
        {
            OutlookFolderType.Inbox, OutlookFolderType.Sent, OutlookFolderType.Drafts, OutlookFolderType.Deleted, OutlookFolderType.Junk,
            OutlookFolderType.Calendar, OutlookFolderType.Contacts, OutlookFolderType.Tasks, OutlookFolderType.Outbox,
        };

        private Dictionary<string, OutlookFolderType> BuildStoreFolderTypeMap(Outlook.Store store)
        {
            var map = new Dictionary<string, OutlookFolderType>(StringComparer.OrdinalIgnoreCase);
            if (store == null) return map;
            for (int i = 0; i < s_knownDefaultFolders.Length; i++)
            {
                Outlook.MAPIFolder f = null;
                try
                {
                    f = store.GetDefaultFolder(s_knownDefaultFolders[i]);
                    if (f == null) continue;
                    string eid = "";
                    try { eid = f.EntryID ?? ""; } catch { }
                    if (!string.IsNullOrEmpty(eid) && !map.ContainsKey(eid))
                        map[eid] = s_knownDefaultFolderTypes[i];
                }
                catch { }
                finally { if (f != null) try { Marshal.ReleaseComObject(f); } catch { } }
            }
            return map;
        }

        private static OutlookFolderType LookupFolderType(
            Dictionary<string, OutlookFolderType> typeMap, string entryId, bool isStoreRoot, int defaultItemType)
        {
            if (isStoreRoot) return OutlookFolderType.StoreRoot;
            if (!string.IsNullOrEmpty(entryId) && typeMap != null &&
                typeMap.TryGetValue(entryId, out OutlookFolderType t)) return t;
            if (defaultItemType == 0) return OutlookFolderType.Mail;
            return OutlookFolderType.OtherSystem;
        }

        // ????????????????????????????????????????????????????????????????????????????????
        // fetch_folder_roots: list all Outlook stores + each store root folder only.
        // ????????????????????????????????????????????????????????????????????????????????
        private async Task HandleFetchFolderRootsAsync(OutlookCommand cmd)
        {
            var req = cmd.FolderDiscoveryRequest;
            string syncId = req?.SyncId ?? ("folder-sync-" + Guid.NewGuid().ToString("N").Substring(0, 8));

            try
            {
                await _signalRClient.BeginFolderSyncAsync(new FolderSyncBeginDto { SyncId = syncId });

                List<OutlookStoreDto> stores = null;
                List<FolderDto> rootFolders = null;
                _chatPane.Invoke((Action)(() => { ReadStoreRoots(out stores, out rootFolders); }));

                if (stores == null || stores.Count == 0)
                {
                    await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                        "fetch_folder_roots: no stores found or Outlook profile not ready.");
                    return;
                }

                await _signalRClient.PushFolderBatchAsync(new FolderSyncBatchDto
                {
                    SyncId = syncId,
                    Sequence = 1,
                    Reset = req?.Reset ?? true,
                    IsFinal = true,
                    Stores = stores,
                    Folders = rootFolders ?? new List<FolderDto>()
                });

                await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                    $"fetch_folder_roots completed. Stores: {stores.Count}, Folders: {rootFolders?.Count ?? 0}");
            }
            catch (Exception ex)
            {
                try
                {
                    await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                }
                catch { }
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_folder_roots error: " + SanitizeExceptionForLog(ex));
            }
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // fetch_folder_children: list direct children of a specified parent folder.
        // Prefers storeId + parentEntryId; falls back to parentFolderPath.
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        private async Task HandleFetchFolderChildrenAsync(OutlookCommand cmd)
        {
            var req = cmd.FolderDiscoveryRequest;
            string syncId = req?.SyncId ?? ("folder-sync-" + Guid.NewGuid().ToString("N").Substring(0, 8));

            if (req == null || (string.IsNullOrEmpty(req.ParentEntryId) && string.IsNullOrEmpty(req.ParentFolderPath)))
            {
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_folder_children failed: parentEntryId or parentFolderPath required");
                return;
            }

            try
            {
                await _signalRClient.BeginFolderSyncAsync(new FolderSyncBeginDto { SyncId = syncId });

                List<OutlookStoreDto> stores = null;
                List<FolderDto> childFolders = null;
                _chatPane.Invoke((Action)(() =>
                {
                    ReadFolderChildren(req, out stores, out childFolders);
                }));

                if (childFolders == null)
                {
                    await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                    await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                        "fetch_folder_children failed: parent folder not found or store not ready");
                    return;
                }

                await _signalRClient.PushFolderBatchAsync(new FolderSyncBatchDto
                {
                    SyncId = syncId,
                    Sequence = 1,
                    Reset = false,
                    IsFinal = true,
                    Stores = stores ?? new List<OutlookStoreDto>(),
                    Folders = childFolders
                });

                await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                await _signalRClient.ReportCommandResultAsync(cmd.Id, true,
                    $"fetch_folder_children completed. Folders: {childFolders.Count}");
            }
            catch (Exception ex)
            {
                try
                {
                    await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
                }
                catch { }
                await _signalRClient.ReportCommandResultAsync(cmd.Id, false,
                    "fetch_folder_children error: " + SanitizeExceptionForLog(ex));
            }
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // Reads all Outlook stores and their root folders only (no recursion).
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        private void ReadStoreRoots(out List<OutlookStoreDto> stores, out List<FolderDto> rootFolders)
        {
            stores = new List<OutlookStoreDto>();
            rootFolders = new List<FolderDto>();

            var session = this.Application.Session;
            if (session == null) return;

            var outlookStores = session.Stores;
            if (outlookStores == null) return;

            try
            {
                foreach (Outlook.Store store in outlookStores)
                {
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        string displayName = "";
                        try { displayName = store.DisplayName ?? ""; } catch { }
                        if (string.IsNullOrEmpty(displayName)) continue;

                        string storeId = "";
                        try { storeId = store.StoreID ?? ""; } catch { }

                        string storeFilePath = "";
                        try { storeFilePath = store.FilePath ?? ""; } catch { }

                        string storeKind = DetermineStoreKind(store, storeFilePath);

                        root = store.GetRootFolder();
                        if (root == null) continue;

                        string rootPath = root.FolderPath ?? "";
                        string rootEntryId = "";
                        try { rootEntryId = root.EntryID ?? ""; } catch { }

                        bool hasChildren = false;
                        try { hasChildren = root.Folders.Count > 0; } catch { }

                        int rootDefaultItemType = -1;
                        try { rootDefaultItemType = (int)root.DefaultItemType; } catch { }
                        bool rootIsHidden = ReadMapiBoolean(root, MapiPropAttrHidden);
                        bool rootIsSystem = ReadMapiBoolean(root, MapiPropAttrSystem);

                        stores.Add(new OutlookStoreDto
                        {
                            StoreId = storeId,
                            DisplayName = displayName,
                            StoreKind = storeKind,
                            StoreFilePath = storeFilePath,
                            RootFolderPath = rootPath
                        });

                        rootFolders.Add(new FolderDto
                        {
                            Name = root.Name ?? displayName,
                            EntryId = rootEntryId,
                            FolderPath = rootPath,
                            ParentEntryId = "",
                            ParentFolderPath = "",
                            ItemCount = 0,
                            StoreId = storeId,
                            IsStoreRoot = true,
                            FolderType = OutlookFolderType.StoreRoot,
                            DefaultItemType = rootDefaultItemType,
                            IsHidden = rootIsHidden,
                            IsSystem = rootIsSystem,
                            HasChildren = hasChildren,
                            ChildrenLoaded = false,
                            DiscoveryState = "partial"
                        });
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"ReadStoreRoots: store error: {ex.Message}");
                    }
                    finally
                    {
                        if (root != null) try { Marshal.ReleaseComObject(root); } catch { }
                        try { Marshal.ReleaseComObject(store); } catch { }
                    }
                }
            }
            finally
            {
                try { Marshal.ReleaseComObject(outlookStores); } catch { }
            }
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // Reads direct children of the parent folder specified in req.
        // Returns the parent itself (with childrenLoaded=true) plus its children.
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        private void ReadFolderChildren(
            OutlookCommandFolderDiscoveryRequest req,
            out List<OutlookStoreDto> stores,
            out List<FolderDto> folders)
        {
            stores = new List<OutlookStoreDto>();
            folders = null;

            int maxChildren = req.MaxChildren > 0 ? Math.Min(req.MaxChildren, 500) : 100;

            Outlook.MAPIFolder parent = null;
            string parentStoreId = req.StoreId ?? "";

            try
            {
                // Prefer storeId + parentEntryId
                if (!string.IsNullOrEmpty(req.ParentEntryId) && !string.IsNullOrEmpty(req.StoreId))
                {
                    parent = GetFolderByEntryIdInStore(req.StoreId, req.ParentEntryId);
                }

                // Fallback to parentFolderPath
                if (parent == null && !string.IsNullOrEmpty(req.ParentFolderPath))
                {
                    parent = GetFolderByPath(req.ParentFolderPath);
                    if (parent != null && string.IsNullOrEmpty(parentStoreId))
                    {
                        try
                        {
                            Outlook.Store s = parent.Store;
                            try { parentStoreId = s?.StoreID ?? ""; } finally { if (s != null) try { Marshal.ReleaseComObject(s); } catch { } }
                        }
                        catch { }
                    }
                }

                if (parent == null) return;

                folders = new List<FolderDto>();

                // Include the parent itself with childrenLoaded=true
                string parentEntryId = "";
                try { parentEntryId = parent.EntryID ?? ""; } catch { }
                string parentFolderPath = parent.FolderPath ?? "";
                string grandParentEntryId = "";
                string grandParentFolderPath = "";
                try
                {
                    var gp = parent.Parent as Outlook.MAPIFolder;
                    if (gp != null)
                    {
                        try { grandParentEntryId = gp.EntryID ?? ""; } catch { }
                        try { grandParentFolderPath = gp.FolderPath ?? ""; } catch { }
                        try { Marshal.ReleaseComObject(gp); } catch { }
                    }
                }
                catch { }

                bool parentHasChildren = false;
                try { parentHasChildren = parent.Folders.Count > 0; } catch { }

                int parentDefaultItemType = -1;
                try { parentDefaultItemType = (int)parent.DefaultItemType; } catch { }
                bool parentIsHidden = ReadMapiBoolean(parent, MapiPropAttrHidden);
                bool parentIsSystem = ReadMapiBoolean(parent, MapiPropAttrSystem);
                bool parentIsStoreRoot = string.IsNullOrEmpty(grandParentFolderPath);

                // Build folder type map once for this store
                var folderTypeMap = new Dictionary<string, OutlookFolderType>(StringComparer.OrdinalIgnoreCase);
                Outlook.Store parentStore = null;
                try
                {
                    parentStore = parent.Store;
                    folderTypeMap = BuildStoreFolderTypeMap(parentStore);
                }
                catch { }
                finally { if (parentStore != null) try { Marshal.ReleaseComObject(parentStore); } catch { } }

                folders.Add(new FolderDto
                {
                    Name = parent.Name ?? "",
                    EntryId = parentEntryId,
                    FolderPath = parentFolderPath,
                    ParentEntryId = grandParentEntryId,
                    ParentFolderPath = grandParentFolderPath,
                    ItemCount = 0,
                    StoreId = parentStoreId,
                    IsStoreRoot = parentIsStoreRoot,
                    FolderType = LookupFolderType(folderTypeMap, parentEntryId, parentIsStoreRoot, parentDefaultItemType),
                    DefaultItemType = parentDefaultItemType,
                    IsHidden = parentIsHidden,
                    IsSystem = parentIsSystem,
                    HasChildren = parentHasChildren,
                    ChildrenLoaded = true,
                    DiscoveryState = "loaded"
                });

                // Collect direct children only
                Outlook.Folders subFolders = null;
                try
                {
                    subFolders = parent.Folders;
                    int count = 0;
                    foreach (Outlook.MAPIFolder sub in subFolders)
                    {
                        if (count >= maxChildren) { try { Marshal.ReleaseComObject(sub); } catch { } break; }
                        try
                        {
                            string name = sub.Name ?? "";
                            string entryId = "";
                            try { entryId = sub.EntryID ?? ""; } catch { }
                            string folderPath = sub.FolderPath ?? "";
                            int itemCount = 0;
                            try { itemCount = sub.Items.Count; } catch { }
                            bool hasChildren = false;
                            try { hasChildren = sub.Folders.Count > 0; } catch { }

                            int subDefaultItemType = -1;
                            try { subDefaultItemType = (int)sub.DefaultItemType; } catch { }
                            bool subIsHidden = ReadMapiBoolean(sub, MapiPropAttrHidden);
                            bool subIsSystem = ReadMapiBoolean(sub, MapiPropAttrSystem);

                            folders.Add(new FolderDto
                            {
                                Name = name,
                                EntryId = entryId,
                                FolderPath = folderPath,
                                ParentEntryId = parentEntryId,
                                ParentFolderPath = parentFolderPath,
                                ItemCount = itemCount,
                                StoreId = parentStoreId,
                                IsStoreRoot = false,
                                FolderType = LookupFolderType(folderTypeMap, entryId, false, subDefaultItemType),
                                DefaultItemType = subDefaultItemType,
                                IsHidden = subIsHidden,
                                IsSystem = subIsSystem,
                                HasChildren = hasChildren,
                                ChildrenLoaded = false,
                                DiscoveryState = "partial"
                            });
                            count++;
                        }
                        catch { }
                        finally { try { Marshal.ReleaseComObject(sub); } catch { } }
                    }
                }
                finally
                {
                    if (subFolders != null) try { Marshal.ReleaseComObject(subFolders); } catch { }
                }
            }
            finally
            {
                if (parent != null) try { Marshal.ReleaseComObject(parent); } catch { }
            }
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // Utility: determine store kind from ExchangeStoreType / file extension.
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        private string DetermineStoreKind(Outlook.Store store, string storeFilePath)
        {
            try
            {
                var est = store.ExchangeStoreType;
                if (est == Outlook.OlExchangeStoreType.olPrimaryExchangeMailbox ||
                    est == Outlook.OlExchangeStoreType.olExchangeMailbox)
                    return "ost";
            }
            catch { }

            if (!string.IsNullOrEmpty(storeFilePath))
            {
                if (storeFilePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)) return "pst";
                if (storeFilePath.EndsWith(".ost", StringComparison.OrdinalIgnoreCase)) return "ost";
            }
            return "other";
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // Utility: locate a folder in a specific store by its EntryID.
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        private Outlook.MAPIFolder GetFolderByEntryIdInStore(string storeId, string entryId)
        {
            if (string.IsNullOrEmpty(storeId) || string.IsNullOrEmpty(entryId)) return null;
            try
            {
                var item = this.Application.Session.GetItemFromID(entryId, storeId);
                var folder = item as Outlook.MAPIFolder;
                if (folder != null) return folder;
                if (item != null) try { Marshal.ReleaseComObject(item); } catch { }
            }
            catch { }

            // Secondary fallback: scan stores manually
            try
            {
                var stores = this.Application.Session.Stores;
                try
                {
                    foreach (Outlook.Store store in stores)
                    {
                        try
                        {
                            string sid = "";
                            try { sid = store.StoreID ?? ""; } catch { }
                            if (!string.Equals(sid, storeId, StringComparison.OrdinalIgnoreCase)) continue;

                            var root = store.GetRootFolder();
                            var found = FindFolderByEntryId(root, entryId);
                            if (found != null) return found;
                            try { Marshal.ReleaseComObject(root); } catch { }
                        }
                        finally { try { Marshal.ReleaseComObject(store); } catch { } }
                    }
                }
                finally { try { Marshal.ReleaseComObject(stores); } catch { } }
            }
            catch { }
            return null;
        }

        private Outlook.MAPIFolder FindFolderByEntryId(Outlook.MAPIFolder current, string entryId)
        {
            string currentEntryId = "";
            try { currentEntryId = current.EntryID ?? ""; } catch { }
            if (string.Equals(currentEntryId, entryId, StringComparison.OrdinalIgnoreCase))
                return current;

            Outlook.Folders subs = null;
            try
            {
                subs = current.Folders;
                foreach (Outlook.MAPIFolder sub in subs)
                {
                    var found = FindFolderByEntryId(sub, entryId);
                    if (found != null) { try { Marshal.ReleaseComObject(subs); } catch { } return found; }
                    try { Marshal.ReleaseComObject(sub); } catch { }
                }
            }
            finally { if (subs != null) try { Marshal.ReleaseComObject(subs); } catch { } }
            return null;
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // PushFolderSyncAsync: used internally after create/delete/move to push
        // a minimal incremental sync showing updated folder item counts.
        // Only re-pushes directly affected folders (not the whole tree).
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        public async System.Threading.Tasks.Task PushFolderSyncAsync(
            string specificFolderPath = null, string specificFolderPath2 = null)
        {
            var syncId = "folder-sync-" + Guid.NewGuid().ToString("N").Substring(0, 8);
            await _signalRClient.BeginFolderSyncAsync(new FolderSyncBeginDto { SyncId = syncId });

            var stores = new List<OutlookStoreDto>();
            var folders = new List<FolderDto>();

            _chatPane.Invoke((Action)(() =>
            {
                // Push only the specified folder(s) for incremental update
                foreach (var path in new[] { specificFolderPath, specificFolderPath2 })
                {
                    if (string.IsNullOrEmpty(path)) continue;
                    Outlook.MAPIFolder f = null;
                    try
                    {
                        f = GetFolderByPath(path);
                        if (f == null) continue;

                        string storeId = "";
                        try { var s = f.Store; storeId = s?.StoreID ?? ""; try { Marshal.ReleaseComObject(s); } catch { } } catch { }
                        string entryId = "";
                        try { entryId = f.EntryID ?? ""; } catch { }
                        string parentEntryId = "";
                        string parentFolderPath = "";
                        try { var gp = f.Parent as Outlook.MAPIFolder; if (gp != null) { parentEntryId = gp.EntryID ?? ""; parentFolderPath = gp.FolderPath ?? ""; Marshal.ReleaseComObject(gp); } } catch { }
                        int itemCount = 0;
                        try { itemCount = f.Items.Count; } catch { }
                        bool hasChildren = false;
                        try { hasChildren = f.Folders.Count > 0; } catch { }
                        int syncDefaultItemType = -1;
                        try { syncDefaultItemType = (int)f.DefaultItemType; } catch { }

                        var syncTypeMap = new Dictionary<string, OutlookFolderType>(StringComparer.OrdinalIgnoreCase);
                        Outlook.Store syncStore = null;
                        try { syncStore = f.Store; syncTypeMap = BuildStoreFolderTypeMap(syncStore); }
                        catch { }
                        finally { if (syncStore != null) try { Marshal.ReleaseComObject(syncStore); } catch { } }

                        folders.Add(new FolderDto
                        {
                            Name = f.Name ?? "",
                            EntryId = entryId,
                            FolderPath = path,
                            ParentEntryId = parentEntryId,
                            ParentFolderPath = parentFolderPath,
                            ItemCount = itemCount,
                            StoreId = storeId,
                            IsStoreRoot = false,
                            FolderType = LookupFolderType(syncTypeMap, entryId, false, syncDefaultItemType),
                            DefaultItemType = syncDefaultItemType,
                            HasChildren = hasChildren,
                            ChildrenLoaded = false,
                            DiscoveryState = "partial"
                        });
                    }
                    catch { }
                    finally { if (f != null) try { Marshal.ReleaseComObject(f); } catch { } }
                }

                // If no specific folder was given, fall back to store roots only
                if (folders.Count == 0)
                    ReadStoreRoots(out stores, out var rootFolders2);
            }));

            await _signalRClient.PushFolderBatchAsync(new FolderSyncBatchDto
            {
                SyncId = syncId,
                Sequence = 1,
                Reset = false,
                IsFinal = true,
                Stores = stores,
                Folders = folders
            });

            await _signalRClient.CompleteFolderSyncAsync(new FolderSyncCompleteDto { SyncId = syncId });
        }

        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        // Legacy helper kept for search and mail path navigation.
        // ¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w¢w
        public void ReadFoldersFlat(out List<OutlookStoreDto> stores, out List<FolderDto> folders)
        {
            ReadStoreRoots(out stores, out folders);
        }
    }
}
