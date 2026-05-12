using System.Collections.Generic;

namespace OutlookAddIn
{
    /// <summary>
    /// Minimal FolderDto stub for test project.
    /// Mirrors the definition in HubClient.cs without pulling in Framework-only dependencies.
    /// </summary>
    public class FolderDto
    {
        public string Name { get; set; } = "";
        public string EntryId { get; set; } = "";
        public string FolderPath { get; set; } = "";
        public string ParentEntryId { get; set; } = "";
        public string ParentFolderPath { get; set; } = "";
        public int ItemCount { get; set; }
        public string StoreId { get; set; } = "";
        public bool IsStoreRoot { get; set; }
        public string FolderType { get; set; } = "Unknown";
        public int DefaultItemType { get; set; } = -1;
        public bool IsHidden { get; set; }
        public bool IsSystem { get; set; }
        public bool HasChildren { get; set; }
        public bool ChildrenLoaded { get; set; }
        public string DiscoveryState { get; set; } = "partial";
    }
}
