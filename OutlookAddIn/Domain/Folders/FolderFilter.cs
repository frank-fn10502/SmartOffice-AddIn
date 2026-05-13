using System;
using System.Collections.Generic;
using SmartOffice.Hub.Contracts;

namespace OutlookAddIn.Domain.Folders
{
    /// <summary>
    /// Pure logic for determining which Outlook folders to include or skip.
    /// Extracted from ThisAddIn so it can be unit-tested without Outlook COM dependencies.
    /// </summary>
    public static class FolderFilter
    {
        /// <summary>
        /// Maximum folder recursion depth to prevent traversing extremely deep system folder trees.
        /// </summary>
        public const int MaxFolderDepth = 10;

        /// <summary>
        /// Outlook system/hidden folder names that are not user-created.
        /// These are internal folders used by Exchange, Outlook sync, social connectors, etc.
        /// Reference: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        /// </summary>
        public static readonly HashSet<string> SystemFolderNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Sync & conflict folders
            "Sync Issues",
            "Conflicts",
            "Local Failures",
            "Server Failures",
            // RSS
            "RSS Feeds",
            "RSS Subscriptions",
            // Quick Steps / Conversation
            "Quick Step Settings",
            "Conversation Action Settings",
            "Conversation History",
            // Social / People
            "Social Activity Notifications",
            "ExternalContacts",
            "MyContactsExtended",
            "Recipient Cache",
            "PersonMetadata",
            "{A9E2BC46-B3A0-4243-B315-60D991004455}",
            "{06967759-274D-40B2-A3EB-D7F9E73727D7}",
            // Yammer / Teams
            "Yammer Root",
            // Files / Graph
            "Files",
            "GraphFilesAndWorkPagesFolder",
            // Finder (search folders container)
            "Finder",
            // Common Views / Reminders
            "Common Views",
            "Reminders",
            "Shortcuts",
            // Spooler Queue
            "Spooler Queue",
            // Public folders (Chinese/English variants)
            "���θ�Ƨ�",
            "���Τ��",
            "Public Folders",
            "?ffentliche Ordner",
            "Dossiers publics",
            // Additional system folders
            "Recoverable Items",
            "Deletions",
            "Purges",
            "Versions",
            "DiscoveryHolds",
            "Calendar Logging",
            "Audits",
            "AdminAuditLogs",
            "FreeBusy Data",
            "Top of Information Store",
            "System",
            "ExchangeSyncData",
            "AllItems",
            "AllContacts",
            "Freebusy Data",
            "Schedule",
            "GAL Contacts",
            "OAB Version 2",
            "OAB Version 3",
            "OAB Version 4",
            "Offline Address Book",
        };

        /// <summary>
        /// Determines whether a folder name should be skipped (system/hidden folder).
        /// </summary>
        public static bool IsSystemFolder(string folderName)
        {
            if (string.IsNullOrEmpty(folderName))
                return false;

            if (SystemFolderNames.Contains(folderName))
                return true;

            // Skip folders whose name looks like a GUID (internal Exchange folders)
            if (folderName.Length > 30 && folderName.StartsWith("{") && folderName.EndsWith("}"))
                return true;

            return false;
        }

        /// <summary>
        /// Checks whether the given depth exceeds the maximum allowed recursion depth.
        /// </summary>
        public static bool ExceedsMaxDepth(int depth)
        {
            return depth > MaxFolderDepth;
        }

        /// <summary>
        /// Builds a flat folder list from simulated folder descriptors.
        /// </summary>
        public static List<FolderDto> BuildTree(List<TestFolderNode> roots)
        {
            var result = new List<FolderDto>();
            foreach (var root in roots)
            {
                CollectNode(root, result, 0, "");
            }
            return result;
        }

        private static void CollectNode(TestFolderNode node, List<FolderDto> list, int depth, string parentFolderPath)
        {
            if (ExceedsMaxDepth(depth))
                return;

            var dto = new FolderDto
            {
                Name = node.Name,
                FolderPath = node.FolderPath,
                ParentFolderPath = parentFolderPath,
                ItemCount = node.ItemCount,
                StoreId = "",
                IsStoreRoot = depth == 0
            };

            list.Add(dto);

            if (node.Children != null)
            {
                foreach (var child in node.Children)
                {
                    if (!IsSystemFolder(child.Name))
                    {
                        CollectNode(child, list, depth + 1, node.FolderPath);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Represents a simulated Outlook folder for testing purposes.
    /// Mirrors the shape of Outlook.MAPIFolder without COM dependency.
    /// </summary>
    public class TestFolderNode
    {
        public string Name { get; set; }
        public string FolderPath { get; set; }
        public int ItemCount { get; set; }
        public List<TestFolderNode> Children { get; set; }

        public TestFolderNode()
        {
            Name = "";
            FolderPath = "";
            Children = new List<TestFolderNode>();
        }
    }
}
