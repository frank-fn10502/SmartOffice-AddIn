using SmartOffice.Hub.Contracts;
using OutlookAddIn.Domain.Folders;

namespace OutlookAddIn.Tests
{
    public class FolderFilterTests
    {
        // ====================================================================
        // IsSystemFolder - known system folder names
        // ====================================================================

        [Theory]
        [InlineData("Sync Issues")]
        [InlineData("sync issues")]          // case-insensitive
        [InlineData("SYNC ISSUES")]
        [InlineData("Conflicts")]
        [InlineData("Local Failures")]
        [InlineData("Server Failures")]
        [InlineData("RSS Feeds")]
        [InlineData("RSS Subscriptions")]
        [InlineData("Quick Step Settings")]
        [InlineData("Conversation Action Settings")]
        [InlineData("Conversation History")]
        [InlineData("Social Activity Notifications")]
        [InlineData("ExternalContacts")]
        [InlineData("MyContactsExtended")]
        [InlineData("Recipient Cache")]
        [InlineData("PersonMetadata")]
        [InlineData("{A9E2BC46-B3A0-4243-B315-60D991004455}")]
        [InlineData("{06967759-274D-40B2-A3EB-D7F9E73727D7}")]
        [InlineData("Yammer Root")]
        [InlineData("Files")]
        [InlineData("GraphFilesAndWorkPagesFolder")]
        [InlineData("Finder")]
        [InlineData("Common Views")]
        [InlineData("Reminders")]
        [InlineData("Shortcuts")]
        [InlineData("Spooler Queue")]
        public void IsSystemFolder_KnownSystemNames_ReturnsTrue(string name)
        {
            Assert.True(FolderFilter.IsSystemFolder(name));
        }

        // ====================================================================
        // IsSystemFolder - GUID-like names (Exchange internal folders)
        // ====================================================================

        [Theory]
        [InlineData("{12345678-1234-1234-1234-123456789ABC}")]
        [InlineData("{ABCDEFAB-CDEF-ABCD-EFAB-CDEFABCDEFAB}")]
        public void IsSystemFolder_GuidLikeNames_ReturnsTrue(string name)
        {
            Assert.True(FolderFilter.IsSystemFolder(name));
        }

        // ====================================================================
        // IsSystemFolder - normal user folder names
        // ====================================================================

        [Theory]
        [InlineData("Inbox")]
        [InlineData("Sent Items")]
        [InlineData("Drafts")]
        [InlineData("Deleted Items")]
        [InlineData("Junk Email")]
        [InlineData("Archive")]
        [InlineData("Outbox")]
        [InlineData("Calendar")]
        [InlineData("Contacts")]
        [InlineData("Tasks")]
        [InlineData("Notes")]
        [InlineData("Journal")]
        [InlineData("My Project")]
        [InlineData("客戶資料")]
        [InlineData("2024 Reports")]
        public void IsSystemFolder_NormalFolderNames_ReturnsFalse(string name)
        {
            Assert.False(FolderFilter.IsSystemFolder(name));
        }

        // ====================================================================
        // IsSystemFolder - edge cases
        // ====================================================================

        [Theory]
        [InlineData("")]
        [InlineData(null)]
        public void IsSystemFolder_EmptyOrNull_ReturnsFalse(string? name)
        {
            Assert.False(FolderFilter.IsSystemFolder(name!));
        }

        [Fact]
        public void IsSystemFolder_ShortBraceString_ReturnsFalse()
        {
            // Short braced string should NOT be treated as GUID
            Assert.False(FolderFilter.IsSystemFolder("{short}"));
        }

        // ====================================================================
        // ExceedsMaxDepth
        // ====================================================================

        [Theory]
        [InlineData(0, false)]
        [InlineData(5, false)]
        [InlineData(10, false)]  // MaxFolderDepth == 10, depth == 10 is still OK
        [InlineData(11, true)]
        [InlineData(100, true)]
        public void ExceedsMaxDepth_ReturnsExpected(int depth, bool expected)
        {
            Assert.Equal(expected, FolderFilter.ExceedsMaxDepth(depth));
        }

        // ====================================================================
        // BuildTree - realistic Outlook folder structure simulation
        // FolderDto is now a flat list; parent-child is expressed via ParentFolderPath.
        // ====================================================================

        // Helper: direct children of a folder in the flat list.
        private static List<FolderDto> ChildrenOf(List<FolderDto> all, FolderDto parent)
            => all.Where(f => f.ParentFolderPath == parent.FolderPath).ToList();

        /// <summary>
        /// Simulates a real Outlook mailbox with both user folders and system folders.
        /// Verifies that system folders are excluded from the output tree.
        /// </summary>
        [Fact]
        public void BuildTree_RealisticMailbox_ExcludesSystemFolders()
        {
            // Arrange: simulate what Outlook DefaultStore.GetRootFolder() returns
            var root = new TestFolderNode
            {
                Name = "Mailbox - Test User",
                FolderPath = "\\\\Mailbox - Test User",
                ItemCount = 0,
                Children = new List<TestFolderNode>
                {
                    // Normal user folders
                    new TestFolderNode
                    {
                        Name = "Inbox",
                        FolderPath = "\\\\Mailbox - Test User\\Inbox",
                        ItemCount = 25,
                        Children = new List<TestFolderNode>
                        {
                            new TestFolderNode { Name = "Projects", FolderPath = "\\\\Mailbox - Test User\\Inbox\\Projects", ItemCount = 8 },
                            new TestFolderNode { Name = "Newsletters", FolderPath = "\\\\Mailbox - Test User\\Inbox\\Newsletters", ItemCount = 12 }
                        }
                    },
                    new TestFolderNode { Name = "Sent Items", FolderPath = "\\\\Mailbox - Test User\\Sent Items", ItemCount = 100 },
                    new TestFolderNode { Name = "Drafts", FolderPath = "\\\\Mailbox - Test User\\Drafts", ItemCount = 3 },
                    new TestFolderNode { Name = "Deleted Items", FolderPath = "\\\\Mailbox - Test User\\Deleted Items", ItemCount = 50 },
                    new TestFolderNode { Name = "Junk Email", FolderPath = "\\\\Mailbox - Test User\\Junk Email", ItemCount = 7 },
                    new TestFolderNode { Name = "Outbox", FolderPath = "\\\\Mailbox - Test User\\Outbox", ItemCount = 0 },
                    new TestFolderNode { Name = "Calendar", FolderPath = "\\\\Mailbox - Test User\\Calendar", ItemCount = 200 },
                    new TestFolderNode { Name = "Contacts", FolderPath = "\\\\Mailbox - Test User\\Contacts", ItemCount = 50 },
                    new TestFolderNode { Name = "Tasks", FolderPath = "\\\\Mailbox - Test User\\Tasks", ItemCount = 5 },
                    new TestFolderNode { Name = "Notes", FolderPath = "\\\\Mailbox - Test User\\Notes", ItemCount = 2 },
                    new TestFolderNode { Name = "Archive", FolderPath = "\\\\Mailbox - Test User\\Archive", ItemCount = 500 },

                    // === System/hidden folders that SHOULD be filtered out ===
                    new TestFolderNode
                    {
                        Name = "Sync Issues",
                        FolderPath = "\\\\Mailbox - Test User\\Sync Issues",
                        ItemCount = 15,
                        Children = new List<TestFolderNode>
                        {
                            new TestFolderNode { Name = "Conflicts", FolderPath = "\\\\Mailbox - Test User\\Sync Issues\\Conflicts", ItemCount = 3 },
                            new TestFolderNode { Name = "Local Failures", FolderPath = "\\\\Mailbox - Test User\\Sync Issues\\Local Failures", ItemCount = 5 },
                            new TestFolderNode { Name = "Server Failures", FolderPath = "\\\\Mailbox - Test User\\Sync Issues\\Server Failures", ItemCount = 7 },
                        }
                    },
                    new TestFolderNode { Name = "RSS Feeds", FolderPath = "\\\\Mailbox - Test User\\RSS Feeds", ItemCount = 0 },
                    new TestFolderNode { Name = "Quick Step Settings", FolderPath = "\\\\Mailbox - Test User\\Quick Step Settings", ItemCount = 0 },
                    new TestFolderNode { Name = "Conversation Action Settings", FolderPath = "\\\\Mailbox - Test User\\Conversation Action Settings", ItemCount = 0 },
                    new TestFolderNode { Name = "Conversation History", FolderPath = "\\\\Mailbox - Test User\\Conversation History", ItemCount = 0 },
                    new TestFolderNode { Name = "PersonMetadata", FolderPath = "\\\\Mailbox - Test User\\PersonMetadata", ItemCount = 0 },
                    new TestFolderNode { Name = "ExternalContacts", FolderPath = "\\\\Mailbox - Test User\\ExternalContacts", ItemCount = 0 },
                    new TestFolderNode { Name = "MyContactsExtended", FolderPath = "\\\\Mailbox - Test User\\MyContactsExtended", ItemCount = 0 },
                    new TestFolderNode { Name = "Recipient Cache", FolderPath = "\\\\Mailbox - Test User\\Recipient Cache", ItemCount = 0 },
                    new TestFolderNode { Name = "Social Activity Notifications", FolderPath = "\\\\Mailbox - Test User\\Social Activity Notifications", ItemCount = 0 },
                    new TestFolderNode { Name = "Yammer Root", FolderPath = "\\\\Mailbox - Test User\\Yammer Root", ItemCount = 0 },
                    new TestFolderNode { Name = "Files", FolderPath = "\\\\Mailbox - Test User\\Files", ItemCount = 0 },
                    new TestFolderNode { Name = "GraphFilesAndWorkPagesFolder", FolderPath = "\\\\Mailbox - Test User\\GraphFilesAndWorkPagesFolder", ItemCount = 0 },
                    new TestFolderNode { Name = "Finder", FolderPath = "\\\\Mailbox - Test User\\Finder", ItemCount = 0 },
                    new TestFolderNode { Name = "Common Views", FolderPath = "\\\\Mailbox - Test User\\Common Views", ItemCount = 0 },
                    new TestFolderNode { Name = "Reminders", FolderPath = "\\\\Mailbox - Test User\\Reminders", ItemCount = 0 },
                    new TestFolderNode { Name = "Shortcuts", FolderPath = "\\\\Mailbox - Test User\\Shortcuts", ItemCount = 0 },
                    new TestFolderNode { Name = "Spooler Queue", FolderPath = "\\\\Mailbox - Test User\\Spooler Queue", ItemCount = 0 },
                    new TestFolderNode { Name = "{A9E2BC46-B3A0-4243-B315-60D991004455}", FolderPath = "\\\\Mailbox - Test User\\{A9E2BC46-B3A0-4243-B315-60D991004455}", ItemCount = 0 },
                    new TestFolderNode { Name = "{06967759-274D-40B2-A3EB-D7F9E73727D7}", FolderPath = "\\\\Mailbox - Test User\\{06967759-274D-40B2-A3EB-D7F9E73727D7}", ItemCount = 0 },
                }
            };

            // Act
            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });

            // Assert: should have exactly 1 root (others are children / excluded)
            var roots = result.Where(f => f.IsStoreRoot).ToList();
            Assert.Single(roots);
            var mailbox = roots[0];
            Assert.Equal("Mailbox - Test User", mailbox.Name);

            // Collect all direct-child names
            var subNames = new HashSet<string>(ChildrenOf(result, mailbox).Select(f => f.Name));

            // Normal folders should be present
            Assert.Contains("Inbox", subNames);
            Assert.Contains("Sent Items", subNames);
            Assert.Contains("Drafts", subNames);
            Assert.Contains("Deleted Items", subNames);
            Assert.Contains("Junk Email", subNames);
            Assert.Contains("Outbox", subNames);
            Assert.Contains("Calendar", subNames);
            Assert.Contains("Contacts", subNames);
            Assert.Contains("Tasks", subNames);
            Assert.Contains("Notes", subNames);
            Assert.Contains("Archive", subNames);

            // System folders should NOT be present
            Assert.DoesNotContain("Sync Issues", subNames);
            Assert.DoesNotContain("RSS Feeds", subNames);
            Assert.DoesNotContain("Quick Step Settings", subNames);
            Assert.DoesNotContain("Conversation Action Settings", subNames);
            Assert.DoesNotContain("Conversation History", subNames);
            Assert.DoesNotContain("PersonMetadata", subNames);
            Assert.DoesNotContain("ExternalContacts", subNames);
            Assert.DoesNotContain("MyContactsExtended", subNames);
            Assert.DoesNotContain("Recipient Cache", subNames);
            Assert.DoesNotContain("Social Activity Notifications", subNames);
            Assert.DoesNotContain("Yammer Root", subNames);
            Assert.DoesNotContain("Files", subNames);
            Assert.DoesNotContain("GraphFilesAndWorkPagesFolder", subNames);
            Assert.DoesNotContain("Finder", subNames);
            Assert.DoesNotContain("Common Views", subNames);
            Assert.DoesNotContain("Reminders", subNames);
            Assert.DoesNotContain("Shortcuts", subNames);
            Assert.DoesNotContain("Spooler Queue", subNames);
            Assert.DoesNotContain("{A9E2BC46-B3A0-4243-B315-60D991004455}", subNames);
            Assert.DoesNotContain("{06967759-274D-40B2-A3EB-D7F9E73727D7}", subNames);

            // Only normal folders should remain as direct children (11 user folders)
            Assert.Equal(11, ChildrenOf(result, mailbox).Count);

            // Inbox should still have its children
            var inbox = result.FirstOrDefault(f => f.Name == "Inbox" && f.ParentFolderPath == mailbox.FolderPath);
            Assert.NotNull(inbox);
            var inboxChildren = ChildrenOf(result, inbox!);
            Assert.Equal(2, inboxChildren.Count);
            Assert.Contains(inboxChildren, f => f.Name == "Projects");
            Assert.Contains(inboxChildren, f => f.Name == "Newsletters");
        }

        /// <summary>
        /// Verifies that nested system folders inside user folders are also filtered out.
        /// </summary>
        [Fact]
        public void BuildTree_SystemFolderNestedInsideUserFolder_IsExcluded()
        {
            var root = new TestFolderNode
            {
                Name = "Mailbox",
                FolderPath = "\\\\Mailbox",
                Children = new List<TestFolderNode>
                {
                    new TestFolderNode
                    {
                        Name = "Inbox",
                        FolderPath = "\\\\Mailbox\\Inbox",
                        ItemCount = 10,
                        Children = new List<TestFolderNode>
                        {
                            new TestFolderNode { Name = "My Project", FolderPath = "\\\\Mailbox\\Inbox\\My Project", ItemCount = 5 },
                            // This system folder might appear nested under Inbox in some Exchange configs
                            new TestFolderNode { Name = "PersonMetadata", FolderPath = "\\\\Mailbox\\Inbox\\PersonMetadata", ItemCount = 0 },
                        }
                    }
                }
            };

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });
            var inbox = result.First(f => f.Name == "Inbox");
            var inboxChildren = ChildrenOf(result, inbox);

            Assert.Single(inboxChildren);
            Assert.Equal("My Project", inboxChildren[0].Name);
        }

        /// <summary>
        /// Verifies the depth limit is enforced: folders beyond MaxFolderDepth are not included.
        /// </summary>
        [Fact]
        public void BuildTree_DeeplyNestedFolders_StopsAtMaxDepth()
        {
            // Build a chain: root -> level1 -> level2 -> ... -> level(MaxDepth+2)
            TestFolderNode deepest = new TestFolderNode
            {
                Name = $"Level{FolderFilter.MaxFolderDepth + 2}",
                FolderPath = $"\\\\Level{FolderFilter.MaxFolderDepth + 2}",
                ItemCount = 999
            };

            var current = deepest;
            for (int i = FolderFilter.MaxFolderDepth + 1; i >= 0; i--)
            {
                current = new TestFolderNode
                {
                    Name = $"Level{i}",
                    FolderPath = $"\\\\Level{i}",
                    ItemCount = i,
                    Children = new List<TestFolderNode> { current }
                };
            }

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { current });

            // Walk the flat list via ParentFolderPath to measure actual depth.
            int actualDepth = 0;
            var node = result[0];
            while (true)
            {
                var child = result.FirstOrDefault(f => f.ParentFolderPath == node.FolderPath);
                if (child == null) break;
                actualDepth++;
                node = child;
            }

            // The deepest node that gets included should be at depth == MaxFolderDepth
            // (depth 0 is root, so MaxFolderDepth levels of children)
            Assert.True(actualDepth <= FolderFilter.MaxFolderDepth,
                $"Tree depth {actualDepth} exceeded max {FolderFilter.MaxFolderDepth}");
        }

        /// <summary>
        /// Empty mailbox should produce a single root with no subfolders.
        /// </summary>
        [Fact]
        public void BuildTree_EmptyMailbox_ReturnsSingleRoot()
        {
            var root = new TestFolderNode
            {
                Name = "Mailbox - Empty User",
                FolderPath = "\\\\Mailbox - Empty User",
                ItemCount = 0,
                Children = new List<TestFolderNode>()
            };

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });

            Assert.Single(result);
            Assert.Equal("Mailbox - Empty User", result[0].Name);
            Assert.Empty(ChildrenOf(result, result[0]));
        }

        /// <summary>
        /// A mailbox that contains ONLY system folders should have zero subfolders in the output.
        /// </summary>
        [Fact]
        public void BuildTree_OnlySystemFolders_SubfoldersEmpty()
        {
            var root = new TestFolderNode
            {
                Name = "Mailbox",
                FolderPath = "\\\\Mailbox",
                Children = new List<TestFolderNode>
                {
                    new TestFolderNode { Name = "Sync Issues", FolderPath = "\\\\Mailbox\\Sync Issues" },
                    new TestFolderNode { Name = "PersonMetadata", FolderPath = "\\\\Mailbox\\PersonMetadata" },
                    new TestFolderNode { Name = "Finder", FolderPath = "\\\\Mailbox\\Finder" },
                    new TestFolderNode { Name = "{A9E2BC46-B3A0-4243-B315-60D991004455}", FolderPath = "\\\\Mailbox\\{A9E2BC46-B3A0-4243-B315-60D991004455}" },
                }
            };

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });

            Assert.Single(result);
            Assert.Empty(ChildrenOf(result, result[0]));
        }

        /// <summary>
        /// Tree structure should be preserved: root contains children, not siblings.
        /// This specifically tests the bug where folders were incorrectly promoted to top-level.
        /// </summary>
        [Fact]
        public void BuildTree_PreservesHierarchy_ChildrenAreNotPromotedToTopLevel()
        {
            var root = new TestFolderNode
            {
                Name = "Mailbox - User",
                FolderPath = "\\\\Mailbox - User",
                Children = new List<TestFolderNode>
                {
                    new TestFolderNode
                    {
                        Name = "Inbox",
                        FolderPath = "\\\\Mailbox - User\\Inbox",
                        ItemCount = 18,
                        Children = new List<TestFolderNode>
                        {
                            new TestFolderNode
                            {
                                Name = "Projects",
                                FolderPath = "\\\\Mailbox - User\\Inbox\\Projects",
                                ItemCount = 7,
                                Children = new List<TestFolderNode>
                                {
                                    new TestFolderNode { Name = "SmartOffice", FolderPath = "\\\\Mailbox - User\\Inbox\\Projects\\SmartOffice", ItemCount = 3 }
                                }
                            }
                        }
                    },
                    new TestFolderNode { Name = "Sent Items", FolderPath = "\\\\Mailbox - User\\Sent Items", ItemCount = 9 },
                }
            };

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });

            // Only 1 top-level entry (IsStoreRoot = true, parentFolderPath empty)
            var roots = result.Where(f => f.IsStoreRoot).ToList();
            Assert.Single(roots);

            // Inbox and Sent Items are children of Mailbox, NOT top-level
            var mailboxNode = roots[0];
            Assert.Equal(2, ChildrenOf(result, mailboxNode).Count);

            // Projects is under Inbox, NOT top-level or under Mailbox
            var inbox = result.FirstOrDefault(f => f.Name == "Inbox" && f.ParentFolderPath == mailboxNode.FolderPath);
            Assert.NotNull(inbox);
            var inboxChildren = ChildrenOf(result, inbox!);
            Assert.Single(inboxChildren);
            Assert.Equal("Projects", inboxChildren[0].Name);

            // SmartOffice is under Projects
            var projectsChildren = ChildrenOf(result, inboxChildren[0]);
            Assert.Single(projectsChildren);
            Assert.Equal("SmartOffice", projectsChildren[0].Name);
        }

        /// <summary>
        /// Verify ItemCount values are preserved correctly in the output.
        /// </summary>
        [Fact]
        public void BuildTree_PreservesItemCount()
        {
            var root = new TestFolderNode
            {
                Name = "Mailbox",
                FolderPath = "\\\\Mailbox",
                ItemCount = 0,
                Children = new List<TestFolderNode>
                {
                    new TestFolderNode { Name = "Inbox", FolderPath = "\\\\Mailbox\\Inbox", ItemCount = 42 },
                    new TestFolderNode { Name = "Archive", FolderPath = "\\\\Mailbox\\Archive", ItemCount = 1000 },
                }
            };

            var result = FolderFilter.BuildTree(new List<TestFolderNode> { root });

            Assert.Equal(0, result[0].ItemCount);
            Assert.Equal(42, result.First(f => f.Name == "Inbox").ItemCount);
            Assert.Equal(1000, result.First(f => f.Name == "Archive").ItemCount);
        }
    }
}
