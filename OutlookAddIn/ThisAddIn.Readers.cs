using System;
using OutlookAddIn.Contracts;

namespace OutlookAddIn
{
    /// <summary>
    /// This file has been split into multiple focused files for better maintainability:
    /// - ThisAddIn.Folders.cs: Folder reading operations (ReadFolders*, CollectFolders*)
    /// - ThisAddIn.Rules.cs: Rules reading (ReadRules, DescribeComObject)
    /// - ThisAddIn.Categories.cs: Categories reading (ReadCategories)
    /// - ThisAddIn.Calendar.cs: Calendar events reading (ReadCalendarEvents)
    /// - ThisAddIn.Mails.cs: Mail reading operations (ReadMails, GetFolderByPath, NavigateToFolder)
    /// 
    /// All functionality has been preserved in the split files.
    /// This file can be safely deleted after verifying the build succeeds.
    /// </summary>
    public partial class ThisAddIn
    {
        // All reader methods have been moved to specialized files.
        // See file header comment for details.
    }
}
