using System;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.OutlookServices.Folders
{
    internal sealed class OutlookFolderLocator
    {
        private readonly Outlook.Application _application;

        public OutlookFolderLocator(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public Outlook.MAPIFolder GetFolderByPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return null;

            Outlook.Stores stores = null;
            try
            {
                stores = _application.Session.Stores;
                foreach (Outlook.Store store in stores)
                {
                    Outlook.MAPIFolder root = null;
                    try
                    {
                        root = store.GetRootFolder();
                        var found = NavigateToFolder(root, path);
                        root = null;
                        if (found != null)
                            return found;
                    }
                    catch { }
                    finally
                    {
                        Release(root);
                        Release(store);
                    }
                }
            }
            catch { }
            finally
            {
                Release(stores);
            }

            return null;
        }

        private static Outlook.MAPIFolder NavigateToFolder(Outlook.MAPIFolder current, string targetPath)
        {
            if (current == null || string.IsNullOrEmpty(targetPath))
                return null;

            try
            {
                if (string.Equals(current.FolderPath, targetPath, StringComparison.OrdinalIgnoreCase))
                    return current;
            }
            catch { }

            Outlook.Folders subFolders = null;
            try
            {
                subFolders = current.Folders;
                foreach (Outlook.MAPIFolder sub in subFolders)
                {
                    Outlook.MAPIFolder found = null;
                    try
                    {
                        var subPath = sub.FolderPath ?? "";
                        if (targetPath.StartsWith(subPath, StringComparison.OrdinalIgnoreCase))
                            found = NavigateToFolder(sub, targetPath);

                        if (found != null)
                            return found;
                    }
                    catch { }
                    finally
                    {
                        if (found == null)
                            Release(sub);
                    }
                }
            }
            finally
            {
                Release(subFolders);
                Release(current);
            }

            return null;
        }

        private static void Release(object obj)
        {
            if (obj == null)
                return;

            try
            {
                if (Marshal.IsComObject(obj))
                    Marshal.ReleaseComObject(obj);
            }
            catch { }
        }
    }
}
