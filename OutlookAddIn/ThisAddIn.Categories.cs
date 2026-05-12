using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn
{
    public partial class ThisAddIn
    {
        /// <summary>Read Outlook master category list from default store.</summary>
        public List<OutlookCategoryDto> ReadCategories()
        {
            var list = new List<OutlookCategoryDto>();
            try
            {
                var ns = this.Application.Session;
                var categories = ns.Categories;
                if (categories != null)
                {
                    foreach (Outlook.Category cat in categories)
                    {
                        try
                        {
                            list.Add(new OutlookCategoryDto
                            {
                                Name = cat.Name ?? "",
                                Color = cat.Color.ToString(),
                                ColorValue = (int)cat.Color,
                                ShortcutKey = cat.ShortcutKey.ToString()
                            });
                        }
                        catch { }
                        finally
                        {
                            try { Marshal.ReleaseComObject(cat); } catch { }
                        }
                    }
                    try { Marshal.ReleaseComObject(categories); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("ReadCategories error: " + ex);
            }
            return list;
        }
    }
}
