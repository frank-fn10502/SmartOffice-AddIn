using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SmartOffice.Hub.Contracts;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAddIn.Outlook.Categories
{
    internal sealed class OutlookCategoryReader
    {
        private readonly Outlook.Application _application;

        public OutlookCategoryReader(Outlook.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public List<OutlookCategoryDto> ReadCategories()
        {
            var list = new List<OutlookCategoryDto>();
            try
            {
                var session = _application.Session;
                var categories = session.Categories;
                if (categories != null)
                {
                    foreach (Outlook.Category category in categories)
                    {
                        try
                        {
                            list.Add(new OutlookCategoryDto
                            {
                                Name = category.Name ?? "",
                                Color = category.Color.ToString(),
                                ColorValue = (int)category.Color,
                                ShortcutKey = category.ShortcutKey.ToString()
                            });
                        }
                        catch { }
                        finally
                        {
                            try { Marshal.ReleaseComObject(category); } catch { }
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
