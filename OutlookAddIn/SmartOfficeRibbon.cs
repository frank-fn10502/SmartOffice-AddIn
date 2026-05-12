using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn
{
    [ComVisible(true)]
    public class SmartOfficeRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("OutlookAddIn.SmartOfficeRibbon.xml"))
            {
                if (stream == null)
                {
                    // Fallback: read from file next to assembly
                    var dir = Path.GetDirectoryName(asm.Location);
                    return File.ReadAllText(Path.Combine(dir, "SmartOfficeRibbon.xml"));
                }
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        public void OnToggleChatPane(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.ToggleChatPane();
        }
    }
}
