using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn.Ribbon
{
    [ComVisible(true)]
    public class SmartOfficeRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID)
        {
            var asm = Assembly.GetExecutingAssembly();
            using (var stream = asm.GetManifestResourceStream("OutlookAddIn.Ribbon.SmartOfficeRibbon.xml")
                ?? asm.GetManifestResourceStream("OutlookAddIn.SmartOfficeRibbon.xml"))
            {
                if (stream == null)
                {
                    // Fallback: read from file next to assembly
                    var dir = Path.GetDirectoryName(asm.Location);
                    var ribbonPath = Path.Combine(dir, "Ribbon", "SmartOfficeRibbon.xml");
                    if (!File.Exists(ribbonPath))
                        ribbonPath = Path.Combine(dir, "SmartOfficeRibbon.xml");
                    return File.ReadAllText(ribbonPath);
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
