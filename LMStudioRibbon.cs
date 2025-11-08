using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Drawing;
using Office = Microsoft.Office.Core;

namespace OutlookLMStudio
{
    [ComVisible(true)]
    public class LMStudioRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public LMStudioRibbon() { }

        #region IRibbonExtensibility Members
        public string GetCustomUI(string ribbonID)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"" onLoad=""Ribbon_Load"">
  <ribbon>
    <tabs>
      <tab idMso=""TabMail"">
        <group id=""LMStudioGroup"" label=""LMStudio Assistant"">
          <button id=""btnTogglePane""
                  label=""Afficher/Masquer""
                  size=""large""
                  onAction=""OnTogglePaneClick""
                  getImage=""GetToggleImage"" />
          <button id=""btnGenerateSelected""
                  label=""Générer Réponses""
                  size=""large""
                  onAction=""OnGenerateSelectedClick""
                  getImage=""GetGenerateImage""
                  screentip=""Générer des réponses pour les emails sélectionnés"" />
        </group>
      </tab>
    </tabs>
  </ribbon>
  <contextMenus>
    <contextMenu idMso=""ContextMenuMailItem""> 
      <button id=""btnContextGenerate""
              label=""Générer Réponse(s) avec LMStudio""
              onAction=""OnContextGenerateClick""
              imageMso=""CreateAReply""
              insertBeforeMso=""Reply"" />
    </contextMenu>
  </contextMenus>
</customUI>";
        }
        #endregion

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI) => ribbon = ribbonUI;

        public void OnTogglePaneClick(Office.IRibbonControl control)
        {
            try { Globals.ThisAddIn.ToggleTaskPane(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error); }
        }
        public void OnGenerateSelectedClick(Office.IRibbonControl control)
        {
            try { Globals.ThisAddIn.GenerateResponsesForSelection(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error); }
        }
        public void OnContextGenerateClick(Office.IRibbonControl control)
        {
            try { Globals.ThisAddIn.GenerateResponsesForSelection(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error); }
        }
        #endregion

        #region Icon Loading
        private Bitmap LoadResourceBitmap(string key, string[] embeddedFallbackNames, Func<Bitmap> drawFallback)
        {
            // 0) ResourceManager lookup
            try
            {
                var obj = OutlookLMStudio.Properties.Resources.ResourceManager.GetObject(key, OutlookLMStudio.Properties.Resources.Culture);
                if (obj is Bitmap bm) return bm;
                if (obj is Icon ic) return ic.ToBitmap();
            }
            catch { }
            // 1) Strongly typed property (if generated later)
            try
            {
                var resType = typeof(OutlookLMStudio.Properties.Resources);
                var prop = resType.GetProperty(key, BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                if (prop != null)
                {
                    var val = prop.GetValue(null, null);
                    if (val is Bitmap pb) return pb;
                    if (val is Icon pi) return pi.ToBitmap();
                }
            }
            catch { }
            // 2) Embedded resource
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var name = asm.GetManifestResourceNames().FirstOrDefault(n => embeddedFallbackNames.Any(f => n.EndsWith(f, StringComparison.OrdinalIgnoreCase)));
                if (name != null)
                {
                    using (var s = asm.GetManifestResourceStream(name))
                        if (s != null) return new Bitmap(s);
                }
            }
            catch { }
            // 3) Fallback drawing
            return drawFallback();
        }

        private Bitmap LoadGenerateIcon32() => LoadResourceBitmap(
            "ToggleIcon32",
            new[] { "ToggleIcon32.png", "ToggleIcon.png", "Toggle.png" },
            () =>
            {
                var bmp = new Bitmap(32, 32);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    using (var b = new SolidBrush(Color.ForestGreen)) g.FillRectangle(b, 6, 6, 20, 20);
                    using (var p = new Pen(Color.White, 2)) { g.DrawLine(p, 12, 16, 16, 20); g.DrawLine(p, 16, 20, 22, 12); }
                }
                return bmp;
            });

        private Bitmap LoadTogglePaneIcon32() => LoadResourceBitmap(
            "TogglePaneIcon32",
            new[] { "TogglePaneIcon32.png", "PaneIcon32.png", "PaneIcon.png" },
            () =>
            {
                var bmp = new Bitmap(32, 32);
                using (var g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    using (var b = new SolidBrush(Color.SteelBlue)) g.FillEllipse(b, 4, 4, 24, 24);
                    using (var p = new Pen(Color.White, 2)) g.DrawLine(p, 12, 10, 12, 22); // simple vertical line inside
                }
                return bmp;
            });

        public stdole.IPictureDisp GetToggleImage(Office.IRibbonControl control) => PictureDispConverter.ToIPictureDisp(LoadTogglePaneIcon32());
        public stdole.IPictureDisp GetGenerateImage(Office.IRibbonControl control) => PictureDispConverter.ToIPictureDisp(LoadGenerateIcon32());
        #endregion

        #region Helper
        private sealed class PictureDispConverter : System.Windows.Forms.AxHost
        {
            private PictureDispConverter() : base(null) { }
            public static stdole.IPictureDisp ToIPictureDisp(Image image) => (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
        #endregion
    }
}