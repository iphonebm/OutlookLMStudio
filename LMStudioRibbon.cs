using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace OutlookLMStudio
{
    [ComVisible(true)]
    public class LMStudioRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public LMStudioRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            // Retourner le XML directement au lieu de le lire depuis un fichier
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
                  imageMso=""HappyFace"" />
          <button id=""btnGenerateSelected"" 
                  label=""Générer Réponses"" 
                  size=""large"" 
                  onAction=""OnGenerateSelectedClick""
                  imageMso=""CreateAReply""
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

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnTogglePaneClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.ToggleTaskPane();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnGenerateSelectedClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.GenerateResponsesForSelection();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public void OnContextGenerateClick(Office.IRibbonControl control)
        {
            try
            {
                Globals.ThisAddIn.GenerateResponsesForSelection();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Erreur : {ex.Message}", "Erreur",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        #endregion
    }
}