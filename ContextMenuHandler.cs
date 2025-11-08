using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookLMStudio
{
    public class ContextMenuHandler
    {
        public void HandleSelectedMailItems()
        {
            try
            {
                var explorer = Globals.ThisAddIn.Application.ActiveExplorer();
                if (explorer == null || explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    return; // silencieux
                }

                var selectedMailItems = new List<Outlook.MailItem>();
                for (int i = 1; i <= explorer.Selection.Count; i++)
                {
                    var item = explorer.Selection[i];
                    if (item is Outlook.MailItem mailItem)
                        selectedMailItems.Add(mailItem);
                }

                if (selectedMailItems.Count == 0)
                    return;

                // Envoi direct sans confirmation ni fenêtre de résultat
                GenerateResponsesForMultipleEmails(selectedMailItems);
            }
            catch
            {
                // ignorer
            }
        }

        public void HandleMailItemContextMenu(Outlook.MailItem mailItem)
        {
            try
            {
                if (mailItem == null) return;
                GenerateResponseForMail(mailItem);
            }
            catch
            {
                // ignorer
            }
        }

        private void GenerateResponsesForMultipleEmails(List<Outlook.MailItem> mailItems)
        {
            foreach (var mail in mailItems)
            {
                try { GenerateResponseForMail(mail); } catch { }
            }
        }

        private void GenerateResponseForMail(Outlook.MailItem mailItem)
        {
            var promptTemplate = "Veuillez générer une réponse courtoise et professionnelle à cet email.\r\nConservez le même niveau de formalité que l'email original.\r\n\r\nEmail à traiter :\r\n{emailContent}";
            var eventArgs = new GenerateResponseEventArgs(mailItem, promptTemplate);
            Globals.ThisAddIn.OnGenerateResponseRequested(eventArgs);
        }
    }
}