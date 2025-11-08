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
                    MessageBox.Show("Aucun email sélectionné.", "Information", 
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Récupérer tous les emails sélectionnés
                var selectedMailItems = new List<Outlook.MailItem>();
                for (int i = 1; i <= explorer.Selection.Count; i++)
                {
                    var item = explorer.Selection[i];
                    if (item is Outlook.MailItem mailItem)
                    {
                        selectedMailItems.Add(mailItem);
                    }
                }

                if (selectedMailItems.Count == 0)
                {
                    MessageBox.Show("Aucun email trouvé dans la sélection.\n\nAssurez-vous de sélectionner des emails, pas des rendez-vous ou contacts.", 
                        "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Confirmer avec l'utilisateur
                var result = MessageBox.Show(
                    $"Générer des réponses pour {selectedMailItems.Count} email(s) ?\n\n" +
                    $"Les réponses seront créées en séquence.\n" +
                    $"Cela peut prendre quelques instants...",
                    "Confirmation", 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    GenerateResponsesForMultipleEmails(selectedMailItems);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void HandleMailItemContextMenu(Outlook.MailItem mailItem)
        {
            try
            {
                if (mailItem == null) return;
                
                var result = MessageBox.Show(
                    $"Générer une réponse pour cet email ?\n\nSujet : {mailItem.Subject}",
                    "Confirmation", 
                    MessageBoxButtons.YesNo, 
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    GenerateResponseForMail(mailItem);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenerateResponsesForMultipleEmails(List<Outlook.MailItem> mailItems)
        {
            int successCount = 0;
            int errorCount = 0;
            var errors = new List<string>();

            var progressForm = new Form
            {
                Text = "Génération en cours...",
                Width = 400,
                Height = 150,
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false
            };

            var progressLabel = new Label
            {
                Text = $"Traitement de 0 / {mailItems.Count} emails...",
                Dock = DockStyle.Top,
                Height = 30,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                Padding = new Padding(10)
            };

            var progressBar = new ProgressBar
            {
                Dock = DockStyle.Top,
                Height = 30,
                Minimum = 0,
                Maximum = mailItems.Count,
                Value = 0
            };

            var cancelButton = new Button
            {
                Text = "Annuler",
                Dock = DockStyle.Bottom,
                Height = 40
            };

            bool cancelled = false;
            cancelButton.Click += (s, e) =>
            {
                cancelled = true;
                progressForm.Close();
            };

            progressForm.Controls.Add(progressLabel);
            progressForm.Controls.Add(progressBar);
            progressForm.Controls.Add(cancelButton);

            progressForm.Show();
            Application.DoEvents();

            for (int i = 0; i < mailItems.Count && !cancelled; i++)
            {
                var mailItem = mailItems[i];
                progressLabel.Text = $"Traitement de {i + 1} / {mailItems.Count} emails...\nSujet : {(mailItem.Subject.Length > 50 ? mailItem.Subject.Substring(0, 50) + "..." : mailItem.Subject)}";
                progressBar.Value = i + 1;
                Application.DoEvents();

                try
                {
                    GenerateResponseForMail(mailItem);
                    successCount++;
                    System.Threading.Thread.Sleep(500); // Pause entre chaque requête
                }
                catch (Exception ex)
                {
                    errorCount++;
                    errors.Add($"• {mailItem.Subject}: {ex.Message}");
                }
            }

            progressForm.Close();

            // Afficher le résumé
            var summary = $"Génération terminée !\n\n" +
                         $"? Réussis : {successCount}\n" +
                         $"? Erreurs : {errorCount}";

            if (errors.Count > 0)
            {
                summary += $"\n\nDétails des erreurs :\n" + string.Join("\n", errors.Take(5));
                if (errors.Count > 5)
                {
                    summary += $"\n... et {errors.Count - 5} autre(s) erreur(s)";
                }
            }

            MessageBox.Show(summary, "Résultat", MessageBoxButtons.OK, 
                errorCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        private void GenerateResponseForMail(Outlook.MailItem mailItem)
        {
            var promptTemplate = @"Système : Vous êtes un assistant professionnel. 
Veuillez générer une réponse courtoise et professionnelle à cet email.
Conservez le même niveau de formalité que l'email original.

Email à traiter :
{emailContent}

Générez une réponse professionnelle et pertinente.";

            var eventArgs = new GenerateResponseEventArgs(mailItem, promptTemplate);
            Globals.ThisAddIn.OnGenerateResponseRequested(eventArgs);
        }
    }
}