using System;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Drawing;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookLMStudio
{
    public partial class TaskPaneControl : UserControl
    {
        private Button btnGenerateResponse;
        private Button btnSettings;
        private Label lblStatus;
        private TextBox txtPromptTemplate;
        private Label lblSelectedEmail;
        private Outlook.MailItem _currentMailItem;

        public event EventHandler<GenerateResponseEventArgs> GenerateResponseRequested;

        public TaskPaneControl()
        {
            InitializeComponent();
            InitializeEvents();
            CheckLMStudioConnection();
        }

        private void InitializeComponent()
        {
            this.btnGenerateResponse = new Button();
            this.btnSettings = new Button();
            this.lblStatus = new Label();
            this.txtPromptTemplate = new TextBox();
            this.lblSelectedEmail = new Label();

            // Configuration du label de l'email sélectionné
            this.lblSelectedEmail.Text = "Aucun email sélectionné";
            this.lblSelectedEmail.Dock = DockStyle.Top;
            this.lblSelectedEmail.Height = 40;
            this.lblSelectedEmail.AutoSize = false;

            // Configuration du bouton Generate Response
            this.btnGenerateResponse.Text = "Générer une réponse";
            this.btnGenerateResponse.Dock = DockStyle.Top;
            this.btnGenerateResponse.Height = 30;
            this.btnGenerateResponse.Enabled = false;

            // Configuration du bouton Settings
            this.btnSettings.Text = "Paramètres";
            this.btnSettings.Dock = DockStyle.Top;
            this.btnSettings.Height = 30;

            // Configuration du label de statut
            this.lblStatus.Text = "Vérification de la connexion...";
            this.lblStatus.Dock = DockStyle.Top;
            this.lblStatus.Height = 20;

            // Configuration du template de prompt
            this.txtPromptTemplate.Multiline = true;
            this.txtPromptTemplate.ScrollBars = ScrollBars.Vertical;
            this.txtPromptTemplate.Dock = DockStyle.Fill;
            this.txtPromptTemplate.Text = @"Système : Vous êtes un assistant professionnel. 
Veuillez générer une réponse courtoise et professionnelle à cet email.
Conservez le même niveau de formalité que l'email original.

Email à traiter :
{emailContent}

Générez une réponse professionnelle et pertinente.";

            // Ajout des contrôles au TaskPane
            this.Controls.Add(this.txtPromptTemplate);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.btnGenerateResponse);
            this.Controls.Add(this.lblSelectedEmail);

            this.Name = "TaskPaneControl";
            this.Size = new System.Drawing.Size(300, 400);
        }

        private void InitializeEvents()
        {
            this.btnGenerateResponse.Click += BtnGenerateResponse_Click;
            this.btnSettings.Click += BtnSettings_Click;

            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.EmailSelected += ThisAddIn_EmailSelected;
                System.Diagnostics.Debug.WriteLine("TaskPaneControl: Événement EmailSelected enregistré");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("TaskPaneControl: ERREUR - Globals.ThisAddIn est NULL!");
                MessageBox.Show("ERREUR: Globals.ThisAddIn est null - la sélection d'email ne fonctionnera pas!", 
                    "Erreur d'initialisation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void CheckLMStudioConnection()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = await ping.SendPingAsync("localhost");
                    if (reply.Status == IPStatus.Success)
                    {
                        UpdateConnectionStatus(true);
                    }
                    else
                    {
                        UpdateConnectionStatus(false);
                    }
                }
            }
            catch
            {
                UpdateConnectionStatus(false);
            }
        }

        private void UpdateConnectionStatus(bool isConnected)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action(() => UpdateConnectionStatus(isConnected)));
                return;
            }

            lblStatus.Text = isConnected ? "LMStudio: Connecté" : "LMStudio: Non connecté";
            lblStatus.ForeColor = isConnected ? Color.Green : Color.Red;
            btnGenerateResponse.Enabled = isConnected && _currentMailItem != null;
        }

        private void ThisAddIn_EmailSelected(object sender, EmailSelectedEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"TaskPaneControl: ThisAddIn_EmailSelected appelé pour: {e.MailItem?.Subject ?? "null"}");
            _currentMailItem = e.MailItem;
            UpdateSelectedEmailInfo();
        }

        private void UpdateSelectedEmailInfo()
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new System.Action(UpdateSelectedEmailInfo));
                return;
            }

            System.Diagnostics.Debug.WriteLine($"TaskPaneControl: UpdateSelectedEmailInfo - _currentMailItem est {(_currentMailItem == null ? "NULL" : "non-null")}");

            if (_currentMailItem != null)
            {
                lblSelectedEmail.Text = $"Email sélectionné :\n{_currentMailItem.Subject}";
                lblSelectedEmail.BackColor = SystemColors.Highlight;
                lblSelectedEmail.ForeColor = SystemColors.HighlightText;
                btnGenerateResponse.Enabled = true;
                System.Diagnostics.Debug.WriteLine($"TaskPaneControl: Email affiché - {_currentMailItem.Subject}");
            }
            else
            {
                lblSelectedEmail.Text = "Aucun email sélectionné";
                lblSelectedEmail.BackColor = SystemColors.Window;
                lblSelectedEmail.ForeColor = SystemColors.WindowText;
                btnGenerateResponse.Enabled = false;
                System.Diagnostics.Debug.WriteLine("TaskPaneControl: Aucun email sélectionné");
            }
        }

        private void BtnGenerateResponse_Click(object sender, EventArgs e)
        {
            if (_currentMailItem != null)
            {
                btnGenerateResponse.Enabled = false;
                GenerateResponseRequested?.Invoke(this, new GenerateResponseEventArgs(_currentMailItem, txtPromptTemplate.Text));
                btnGenerateResponse.Enabled = true;
            }
        }

        private void BtnSettings_Click(object sender, EventArgs e)
        {
            using (var settingsForm = new SettingsForm())
            {
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    CheckLMStudioConnection();
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.EmailSelected -= ThisAddIn_EmailSelected;
            }
            base.Dispose(disposing);
        }
    }
}