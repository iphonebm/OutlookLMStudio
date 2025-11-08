using System;
using System.Drawing;
using System.Net.NetworkInformation;
using System.Windows.Forms;
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

        private const string DefaultPromptTemplate = "Système : Vous êtes un assistant professionnel.\r\n" +
            "Veuillez générer une réponse courtoise et professionnelle à cet email.\r\n" +
            "Conservez le même niveau de formalité que l'email original.\r\n\r\n" +
            "Email à traiter :\r\n{emailContent}\r\n\r\n" +
            "Générez une réponse professionnelle et pertinente.";

        public TaskPaneControl()
        {
            InitializeComponent();
            InitializeEvents();
            CheckLMStudioConnection();
        }

        private void InitializeComponent()
        {
            SuspendLayout();
            btnGenerateResponse = new Button();
            btnSettings = new Button();
            lblStatus = new Label();
            txtPromptTemplate = new TextBox();
            lblSelectedEmail = new Label();
            lblSelectedEmail.Text = "Aucun email sélectionné";
            lblSelectedEmail.Dock = DockStyle.Top;
            lblSelectedEmail.Height = 40;
            lblSelectedEmail.AutoSize = false;
            lblSelectedEmail.Padding = new Padding(4);
            btnGenerateResponse.Text = "Générer une réponse";
            btnGenerateResponse.Dock = DockStyle.Top;
            btnGenerateResponse.Height = 30;
            btnGenerateResponse.Enabled = false;
            btnSettings.Text = "Paramètres";
            btnSettings.Dock = DockStyle.Top;
            btnSettings.Height = 30;
            lblStatus.Text = "Vérification de la connexion...";
            lblStatus.Dock = DockStyle.Top;
            lblStatus.Height = 20;
            lblStatus.TextAlign = ContentAlignment.MiddleLeft;
            lblStatus.Padding = new Padding(4, 0, 0, 0);
            txtPromptTemplate.Multiline = true;
            txtPromptTemplate.ScrollBars = ScrollBars.Vertical;
            txtPromptTemplate.Dock = DockStyle.Fill;
            txtPromptTemplate.Text = DefaultPromptTemplate;
            txtPromptTemplate.Font = new Font(FontFamily.GenericMonospace, 9f);
            Controls.Add(txtPromptTemplate);
            Controls.Add(lblStatus);
            Controls.Add(btnSettings);
            Controls.Add(btnGenerateResponse);
            Controls.Add(lblSelectedEmail);
            Name = "TaskPaneControl";
            Size = new Size(300, 400);
            ResumeLayout();
        }

        private void InitializeEvents()
        {
            btnGenerateResponse.Click += BtnGenerateResponse_Click;
            btnSettings.Click += BtnSettings_Click;
            if (Globals.ThisAddIn != null)
            {
                Globals.ThisAddIn.EmailSelected += ThisAddIn_EmailSelected;
            }
            else
            {
                MessageBox.Show("ERREUR: Globals.ThisAddIn est null - la sélection d'email ne fonctionnera pas!", "Erreur d'initialisation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void CheckLMStudioConnection()
        {
            try
            {
                using (var ping = new Ping())
                {
                    var reply = await ping.SendPingAsync("localhost");
                    UpdateConnectionStatus(reply.Status == IPStatus.Success);
                }
            }
            catch
            {
                UpdateConnectionStatus(false);
            }
        }

        private void UpdateConnectionStatus(bool isConnected)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => UpdateConnectionStatus(isConnected)));
                return;
            }
            lblStatus.Text = isConnected ? "LMStudio: Connecté" : "LMStudio: Non connecté";
            lblStatus.ForeColor = isConnected ? Color.Green : Color.Red;
            btnGenerateResponse.Enabled = isConnected && _currentMailItem != null;
        }

        private void ThisAddIn_EmailSelected(object sender, EmailSelectedEventArgs e)
        {
            _currentMailItem = e.MailItem;
            UpdateSelectedEmailInfo();
        }

        private void UpdateSelectedEmailInfo()
        {
            if (InvokeRequired)
            {
                Invoke(new Action(UpdateSelectedEmailInfo));
                return;
            }
            if (_currentMailItem != null)
            {
                lblSelectedEmail.Text = $"Email sélectionné :\n{_currentMailItem.Subject}";
                lblSelectedEmail.BackColor = SystemColors.Highlight;
                lblSelectedEmail.ForeColor = SystemColors.HighlightText;
                btnGenerateResponse.Enabled = true;
            }
            else
            {
                lblSelectedEmail.Text = "Aucun email sélectionné";
                lblSelectedEmail.BackColor = SystemColors.Window;
                lblSelectedEmail.ForeColor = SystemColors.WindowText;
                btnGenerateResponse.Enabled = false;
            }
        }

        private void BtnGenerateResponse_Click(object sender, EventArgs e)
        {
            if (_currentMailItem == null) return;
            try
            {
                btnGenerateResponse.Enabled = false;
                GenerateResponseRequested?.Invoke(this, new GenerateResponseEventArgs(_currentMailItem, txtPromptTemplate.Text));
            }
            finally
            {
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
            if (disposing)
            {
                if (Globals.ThisAddIn != null)
                {
                    Globals.ThisAddIn.EmailSelected -= ThisAddIn_EmailSelected;
                }
            }
            base.Dispose(disposing);
        }
    }
}