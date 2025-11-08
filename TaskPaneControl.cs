using System;
using System.Windows.Forms;
using System.Net.NetworkInformation;
using System.Drawing;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Net;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace OutlookLMStudio
{
    public partial class TaskPaneControl : UserControl
    {
        // UI principal
        private Button btnGenerateResponse;
        private Button btnSettings;
        private FlowLayoutPanel pnlStatus; // contient la puce + file d'attente
        private Label lblDot;              // puce d'état dessinée
        private Label lblQueue;            // affiche "File d'attente: N"
        private TextBox txtPromptTemplate;
        private Label lblSelectedEmail;

        // Panneau paramètres
        private Panel pnlSettings;
        private TextBox txtApiUrl;
        private NumericUpDown nudTimeout;
        private NumericUpDown nudTemperature;
        private NumericUpDown nudMaxTokens;
        private TextBox txtStopSequences;
        private ComboBox cboModelName;
        private Button btnSaveSettings;
        private Button btnRefreshModels;
        private Label lblModelLoading;

        // Contexte
        private Outlook.MailItem _currentMailItem;
        private bool _isConnected;
        private bool _isProcessing;

        public event EventHandler<GenerateResponseEventArgs> GenerateResponseRequested;

        public TaskPaneControl()
        {
            InitializeComponent();
            InitializeEvents();
            LoadSettingsIntoPanel();
            CheckLMStudioConnection();
        }

        private void InitializeComponent()
        {
            btnGenerateResponse = new Button();
            btnSettings = new Button();
            pnlStatus = new FlowLayoutPanel();
            lblDot = new Label();
            lblQueue = new Label();
            txtPromptTemplate = new TextBox();
            lblSelectedEmail = new Label();
            pnlSettings = new Panel();

            AutoScroll = true;

            // Label email sélectionné
            lblSelectedEmail.Text = "Aucun email sélectionné";
            lblSelectedEmail.Dock = DockStyle.Top;
            lblSelectedEmail.Height = 40;
            lblSelectedEmail.TextAlign = ContentAlignment.MiddleLeft;

            // Bouton génération
            btnGenerateResponse.Text = "Générer une réponse";
            btnGenerateResponse.Dock = DockStyle.Top;
            btnGenerateResponse.Height = 30;
            btnGenerateResponse.Enabled = false;

            // Bouton paramètres
            btnSettings.Text = "Paramètres";
            btnSettings.Dock = DockStyle.Top;
            btnSettings.Height = 30;

            // Bandeau statut (puce + file d'attente)
            pnlStatus.Dock = DockStyle.Top;
            pnlStatus.Height = 24;
            pnlStatus.Padding = new Padding(0, 2, 0, 2);
            pnlStatus.WrapContents = false;
            pnlStatus.AutoSize = false;
            pnlStatus.FlowDirection = FlowDirection.LeftToRight;

            // Puce (dessinée)
            lblDot.Width = 24;
            lblDot.Height = 20;
            lblDot.Margin = new Padding(4, 0, 8, 0);
            lblDot.Paint += LblDot_Paint;

            // File d'attente
            lblQueue.AutoSize = true;
            lblQueue.Text = "File d'attente: 0";
            lblQueue.Margin = new Padding(0, 2, 0, 0);

            pnlStatus.Controls.Add(lblDot);
            pnlStatus.Controls.Add(lblQueue);

            // Zone template (prompt)
            txtPromptTemplate.Multiline = true;
            txtPromptTemplate.ScrollBars = ScrollBars.Vertical;
            txtPromptTemplate.Dock = DockStyle.Fill;
            txtPromptTemplate.Text = "Veuillez générer une réponse courtoise et professionnelle à cet email.\r\nConservez le même niveau de formalité que l'email original.\r\n\r\nEmail à traiter :\r\n{emailContent}";

            // Panneau paramètres
            pnlSettings.Dock = DockStyle.Top;
            pnlSettings.Padding = new Padding(8);
            pnlSettings.BorderStyle = BorderStyle.FixedSingle;
            pnlSettings.Visible = false;
            pnlSettings.AutoSize = true;
            pnlSettings.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            BuildSettingsPanelUI();

            // Ordre d'ajout (Fill en dernier)
            Controls.Add(txtPromptTemplate);
            Controls.Add(pnlSettings);
            Controls.Add(pnlStatus);
            Controls.Add(btnSettings);
            Controls.Add(btnGenerateResponse);
            Controls.Add(lblSelectedEmail);

            Name = "TaskPaneControl";
            Size = new Size(300, 400);
        }

        private void BuildSettingsPanelUI()
        {
            txtApiUrl = new TextBox();
            nudTimeout = new NumericUpDown();
            nudTemperature = new NumericUpDown();
            nudMaxTokens = new NumericUpDown();
            txtStopSequences = new TextBox();
            cboModelName = new ComboBox();
            btnSaveSettings = new Button();
            btnRefreshModels = new Button();
            lblModelLoading = new Label();

            int y = 8; int w = 260;
            var lblApiUrl = new Label { Text = "URL de l'API LMStudio:", Location = new Point(8, y), AutoSize = true }; y += 18;
            txtApiUrl.Location = new Point(8, y); txtApiUrl.Size = new Size(w, 20); y += 28;
            var lblTimeout = new Label { Text = "Timeout (secondes):", Location = new Point(8, y), AutoSize = true }; y += 18;
            nudTimeout.Location = new Point(8, y); nudTimeout.Minimum = 1; nudTimeout.Maximum = 300; nudTimeout.Value = 30; y += 28;
            var lblTemperature = new Label { Text = "Température (0.0 - 1.0):", Location = new Point(8, y), AutoSize = true }; y += 18;
            nudTemperature.Location = new Point(8, y); nudTemperature.DecimalPlaces = 2; nudTemperature.Increment = 0.1M; nudTemperature.Minimum = 0; nudTemperature.Maximum = 1; nudTemperature.Value = 0.7M; y += 28;
            var lblMaxTokens = new Label { Text = "Nombre maximum de tokens:", Location = new Point(8, y), AutoSize = true }; y += 18;
            nudMaxTokens.Location = new Point(8, y); nudMaxTokens.Minimum = 100; nudMaxTokens.Maximum = 8000; nudMaxTokens.Value = 2000; y += 28;
            var lblStopSequences = new Label { Text = "Séquences d'arrêt (séparées par |):", Location = new Point(8, y), AutoSize = true }; y += 18;
            txtStopSequences.Location = new Point(8, y); txtStopSequences.Size = new Size(w, 20); y += 28;
            var lblModelName = new Label { Text = "Nom du modèle:", Location = new Point(8, y), AutoSize = true }; y += 18;
            cboModelName.Location = new Point(8, y); cboModelName.Size = new Size(w - 70, 22); cboModelName.DropDownStyle = ComboBoxStyle.DropDownList; cboModelName.Sorted = true;
            cboModelName.SelectionChangeCommitted += (s, e) => SaveSelectedModelImmediate();
            cboModelName.SelectedIndexChanged += (s, e) => SaveSelectedModelImmediate();
            btnRefreshModels.Text = "?"; btnRefreshModels.Location = new Point(8 + w - 60, y - 1); btnRefreshModels.Size = new Size(30, 24); btnRefreshModels.Click += async (s, e) => await LoadModelsAsync();
            lblModelLoading.Location = new Point(8 + w - 30, y + 2); lblModelLoading.Size = new Size(26, 16); lblModelLoading.ForeColor = Color.Gray; y += 34;
            btnSaveSettings.Text = "Enregistrer"; btnSaveSettings.Location = new Point(8, y); btnSaveSettings.Size = new Size(100, 24); btnSaveSettings.Click += BtnSaveSettings_Click;

            pnlSettings.Controls.AddRange(new Control[] { lblApiUrl, txtApiUrl, lblTimeout, nudTimeout, lblTemperature, nudTemperature, lblMaxTokens, nudMaxTokens, lblStopSequences, txtStopSequences, lblModelName, cboModelName, btnRefreshModels, lblModelLoading, btnSaveSettings });
        }

        private void InitializeEvents()
        {
            btnGenerateResponse.Click += BtnGenerateResponse_Click;
            btnSettings.Click += BtnSettings_Click;
            if (Globals.ThisAddIn != null) Globals.ThisAddIn.EmailSelected += ThisAddIn_EmailSelected;
        }

        // Dessin de la puce
        private void LblDot_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            int diameter = 12;
            int x = 4;
            int y = (lblDot.Height - diameter) / 2;
            Color color = !_isConnected ? Color.Red : (_isProcessing ? Color.Orange : Color.Green);
            using (var b = new SolidBrush(color)) using (var p = new Pen(color))
            {
                e.Graphics.FillEllipse(b, x, y, diameter, diameter);
                e.Graphics.DrawEllipse(p, x, y, diameter, diameter);
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
            catch { UpdateConnectionStatus(false); }
        }

        private void UpdateConnectionStatus(bool isConnected)
        {
            _isConnected = isConnected;
            btnGenerateResponse.Enabled = isConnected && _currentMailItem != null && !_isProcessing;
            lblDot.Invalidate();
        }

        public void MarkProcessingStart()
        {
            _isProcessing = true;
            btnGenerateResponse.Enabled = false;
            lblDot.Invalidate();
        }

        public void MarkProcessingEnd()
        {
            _isProcessing = false;
            btnGenerateResponse.Enabled = _isConnected && _currentMailItem != null;
            lblDot.Invalidate();
        }

        public void UpdateQueueLength(int count)
        {
            if (lblQueue.InvokeRequired) { lblQueue.Invoke(new Action<int>(UpdateQueueLength), count); return; }
            lblQueue.Text = $"File d'attente: {count}";
        }

        private async Task LoadModelsAsync()
        {
            try
            {
                lblModelLoading.Text = "...";
                cboModelName.Items.Clear();
                var settings = LMStudioSettings.LoadFromConfig();
                string url = settings.ApiUrl?.TrimEnd('/') + "/v1/models";
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.Accept] = "application/json";
                    string json = await client.DownloadStringTaskAsync(url);
                    var root = JObject.Parse(json);
                    var data = root["data"] as JArray;
                    if (data != null)
                    {
                        foreach (var m in data)
                        {
                            var id = m["id"]?.ToString();
                            if (!string.IsNullOrEmpty(id)) cboModelName.Items.Add(id);
                        }
                    }
                }
                var saved = settings.ModelName;
                if (!string.IsNullOrEmpty(saved))
                {
                    int idx = cboModelName.Items.IndexOf(saved);
                    if (idx >= 0) cboModelName.SelectedIndex = idx; else cboModelName.Items.Insert(0, saved);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Impossible de récupérer la liste des modèles : " + ex.Message, "LMStudio", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally { lblModelLoading.Text = string.Empty; }
        }

        private void SaveSelectedModelImmediate()
        {
            var selected = cboModelName.SelectedItem as string;
            if (!string.IsNullOrEmpty(selected))
            {
                try { var s = LMStudioSettings.LoadFromConfig(); s.ModelName = selected; s.SaveToConfig(); } catch { }
            }
        }

        private void ThisAddIn_EmailSelected(object sender, EmailSelectedEventArgs e)
        {
            _currentMailItem = e.MailItem;
            UpdateSelectedEmailInfo();
        }

        private void UpdateSelectedEmailInfo()
        {
            if (InvokeRequired) { Invoke(new Action(UpdateSelectedEmailInfo)); return; }
            if (_currentMailItem != null)
            {
                lblSelectedEmail.Text = $"Email sélectionné :\n{_currentMailItem.Subject}";
                lblSelectedEmail.BackColor = SystemColors.Highlight;
                lblSelectedEmail.ForeColor = SystemColors.HighlightText;
                btnGenerateResponse.Enabled = _isConnected && !_isProcessing;
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
            if (_currentMailItem != null && !_isProcessing && _isConnected)
            {
                MarkProcessingStart();
                GenerateResponseRequested?.Invoke(this, new GenerateResponseEventArgs(_currentMailItem, txtPromptTemplate.Text));
            }
        }

        private async void BtnSettings_Click(object sender, EventArgs e)
        {
            pnlSettings.Visible = !pnlSettings.Visible;
            if (pnlSettings.Visible) await LoadModelsAsync();
        }

        private void BtnSaveSettings_Click(object sender, EventArgs e)
        {
            try
            {
                var settings = LMStudioSettings.LoadFromConfig();
                settings.ApiUrl = txtApiUrl.Text;
                settings.TimeoutSeconds = (int)nudTimeout.Value;
                settings.Temperature = (double)nudTemperature.Value;
                settings.MaxTokens = (int)nudMaxTokens.Value;
                settings.StopSequences = (txtStopSequences.Text ?? string.Empty).Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                settings.ModelName = cboModelName.SelectedItem?.ToString() ?? cboModelName.Text;
                settings.SaveToConfig();
                MessageBox.Show("Paramètres enregistrés.", "LMStudio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CheckLMStudioConnection();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'enregistrement des paramètres : {ex.Message}", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadSettingsIntoPanel()
        {
            try
            {
                var s = LMStudioSettings.LoadFromConfig();
                txtApiUrl.Text = s.ApiUrl;
                nudTimeout.Value = ClampToRange(s.TimeoutSeconds, (int)nudTimeout.Minimum, (int)nudTimeout.Maximum);
                nudTemperature.Value = ClampDecimalToRange((decimal)s.Temperature, nudTemperature.Minimum, nudTemperature.Maximum);
                nudMaxTokens.Value = ClampToRange(s.MaxTokens, (int)nudMaxTokens.Minimum, (int)nudMaxTokens.Maximum);
                txtStopSequences.Text = string.Join("|", s.StopSequences ?? new string[0]);
            }
            catch { }
        }

        private static int ClampToRange(int v, int min, int max) { if (v < min) return min; if (v > max) return max; return v; }
        private static decimal ClampDecimalToRange(decimal v, decimal min, decimal max) { if (v < min) return min; if (v > max) return max; return v; }

        // Méthode externe appelée quand génération terminée côté add-in
        public void ExternalGenerationEnded() => MarkProcessingEnd();

        protected override void Dispose(bool disposing)
        {
            if (disposing && Globals.ThisAddIn != null) Globals.ThisAddIn.EmailSelected -= ThisAddIn_EmailSelected;
            base.Dispose(disposing);
        }
    }
}