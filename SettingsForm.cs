using System;
using System.Windows.Forms;

namespace OutlookLMStudio
{
    public partial class SettingsForm : Form
    {
        private TextBox txtApiUrl;
        private NumericUpDown nudTimeout;
        private NumericUpDown nudTemperature;
        private NumericUpDown nudMaxTokens;
        private TextBox txtStopSequences;
        private TextBox txtModelName;
        private Button btnSave;
        private Button btnCancel;
        private LMStudioSettings _settings;

        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void InitializeComponent()
        {
            this.txtApiUrl = new TextBox();
            this.nudTimeout = new NumericUpDown();
            this.nudTemperature = new NumericUpDown();
            this.nudMaxTokens = new NumericUpDown();
            this.txtStopSequences = new TextBox();
            this.txtModelName = new TextBox();
            this.btnSave = new Button();
            this.btnCancel = new Button();

            // Configuration du formulaire
            this.Text = "Paramètres LMStudio";
            this.Size = new System.Drawing.Size(400, 400);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            int currentY = 20;
            
            // Configuration des contrôles
            var lblApiUrl = new Label
            {
                Text = "URL de l'API LMStudio:",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.txtApiUrl.Location = new System.Drawing.Point(10, currentY);
            this.txtApiUrl.Size = new System.Drawing.Size(360, 20);

            currentY += 30;
            var lblTimeout = new Label
            {
                Text = "Timeout (secondes):",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.nudTimeout.Location = new System.Drawing.Point(10, currentY);
            this.nudTimeout.Size = new System.Drawing.Size(80, 20);
            this.nudTimeout.Minimum = 1;
            this.nudTimeout.Maximum = 300;
            this.nudTimeout.Value = 30;

            currentY += 30;
            var lblTemperature = new Label
            {
                Text = "Température (0.0 - 1.0):",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.nudTemperature.Location = new System.Drawing.Point(10, currentY);
            this.nudTemperature.Size = new System.Drawing.Size(80, 20);
            this.nudTemperature.DecimalPlaces = 2;
            this.nudTemperature.Increment = 0.1M;
            this.nudTemperature.Minimum = 0.0M;
            this.nudTemperature.Maximum = 1.0M;
            this.nudTemperature.Value = 0.7M;

            currentY += 30;
            var lblMaxTokens = new Label
            {
                Text = "Nombre maximum de tokens:",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.nudMaxTokens.Location = new System.Drawing.Point(10, currentY);
            this.nudMaxTokens.Size = new System.Drawing.Size(120, 20);
            this.nudMaxTokens.Minimum = 100;
            this.nudMaxTokens.Maximum = 8000;
            this.nudMaxTokens.Value = 2000;

            currentY += 30;
            var lblStopSequences = new Label
            {
                Text = "Séquences d'arrêt (séparées par |):",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.txtStopSequences.Location = new System.Drawing.Point(10, currentY);
            this.txtStopSequences.Size = new System.Drawing.Size(360, 20);

            currentY += 30;
            var lblModelName = new Label
            {
                Text = "Nom du modèle:",
                Location = new System.Drawing.Point(10, currentY),
                AutoSize = true
            };

            currentY += 25;
            this.txtModelName.Location = new System.Drawing.Point(10, currentY);
            this.txtModelName.Size = new System.Drawing.Size(360, 20);

            currentY += 40;
            this.btnSave.Text = "Enregistrer";
            this.btnSave.DialogResult = DialogResult.OK;
            this.btnSave.Location = new System.Drawing.Point(200, currentY);
            this.btnSave.Click += BtnSave_Click;

            this.btnCancel.Text = "Annuler";
            this.btnCancel.DialogResult = DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(290, currentY);

            // Ajout des contrôles au formulaire
            this.Controls.AddRange(new Control[] {
                lblApiUrl, this.txtApiUrl,
                lblTimeout, this.nudTimeout,
                lblTemperature, this.nudTemperature,
                lblMaxTokens, this.nudMaxTokens,
                lblStopSequences, this.txtStopSequences,
                lblModelName, this.txtModelName,
                this.btnSave, this.btnCancel
            });
        }

        private void LoadSettings()
        {
            _settings = LMStudioSettings.LoadFromConfig();
            
            txtApiUrl.Text = _settings.ApiUrl;
            nudTimeout.Value = _settings.TimeoutSeconds;
            nudTemperature.Value = (decimal)_settings.Temperature;
            nudMaxTokens.Value = _settings.MaxTokens;
            txtStopSequences.Text = string.Join("|", _settings.StopSequences);
            txtModelName.Text = _settings.ModelName;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                _settings.ApiUrl = txtApiUrl.Text;
                _settings.TimeoutSeconds = (int)nudTimeout.Value;
                _settings.Temperature = (double)nudTemperature.Value;
                _settings.MaxTokens = (int)nudMaxTokens.Value;
                _settings.StopSequences = txtStopSequences.Text.Split('|');
                _settings.ModelName = txtModelName.Text;

                _settings.SaveToConfig();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur lors de l'enregistrement des paramètres : {ex.Message}",
                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.DialogResult = DialogResult.None;
            }
        }
    }
}