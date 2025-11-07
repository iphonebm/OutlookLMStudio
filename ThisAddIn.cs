using System;
using System.Configuration;
using System.Net;
using System.Text;
using System.IO;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookLMStudio
{
    public partial class ThisAddIn
    {
        private WebClient _webClient;
        private TaskPaneControl _taskPaneControl;
        private Microsoft.Office.Tools.CustomTaskPane _customTaskPane;
        private Outlook.Explorers _explorers;
        private Outlook.Inspectors _inspectors;
        private bool _isDisposed;
        private ConcurrentQueue<GenerationRequest> _requestQueue = new ConcurrentQueue<GenerationRequest>();
        private SemaphoreSlim _processingSemaphore = new SemaphoreSlim(1, 1);
        private bool _isProcessing;
        private ContextMenuHandler _contextMenuHandler;

        public event EventHandler<EmailSelectedEventArgs> EmailSelected;

        // Méthodes publiques pour le Ribbon
        public void ToggleTaskPane()
        {
            if (_customTaskPane != null)
            {
                _customTaskPane.Visible = !_customTaskPane.Visible;
                Logger.Log($"Volet TaskPane: {(_customTaskPane.Visible ? "Affiché" : "Masqué")}");
            }
        }

        public bool IsTaskPaneVisible()
        {
            return _customTaskPane != null && _customTaskPane.Visible;
        }

        public void GenerateResponsesForSelection()
        {
            try
            {
                Logger.Log("GenerateResponsesForSelection: Début");
                _contextMenuHandler?.HandleSelectedMailItems();
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur dans GenerateResponsesForSelection", ex);
                MessageBox.Show($"Erreur : {ex.Message}", "Erreur", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new LMStudioRibbon();
        }

        public void OnGenerateResponseRequested(GenerateResponseEventArgs e)
        {
            TaskPaneControl_GenerateResponseRequested(this, e);
        }

        private static class Logger
        {
            private static readonly string LogPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "OutlookLMStudio",
                "logs.txt"
            );

            public static void Log(string message, System.Exception ex = null)
            {
                try
                {
                    var logDir = Path.GetDirectoryName(LogPath);
                    if (!Directory.Exists(logDir))
                        Directory.CreateDirectory(logDir);

                    var logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
                    if (ex != null)
                        logMessage += $"\nException: {ex.Message}\nStack: {ex.StackTrace}";

                    File.AppendAllText(LogPath, logMessage + "\n\n");
                }
                catch
                {
                    // Ignorer les erreurs de journalisation
                }
            }
        }

        private void InitializeAddInComponents()
        {
            try
            {
                Logger.Log("Initialisation des composants du complément");
                
                _webClient = new WebClient
                {
                    Encoding = Encoding.UTF8
                };
                _webClient.Headers[HttpRequestHeader.ContentType] = "application/json";

                _taskPaneControl = new TaskPaneControl();
                _customTaskPane = this.CustomTaskPanes.Add(_taskPaneControl, "LMStudio Assistant");
                _customTaskPane.Visible = true;
                _customTaskPane.Width = 300;

                _taskPaneControl.GenerateResponseRequested += TaskPaneControl_GenerateResponseRequested;
                
                // Initialiser le menu contextuel
                _contextMenuHandler = new ContextMenuHandler();
                
                Logger.Log("Composants initialisés avec succès");
                Logger.Log("Ruban LMStudio Assistant ajouté à la barre d'outils");
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur lors de l'initialisation des composants", ex);
                throw;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                Logger.Log("=== Démarrage du complément OutlookLMStudio ===");
                InitializeAddInComponents();
                SetupOutlookHandlers();
                Logger.Log("=== Complément démarré avec succès ===");
            }
            catch (System.Exception ex)
            {
                Logger.Log("ERREUR CRITIQUE lors du démarrage", ex);
                MessageBox.Show($"Erreur lors de démarrage du complément LMStudio :\n\n{ex.Message}\n\nConsultez les logs pour plus de détails.",
                    "Erreur OutlookLMStudio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task<bool> CheckLMStudioAvailability(string apiUrl)
        {
            try
            {
                using (var client = new WebClient())
                {
                    client.Headers[HttpRequestHeader.Accept] = "application/json";
                    var response = await client.DownloadStringTaskAsync($"{apiUrl}/v1/models");
                    return !string.IsNullOrEmpty(response);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log($"LMStudio non disponible à {apiUrl}", ex);
                return false;
            }
        }

        private async Task<string> GenerateResponse(Outlook.MailItem mailItem, string promptTemplate)
        {
            var settings = LMStudioSettings.LoadFromConfig();

            if (!await CheckLMStudioAvailability(settings.ApiUrl))
            {
                throw new System.Exception("LMStudio n'est pas accessible. Vérifiez que le service est démarré sur " + settings.ApiUrl);
            }

            var prompt = BuildPrompt(mailItem, promptTemplate);
            
            // Format correct pour l'API LMStudio (compatible OpenAI)
            var request = new
            {
                model = settings.ModelName,
                prompt = prompt,
                max_tokens = settings.MaxTokens,
                temperature = settings.Temperature,
                top_p = 1.0,
                n = 1,
                stream = false,
                stop = settings.StopSequences
            };

            var requestJson = JsonConvert.SerializeObject(request);
            Logger.Log($"Requête LMStudio - Modèle: {settings.ModelName}, Tokens: {settings.MaxTokens}, Temp: {settings.Temperature}");
            Logger.Log($"JSON envoyé: {requestJson}");

            try
            {
                using (var client = new WebClient())
                {
                    client.Encoding = Encoding.UTF8;
                    client.Headers[HttpRequestHeader.ContentType] = "application/json";
                    client.Headers[HttpRequestHeader.Accept] = "application/json";
                    
                    var apiEndpoint = $"{settings.ApiUrl}/v1/completions";
                    Logger.Log($"Envoi vers: {apiEndpoint}");
                    
                    var response = await client.UploadStringTaskAsync(apiEndpoint, "POST", requestJson);
                    Logger.Log($"Réponse brute reçue: {response}");
                    
                    var result = JsonConvert.DeserializeObject<LMStudioResponse>(response);

                    if (result?.Choices == null || result.Choices.Length == 0)
                    {
                        Logger.Log("ERREUR: Aucun choix dans la réponse");
                        throw new System.Exception("Aucune réponse générée par LMStudio");
                    }

                    var generatedText = result.Choices[0].Text?.Trim() ?? string.Empty;
                    Logger.Log($"Texte généré: {generatedText.Length} caractères");
                    return generatedText;
                }
            }
            catch (WebException ex)
            {
                // Lire la réponse d'erreur pour plus de détails
                string errorDetails = string.Empty;
                if (ex.Response != null)
                {
                    using (var reader = new System.IO.StreamReader(ex.Response.GetResponseStream()))
                    {
                        errorDetails = reader.ReadToEnd();
                        Logger.Log($"Détails de l'erreur serveur: {errorDetails}");
                    }
                }
                
                if (ex.Status == WebExceptionStatus.Timeout)
                {
                    Logger.Log($"Timeout après {settings.TimeoutSeconds} secondes", ex);
                    throw new System.Exception($"La requête a dépassé le délai d'attente de {settings.TimeoutSeconds} secondes. Augmentez le timeout dans les paramètres.");
                }
                
                Logger.Log($"Erreur réseau: {ex.Message} - Détails: {errorDetails}", ex);
                throw new System.Exception($"Erreur LMStudio: {ex.Message}\n\nDétails: {errorDetails}\n\nVérifiez que:\n1. LMStudio est démarré\n2. Un modèle est chargé\n3. Le serveur local est actif sur {settings.ApiUrl}");
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur lors de la génération de la réponse", ex);
                throw;
            }
        }

        private void SetupOutlookHandlers()
        {
            try
            {
                _explorers = this.Application.Explorers;
                _inspectors = this.Application.Inspectors;

                var explorer = this.Application.ActiveExplorer();
                if (explorer != null)
                {
                    explorer.SelectionChange += Explorer_SelectionChange;
                    Logger.Log("Gestionnaire de sélection configuré pour l'explorateur actif");
                }

                _explorers.NewExplorer += Explorers_NewExplorer;
                _inspectors.NewInspector += Inspectors_NewInspector;
                
                Logger.Log("Tous les gestionnaires Outlook configurés");
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur lors de la configuration des gestionnaires Outlook", ex);
                throw;
            }
        }

        private void Explorers_NewExplorer(Outlook.Explorer explorer)
        {
            try
            {
                explorer.SelectionChange += Explorer_SelectionChange;
                Logger.Log("Nouvelle fenêtre explorateur détectée");
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur lors de l'ajout du gestionnaire au nouvel explorateur", ex);
            }
        }

        private void Inspectors_NewInspector(Outlook.Inspector inspector)
        {
            try
            {
                if (inspector.CurrentItem is Outlook.MailItem mailItem)
                {
                    OnEmailSelected(mailItem);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log("Erreur dans Inspectors_NewInspector", ex);
            }
        }

        private void Explorer_SelectionChange()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Explorer_SelectionChange: Méthode appelée");
                Logger.Log("Explorer_SelectionChange: Événement déclenché");

                var explorer = this.Application.ActiveExplorer();
                if (explorer == null)
                {
                    System.Diagnostics.Debug.WriteLine("Explorer_SelectionChange: Explorer est NULL");
                    Logger.Log("Explorer_SelectionChange: Explorer est NULL");
                    return;
                }

                System.Diagnostics.Debug.WriteLine($"Explorer_SelectionChange: Selection.Count = {explorer.Selection?.Count ?? 0}");
                Logger.Log($"Explorer_SelectionChange: Selection.Count = {explorer.Selection?.Count ?? 0}");

                if (explorer.Selection == null || explorer.Selection.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("Explorer_SelectionChange: Aucune sélection");
                    Logger.Log("Explorer_SelectionChange: Aucune sélection");
                    return;
                }

                var item = explorer.Selection[1];
                System.Diagnostics.Debug.WriteLine($"Explorer_SelectionChange: Type d'élément = {item?.GetType().Name ?? "null"}");
                Logger.Log($"Explorer_SelectionChange: Type d'élément = {item?.GetType().Name ?? "null"}");

                if (item is Outlook.MailItem mailItem)
                {
                    System.Diagnostics.Debug.WriteLine($"Explorer_SelectionChange: MailItem trouvé - {mailItem.Subject}");
                    Logger.Log($"Explorer_SelectionChange: MailItem trouvé - {mailItem.Subject}");
                    OnEmailSelected(mailItem);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Explorer_SelectionChange: L'élément n'est PAS un MailItem");
                    Logger.Log($"Explorer_SelectionChange: L'élément n'est PAS un MailItem");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Explorer_SelectionChange: EXCEPTION - {ex.Message}");
                Logger.Log("Erreur lors de la sélection d'email", ex);
            }
        }

        private void OnEmailSelected(Outlook.MailItem mailItem)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"OnEmailSelected: Début - {mailItem.Subject}");
                Logger.Log($"OnEmailSelected: Début - {mailItem.Subject}");

                System.Diagnostics.Debug.WriteLine($"OnEmailSelected: EmailSelected a {(EmailSelected == null ? "0" : EmailSelected.GetInvocationList().Length.ToString())} abonnés");

                EmailSelected?.Invoke(this, new EmailSelectedEventArgs(mailItem));

                System.Diagnostics.Debug.WriteLine($"OnEmailSelected: Événement déclenché avec succès");
                Logger.Log($"Email sélectionné: {mailItem.Subject}");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnEmailSelected: EXCEPTION - {ex.Message}");
                Logger.Log("Erreur lors de l'événement EmailSelected", ex);
            }
        }

        private readonly struct GenerationRequest
        {
            public Outlook.MailItem MailItem { get; }
            public string PromptTemplate { get; }
            public TaskCompletionSource<string> CompletionSource { get; }

            public GenerationRequest(Outlook.MailItem mailItem, string promptTemplate)
            {
                MailItem = mailItem;
                PromptTemplate = promptTemplate;
                CompletionSource = new TaskCompletionSource<string>();
            }
        }

        private async void TaskPaneControl_GenerateResponseRequested(object sender, GenerateResponseEventArgs e
)
        {
            try
            {
                Logger.Log($"=== Début génération pour: {e.MailItem.Subject} ===");
                
                var request = new GenerationRequest(e.MailItem, e.PromptTemplate);
                _requestQueue.Enqueue(request);
                await ProcessQueueAsync();

                var response = await request.CompletionSource.Task;
                if (!string.IsNullOrEmpty(response))
                {
                    var replyMail = e.MailItem.Reply();
                    replyMail.Body = response + "\n\n--- Généré automatiquement par LMStudio ---";
                    replyMail.Display();
                    Logger.Log("=== Réponse générée et affichée avec succès ===");
                }
                else
                {
                    Logger.Log("ATTENTION: Réponse vide reçue de LMStudio");
                    MessageBox.Show("La réponse générée est vide. Vérifiez votre modèle LMStudio.",
                        "Réponse vide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (System.Exception ex)
            {
                Logger.Log("ERREUR lors de la génération de la réponse", ex);
                MessageBox.Show($"Erreur lors de la génération de la réponse :\n\n{ex.Message}\n\nConsultez les logs pour plus de détails.",
                    "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async Task ProcessQueueAsync()
        {
            if (_isProcessing)
            {
                Logger.Log("Traitement déjà en cours, requête en attente");
                return;
            }

            await _processingSemaphore.WaitAsync();
            try
            {
                _isProcessing = true;
                while (_requestQueue.TryDequeue(out var request))
                {
                    try
                    {
                        Logger.Log("Traitement d'une requête de la file");
                        var response = await GenerateResponse(request.MailItem, request.PromptTemplate);
                        request.CompletionSource.SetResult(response);
                    }
                    catch (System.Exception ex)
                    {
                        Logger.Log("Erreur lors du traitement d'une requête", ex);
                        request.CompletionSource.SetException(ex);
                    }
                }
            }
            finally
            {
                _isProcessing = false;
                _processingSemaphore.Release();
            }
        }

        private string BuildPrompt(Outlook.MailItem mailItem, string promptTemplate)
        {
            var emailContent = new StringBuilder();
            emailContent.AppendLine($"De: {mailItem.SenderName} <{mailItem.SenderEmailAddress}>");
            emailContent.AppendLine($"À: {mailItem.To}");
            if (!string.IsNullOrEmpty(mailItem.CC))
                emailContent.AppendLine($"Cc: {mailItem.CC}");
            emailContent.AppendLine($"Objet: {mailItem.Subject}");
            emailContent.AppendLine($"Date: {mailItem.ReceivedTime:g}");
            emailContent.AppendLine();
            emailContent.AppendLine("Corps du message:");
            emailContent.AppendLine(new string('-', 50));

            var body = mailItem.Body ?? string.Empty;
            body = body.Replace("\r\n", "\n").Trim();
            emailContent.AppendLine(body);
            emailContent.AppendLine(new string('-', 50));

            var finalPrompt = promptTemplate.Replace("{emailContent}", emailContent.ToString());
            Logger.Log($"Prompt construit: {finalPrompt.Length} caractères");
            return finalPrompt;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Logger.Log("=== Arrêt du complément OutlookLMStudio ===");
            CleanupResources();
        }

        private void CleanupResources()
        {
            if (!_isDisposed)
            {
                try
                {
                    Logger.Log("Début du nettoyage des ressources");
                    
                    _webClient?.Dispose();
                    _processingSemaphore?.Dispose();

                    if (this.Application?.ActiveExplorer() != null)
                    {
                        this.Application.ActiveExplorer().SelectionChange -= Explorer_SelectionChange;
                    }

                    if (_explorers != null)
                    {
                        _explorers.NewExplorer -= Explorers_NewExplorer;
                    }

                    if (_inspectors != null)
                    {
                        _inspectors.NewInspector -= Inspectors_NewInspector;
                    }

                    while (_requestQueue.TryDequeue(out var request))
                    {
                        request.CompletionSource.TrySetCanceled();
                    }

                    Logger.Log("Ressources nettoyées avec succès");
                }
                catch (System.Exception ex)
                {
                    Logger.Log("Erreur lors du nettoyage des ressources", ex);
                }
                finally
                {
                    _isDisposed = true;
                    Logger.Log("=== Complément arrêté ===");
                }
            }
        }

        #region Code généré par VSTO
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
