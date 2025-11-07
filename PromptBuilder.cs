using System;
using System.Text;

namespace OutlookLMStudio
{
    /// <summary>
    /// Classe helper pour construire des prompts pour LMStudio
    /// Séparée de ThisAddIn pour faciliter les tests unitaires
    /// </summary>
    public static class PromptBuilder
    {
        /// <summary>
        /// Construit un prompt à partir des données d'email et d'un template
        /// </summary>
        public static string BuildPrompt(EmailData emailData, string promptTemplate)
        {
            if (emailData == null)
                throw new ArgumentNullException(nameof(emailData));
            
            if (string.IsNullOrEmpty(promptTemplate))
                throw new ArgumentException("Le template de prompt ne peut pas être vide", nameof(promptTemplate));

            var emailContent = FormatEmailContent(emailData);
            return promptTemplate.Replace("{emailContent}", emailContent);
        }

        /// <summary>
        /// Formate le contenu d'un email pour le prompt
        /// </summary>
        private static string FormatEmailContent(EmailData emailData)
        {
            var content = new StringBuilder();
            
            content.AppendLine($"De: {emailData.SenderName} <{emailData.SenderEmailAddress}>");
            content.AppendLine($"À: {emailData.To}");
            
            if (!string.IsNullOrEmpty(emailData.CC))
                content.AppendLine($"Cc: {emailData.CC}");
            
            content.AppendLine($"Objet: {emailData.Subject}");
            content.AppendLine($"Date: {emailData.ReceivedTime:g}");
            content.AppendLine();
            content.AppendLine("Corps du message:");
            content.AppendLine(new string('-', 50));

            var body = emailData.Body ?? string.Empty;
            body = body.Replace("\r\n", "\n").Trim();
            content.AppendLine(body);
            
            content.AppendLine(new string('-', 50));

            return content.ToString();
        }
    }

    /// <summary>
    /// Classe représentant les données d'un email pour le PromptBuilder
    /// </summary>
    public class EmailData
    {
        public string SenderName { get; set; }
        public string SenderEmailAddress { get; set; }
        public string To { get; set; }
        public string CC { get; set; }
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime ReceivedTime { get; set; }
    }
}