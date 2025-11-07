using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookLMStudio
{
    public class EmailSelectedEventArgs : EventArgs
    {
        public Outlook.MailItem MailItem { get; }

        public EmailSelectedEventArgs(Outlook.MailItem mailItem)
        {
            MailItem = mailItem;
        }
    }

    public class GenerateResponseEventArgs : EventArgs
    {
        public Outlook.MailItem MailItem { get; }
        public string PromptTemplate { get; }

        public GenerateResponseEventArgs(Outlook.MailItem mailItem, string promptTemplate)
        {
            MailItem = mailItem;
            PromptTemplate = promptTemplate;
        }
    }
}