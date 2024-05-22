using System.Globalization;

namespace ExchangeApp.Helper
{
    public class MailWithAttachmentObj
    {
        public string Subject { get; set; }
        public string Body { get; set; }

        public byte[] Attachments { get; set; }
        public List<string> Attendees { get; set; } = new List<string>();




    }
}
