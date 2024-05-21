using System.Globalization;

namespace ExchangeApp.Helper
{
    public class MailWithAttachmentObj
    {
        public string Subject { get; set; }
        public string Body { get; set; }

        public byte[] Attachments { get; set; }
        public List<string> Attendees { get; set; } = new List<string>();

        //public byte[] StringWithSeparatorToByteArray(string input, string separator)
        //{
        //    if (string.IsNullOrWhiteSpace(input))
        //    {
        //        return new byte[] { };
        //    }

        //    string[] stringParts = input.Split(new string[] { separator }, StringSplitOptions.None);
        //    List<byte> byteList = new List<byte>();

        //    foreach (string part in stringParts)
        //    {
        //        if (!string.IsNullOrWhiteSpace(part))
        //        {
        //            if (byte.TryParse(part, NumberStyles.HexNumber, null, out byte byteValue))
        //            {
        //                byteList.Add(byteValue);
        //            }
        //            else
        //            {
        //                throw new FormatException($"Invalid byte value: {part}");
        //            }
        //        }
        //    }

        //    return byteList.ToArray();
        //}


    }
}
