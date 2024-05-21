using ExchangeApp.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using System.Globalization;
using System.Runtime.Intrinsics.Arm;

namespace ExchangeApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MailController : ControllerBase
    {
        private readonly ExchangeSettings _exchangeSettings;

        public MailController(IOptions<ExchangeSettings> exchangeSettings)
        {
            _exchangeSettings = exchangeSettings.Value;

        }
        [HttpPost]
        [Route("CreateMailWithAttachmentInitial")]
        public async Task<IActionResult> CreateMailWithAttachmentInitial()
        {
            try
            {
                var formCollection = await Request.ReadFormAsync();

                // Read the JSON part
                var jsonPart = formCollection["json"];
                var mailWithAttachmentObj = JsonConvert.DeserializeObject<MailWithAttachmentObj>(jsonPart);

                // Read the file part
                IFormFile filePart = formCollection.Files["file"];
                byte[] fileBytes;

                using (var memoryStream = new MemoryStream())
                {
                    await filePart.CopyToAsync(memoryStream);
                    memoryStream.Flush();
                    fileBytes = memoryStream.ToArray();
                }

                // Initialize the Exchange service
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
                    Url = new Uri("https://mail.dpp.gov.ae/ews/exchange.asmx")
                };

                // Create the email message
                var email = new EmailMessage(service)
                {
                    Subject = mailWithAttachmentObj.Subject,
                    Body = new MessageBody(mailWithAttachmentObj.Body)
                };

                // Add recipients
                foreach (var attendeeEmail in mailWithAttachmentObj.Attendees)
                {
                    if (!string.IsNullOrWhiteSpace(attendeeEmail))
                    {
                        email.ToRecipients.Add(attendeeEmail.Trim());
                    }
                }

                // Add attachment
                string outputPath  = @"\\srv-dppapp02\e\ExportAttattchment\بطاقة_تقييم_مبدئي.pdf";
                email.Attachments.AddFileAttachment(outputPath);
                //email.Attachments.AddFileAttachment("aa.pdf", fileBytes);
                // Send the email
                email.SendAndSaveCopy();

                FileHelper.DeleteFile(outputPath);

                return Ok(new { message = "Email sent successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "Error sending email: " + ex.Message });
            }
        }
        [HttpPost]
        [Route("CreateMailWithAttachmentFinal")]
        public async Task<IActionResult> CreateMailWithAttachmentFinal()
        {
            try
            {
                var formCollection = await Request.ReadFormAsync();

                // Read the JSON part
                var jsonPart = formCollection["json"];
                var mailWithAttachmentObj = JsonConvert.DeserializeObject<MailWithAttachmentObj>(jsonPart);

                // Read the file part
                IFormFile filePart = formCollection.Files["file"];
                byte[] fileBytes;

                using (var memoryStream = new MemoryStream())
                {
                    await filePart.CopyToAsync(memoryStream);
                    memoryStream.Flush();
                    fileBytes = memoryStream.ToArray();
                }

                // Initialize the Exchange service
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
                    Url = new Uri("https://mail.dpp.gov.ae/ews/exchange.asmx")
                };

                // Create the email message
                var email = new EmailMessage(service)
                {
                    Subject = mailWithAttachmentObj.Subject,
                    Body = new MessageBody(mailWithAttachmentObj.Body)
                };

                // Add recipients
                foreach (var attendeeEmail in mailWithAttachmentObj.Attendees)
                {
                    if (!string.IsNullOrWhiteSpace(attendeeEmail))
                    {
                        email.ToRecipients.Add(attendeeEmail.Trim());
                    }
                }

                // Add attachment
                string outputPath = @"\\srv-dppapp02\e\ExportAttattchment\بطاقة_تقييم_نهائي.pdf";
                email.Attachments.AddFileAttachment(outputPath);
                //email.Attachments.AddFileAttachment("aa.pdf", fileBytes);
                // Send the email
                email.SendAndSaveCopy();

                FileHelper.DeleteFile(outputPath);

                return Ok(new { message = "Email sent successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "Error sending email: " + ex.Message });
            }
        }

        //public IActionResult CreateMailWithAttachment([FromBody] MailWithAttattchementobj request)
        //{
        //    try
        //    {
        //        // Initialize the Exchange service
        //        var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
        //        {
        //            Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
        //            Url = new Uri("https://mail.dpp.gov.ae/ews/exchange.asmx")
        //        };

        //        // Create the email message
        //        var email = new EmailMessage(service)
        //        {
        //            Subject = request.Subject,
        //            Body = new MessageBody(request.body)
        //        };

        //        // Add recipients
        //        foreach (var attendeeEmail in request.Attendees)
        //        {
        //            if (!string.IsNullOrWhiteSpace(attendeeEmail))
        //            {
        //                email.ToRecipients.Add(attendeeEmail.Trim());
        //            }
        //        }

        //        // Add attachment
        //        MailWithAttattchementobj requestatt= new MailWithAttattchementobj();
        //        byte[] ss ;
        //         ss = requestatt.StringWithSeparatorToByteArray(request.Attachments, ",");
        //            email.Attachments.AddFileAttachment("attachment.pdf", ss);


        //        string filePath = @"\\srv-dppapp02\e\PioneerExcellence\outputApi.pdf"; // Specify the full path to the PDF file

        //        // Write the byte array to a file
        //        using (FileStream fileStream = new FileStream(filePath, FileMode.Create))
        //        {
        //            fileStream.Write(ss, 0, ss.Length);
        //        }



        //        // Send the email
        //        email.SendAndSaveCopy();

        //        return Ok(new { message = "Email sent successfully." });
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(new { message = "Error sending email: " + ex.Message });
        //    }
        //}
    }
}
