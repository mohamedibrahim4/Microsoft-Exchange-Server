using ExchangeApp.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using System.Globalization;
using System.Runtime.Intrinsics.Arm;
using System.Runtime.Serialization;

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

                // Initialize the Exchange service
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
                    Url = new Uri("Echange Server URL")
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
                var fileStream = filePart.OpenReadStream();

                email.Attachments.AddFileAttachment("Initial Report.pdf", fileStream);
                await fileStream.FlushAsync();
                // Send the email
                email.SendAndSaveCopy();

                //FileHelper.DeleteFile(outputPath);

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

                // Initialize the Exchange service
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
                    Url = new Uri("Echange Server URL")
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
                var fileStream = filePart.OpenReadStream();

                email.Attachments.AddFileAttachment("FinalReport.pdf", fileStream);
                await fileStream.FlushAsync();
                // Send the email
                email.SendAndSaveCopy();

                //FileHelper.DeleteFile(outputPath);

                return Ok(new { message = "Email sent successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "Error sending email: " + ex.Message });
            }
        }

    }
}
