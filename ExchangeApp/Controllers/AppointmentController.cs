using ExchangeApp.Helper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Options;

namespace ExchangeApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AppointmentController : ControllerBase
    {
        private readonly ExchangeSettings _exchangeSettings;

        public AppointmentController(IOptions<ExchangeSettings> exchangeSettings)
        {
            _exchangeSettings = exchangeSettings.Value;
        }

        [HttpPost]
        public IActionResult CreateAppointment([FromBody] AppointmentRequest request)
        {
            try
            {
                var service = new ExchangeService(ExchangeVersion.Exchange2010_SP2)
                {
                    Credentials = new WebCredentials(_exchangeSettings.UserName, _exchangeSettings.Password),
                    Url = new Uri("https://mail.dpp.gov.ae/ews/exchange.asmx")
                };

                var appointment = new Appointment(service)
                {
                    Subject = request.Subject,
                    Location = request.Location,
                    Start = request.StartTime,
                    End = request.EndTime,
                    Importance = Importance.High
                };



                foreach (var attendeeEmail in request.Attendees)
                {
                    if (!string.IsNullOrWhiteSpace(attendeeEmail)) 
                    {
                        appointment.RequiredAttendees.Add(new Attendee(attendeeEmail.Trim()));
                    }
                }



                appointment.Save(SendInvitationsMode.SendToAllAndSaveCopy);

                return Ok(new { message = "Appointment created successfully." });
            }
            catch (Exception ex)
            {
                return BadRequest(new { message = "Error creating appointment: " + ex.Message });
            }
        }
    }
}
