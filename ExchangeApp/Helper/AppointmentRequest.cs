namespace ExchangeApp.Helper
{
    public class AppointmentRequest
    {
        // test comment
        public string Subject { get; set; }
        public string Location { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public List<string> Attendees { get; set; }
    }
}
