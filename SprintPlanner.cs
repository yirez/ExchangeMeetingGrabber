
namespace SprintPlanner
{
    public class SprintManager
    {

        private DateTime firstOccurrenceOfMeeting = new DateTime(2000, 1, 1);  
        private const string MEETING_ROOM_CONST = "roomMail@company.com"; 
        private const string MEETING_ROOM_CONST2 = "room2Mail@company.com"; 
        private const string attendee1 = "attendee@company.com";
        private const string attendee2 = "attendee2@company.com";


        private const int MAX_DAYS_ALLOWED = 15; //whatever max time allowed(in days)
        private readonly DateTime maxDateReach = DateTime.Now.Date.AddDays(MAX_DAYS_ALLOWED);
        private ExchangeService service;
        public void ManageSpringMeetings()
        { 
            //is the farthest reachable date a 2 weeks distance from the last meeting?
            if ((int)((maxDateReach - firstProjectMeeting).TotalDays % 14) == 0)
            {
                planMeeting("Meeting Subject", MeetingRoomEnum.SomeRoom, new TimeSpan(8, 0, 0), new TimeSpan(12, 0, 0), new List<string>() { attendee1,attendee2 }); 
            }   
        }
 
        private void planMeeting(string subject, MeetingRoom room, TimeSpan startTime, TimeSpan endTime, List<string> attendees)
        { 
            prepareEWSService(); 
            Appointment meeting = prepareMeeting(subject); 
            prepareMeetingTimes(ref meeting, startTime, endTime); 
            prepareAttendees(ref meeting, room, attendees); 
            saveMeeting(ref meeting); 
        }

        #region Preparation Methods

        private void prepareEWSService()
        { 
            if (service == null)
            {
                service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                service.Credentials = new WebCredentials("yourmail@company.com", "yourpass");
                service.AutodiscoverUrl("yourmail@company.com", RedirectionUrlValidationCallback);
            } 
        }

        private bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }


        private Appointment prepareMeeting(string subject)
        { 
            Appointment meeting = new Appointment(service);
            meeting.Subject = subject;
            meeting.ReminderMinutesBeforeStart = 60; 
            return meeting;
        }
        private void prepareAttendees(ref Appointment meeting, MeetingRoom room, List<string> attendees)
        {
            addRoomsToMeeting(ref meeting, room);
            addAttendeesToMeeting(ref meeting, attendees);
        }
        private void saveMeeting(ref Appointment meeting)
        {
            meeting.Save(SendInvitationsMode.SendToAllAndSaveCopy);
            Item item = Item.Bind(service, meeting.Id, new PropertySet(ItemSchema.Subject));
        }
        private void addAttendeesToMeeting(ref Appointment meeting, List<string> attendees)
        {
            if (attendees.Any())
            {
                foreach (string attendee in attendees)
                {
                    meeting.RequiredAttendees.Add(attendee);
                }
            }
        }
        private void addRoomsToMeeting(ref Appointment meeting, MeetingRoom room)
        {
            if (room == MeetingRoomEnum.SomeRoom)
            {
                meeting.RequiredAttendees.Add(MEETING_ROOM_CONST);
                meeting.Location = "whatever location text you want";
            }
            else if (room == MeetingRoom.OtherRoom)
            {
                meeting.RequiredAttendees.Add(MEETING_ROOM_CONST2);
                meeting.Location = "other location text you want";
            } 
            else
            {
                throw new Exception("No rooms set");
            }
        }
        #endregion 

        #region methods to prepare meeting times.
        private void prepareMeetingTimes(ref Appointment meeting, TimeSpan startTime, TimeSpan endTime)
        {

            meeting.Start = new DateTime(maxDateReach.Year, maxDateReach.Month, maxDateReach.Day, startTime.Hours, startTime.Minutes, 0);
            meeting.End = new DateTime(maxDateReach.Year, maxDateReach.Month, maxDateReach.Day, endTime.Hours, endTime.Minutes, 0);
        }
        #endregion 
    }
}
