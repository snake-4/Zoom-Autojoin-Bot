using FileHelpers;
using System;
using System.Threading;

namespace YAZABNET
{
    class Program
    {
        static void Main(string[] args)
        {
            const string versionString = "0.0.1";

            Console.WriteLine("Zoom autojoin bot by SnakePin, this time made with .NET!");
            Console.WriteLine("Version: " + versionString);

            Console.WriteLine("Enter CSV file path: ");
            var csvFilePath = Console.ReadLine();

            var engine = new FileHelperEngine<ScheduleEntry>();
            engine.BeforeReadRecord += new FileHelpers.Events.BeforeReadHandler<ScheduleEntry>(
                (EngineBase _engine, FileHelpers.Events.BeforeReadEventArgs<ScheduleEntry> e) => e.SkipThisRecord = e.RecordLine.StartsWith("#")
            );
            var result = engine.ReadFile(csvFilePath);

            Console.WriteLine("Fully loaded and initialized.");

            while (true)
            {
                foreach (var schedule in result)
                {
                    //0 to the 6 starting from Monday
                    int day = ((DateTime.Now.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)DateTime.Now.DayOfWeek) - 1;

                    if ((schedule.DaysInteger & (1 << day)) != 0)
                    {
                        DateTime meetingStartDate = DateTime.Today + Utils.TimeSpanFrom24HString(schedule.TimeIn24H);
                        DateTime meetingEndDate = meetingStartDate + TimeSpan.FromSeconds(schedule.MeetingTimeInSeconds);
                        DateTime currentDate = DateTime.Now;

                        if ((currentDate > meetingStartDate) && (currentDate < meetingEndDate))
                        {
                            int secondsUntilMeetingEnds = (int)(meetingEndDate - currentDate).TotalSeconds;

                            Console.WriteLine($"Joining a session. ID: {schedule.MeetingID} Password: {schedule.MeetingPSW} MeetingTimeInSeconds: {schedule.MeetingTimeInSeconds}");
                            try
                            {
                                ZoomAutomationFunctions.ZoomJoinLeaveTask(schedule.MeetingID, schedule.MeetingPSW, secondsUntilMeetingEnds);
                            }
                            catch (Exception exc)
                            {
                                Console.WriteLine("Failed joining session. Sorry! Exception: " + exc);
                            }
                        }
                    }
                }
                Thread.Sleep(1000);
            }
        }
    }

    [DelimitedRecord(",")]
    [IgnoreFirst()]
    public class ScheduleEntry
    {
        public int DaysInteger;
        public string TimeIn24H;

        public string MeetingID;
        public string MeetingPSW;

        public int MeetingTimeInSeconds;

        [FieldValueDiscarded]
        public string Comment;
    }
}
