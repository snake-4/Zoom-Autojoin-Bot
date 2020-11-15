using FileHelpers;
using System;
using System.IO;
using System.Threading;

namespace YAZABNET
{
    class Program
    {
        static string GetVersionString()
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            System.Diagnostics.FileVersionInfo fvi = System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location);
            return fvi.FileVersion;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("Zoom autojoin bot by SnakePin, this time made with .NET!");
            Console.WriteLine("Version: " + GetVersionString());

            string csvFilePath;
            if (args.Length >= 1 && !string.IsNullOrWhiteSpace(args[0]))
            {
                csvFilePath = args[0];
            }
            else
            {
                Console.WriteLine("Enter CSV file path: ");
                csvFilePath = Console.ReadLine(); ;
            }

            if(!File.Exists(csvFilePath))
            {
                Console.WriteLine("Invalid file path specified!");
                return;
            }

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
                    //1 to the 7 starting from Monday
                    int day = (DateTime.Now.DayOfWeek == DayOfWeek.Sunday) ? 7 : (int)DateTime.Now.DayOfWeek;

                    if (schedule.DayOfWeek == day)
                    {
                        DateTime meetingStartDate = DateTime.Today + Utils.TimeSpanFrom24HString(schedule.TimeIn24H);
                        DateTime meetingEndDate = meetingStartDate + TimeSpan.FromSeconds(schedule.MeetingTimeInSeconds);
                        DateTime currentDate = DateTime.Now;

                        if ((currentDate > meetingStartDate) && (currentDate < meetingEndDate))
                        {
                            int secondsUntilMeetingEnds = (int)(meetingEndDate - currentDate).TotalSeconds;

                            Console.WriteLine("Joining a session.");
                            Console.WriteLine($"-> Comment: {schedule.Comment}");
                            Console.WriteLine($"-> ID: {schedule.MeetingID}");
                            Console.WriteLine($"-> Password: {schedule.MeetingPSW}");
                            Console.WriteLine($"-> MeetingTimeInSeconds: {schedule.MeetingTimeInSeconds}");

                            try
                            {
                                ZoomAutomationFunctions.ZoomJoinLeaveTask(schedule.MeetingID, schedule.MeetingPSW, secondsUntilMeetingEnds);
                                Console.WriteLine($"Session finished.");
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
        public int DayOfWeek;
        public string TimeIn24H;

        public string MeetingID;
        public string MeetingPSW;

        public int MeetingTimeInSeconds;
        public string Comment;
    }
}
