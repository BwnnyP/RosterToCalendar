using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace RosterToCalendarEvent
{
    class Program
    {
        public int NumShifts, NumPhoffS = 0;
        public double NumHours, NumPhoffH = 0;
        bool AddShift = true;
        bool PHOFF = false;
        public DateTime startTime, endTime;

        // Required for Google integration
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        static string[] Scopes = { CalendarService.Scope.Calendar };
        static string ApplicationName = "Google Calendar API .NET Quickstart";

        // Create Calendar Event in Google
        public void ConnectToAPI(Event ev, string calendarId)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                // Console.WriteLine("Credential file saved to: " + credPath);      // Uneeded
            }

            // Create Google Calendar API service.
            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define parameters of request.
            EventsResource.ListRequest request = service.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // Create Event
            service.Events.Insert(ev, calendarId).Execute();
        }

        public int GetMonthNum(string month)
        {
            // Create array with months
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

            int index = Array.IndexOf(months, month) + 1;

            return index;
        } 
        public void GetShiftDateTime(string shift)
        {
            // Split into words
            string[] words = shift.Split(' ');

            if (words.Length < 3 | words[0] == "Access")
            {
                AddShift = false;
            }
            else
            {
                // Pass each line from message through to another statement and break up into numbers
                string[] numbers = Regex.Split(shift, @"\D+");

                if (words[3] == "PHOFF")
                {
                    NumPhoffS++;
                    NumPhoffH += (Int32.Parse(numbers[2]) + (double.Parse(numbers[3]) / 100));
                    PHOFF = true;
                }
                else
                {
                    // Generate Shift
                    startTime = new DateTime(DateTime.Now.Year, GetMonthNum(words[2]), Int32.Parse(numbers[1]), Int32.Parse(numbers[2]), Int32.Parse(numbers[3]), 0);
                    endTime = new DateTime(DateTime.Now.Year, GetMonthNum(words[2]), Int32.Parse(numbers[1]), Int32.Parse(numbers[4]), Int32.Parse(numbers[5]), 0);

                    // Add hours
                    NumHours += ((endTime - startTime).TotalMinutes) / 60;
                    NumShifts++;
                }
            }
        }

        static void Main(string[] args)
        {
            // Set up public variables
            Program p = new Program();

            // Shifts input
            Console.WriteLine("Please paste the text you received below and press enter twice!");

            List<string> input = new List<String>();

            string line;
            while ((line = Console.ReadLine()) != null && line != "")
            {
                input.Add(line);
            }

            Console.WriteLine("Thank you! Adding to your calendar now!");

            /*
            Console.WriteLine("Thank you!\nNow please enter the path to save the CSV file:");
            string path = Console.ReadLine();

            // Set up CSV
            StringBuilder roster = new StringBuilder();

            // Write headings in CSV
            var newLine = string.Format("{0},{1},{2},{3},{4}", "Subject", "Start Date", "Start Time", "End Date", "End Time");
            roster.AppendLine(newLine);
            */

            // Add shifts to Google calendar
            for (int i = 0; i < input.Count; i++)
            {
                // Get shift start and end time
                p.GetShiftDateTime(input[i]);

                if (p.AddShift) // Check if shift to add to Google calendar
                {
                    /*
                    // Append line of CSV
                    newLine = string.Format("Work,{0}/{1}/{2},{3}:{4},{0}/{1}/{2},{5}:{6}", p.startTime.Day, p.startTime.Month, DateTime.Today.Year, p.startTime.Hour, p.startTime.Minute, p.endTime.Hour, p.endTime.Minute);
                    roster.AppendLine(newLine);
                    */

                    // Add to Google
                    var ev = new Event();
                    EventDateTime start = new EventDateTime();
                    start.DateTime = new DateTime(DateTime.Now.Year, p.startTime.Month, p.startTime.Day, p.startTime.Hour, p.startTime.Minute, 0);

                    EventDateTime end = new EventDateTime();
                    end.DateTime = new DateTime(DateTime.Now.Year, p.endTime.Month, p.endTime.Day, p.endTime.Hour, p.endTime.Minute, 0);

                    ev.Start = start;
                    ev.End = end;

                    ev.Summary = "Work";
                    string calendarId = "primary";

                    p.ConnectToAPI(ev, calendarId);
                }
                p.AddShift = true; // Reset back to add shift
            }

            // Number of shifts and hours output to console
            Console.WriteLine("\nNumber of shifts this month: " + p.NumShifts);
            Console.WriteLine("Number of hours this month: " + p.NumHours);

            if (p.PHOFF)
            {
                Console.WriteLine("Number of PHOFF shifts this month: " + p.NumPhoffS);
                Console.WriteLine("Number of PHOFF hours this month: " + p.NumPhoffH);
            }

            /*
            // Write to CSV
            File.WriteAllText(path + "\\Roster.csv", roster.ToString());
            */

            Console.WriteLine("\nDone! Your shifts have been saved to you Google calendar!");
            Console.ReadLine();
        }
    }
}
