using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;

namespace RosterToCalendarEvent
{
    class Program
    {
        public int NumShifts = 0;
        public double NumHours = 0;
        public DateTime startTime, endTime;

        public int GetMonthNum(string month)
        {
            // Create array with months
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

            int index = Array.IndexOf(months, month) + 1;

            return index;
        } 
        public void GetShiftDateTime(string shift)
        {
            // Pass each line from message through to another statement and break up into numbers
            string[] numbers = Regex.Split(shift, @"\D+");
    
            // Split into words
            string[] words = shift.Split(' ');

            // Generate Shift
            startTime = new DateTime(DateTime.Now.Year, GetMonthNum(words[2]), Int32.Parse(numbers[1]), Int32.Parse(numbers[2]), Int32.Parse(numbers[3]), 0);
            endTime = new DateTime(DateTime.Now.Year, GetMonthNum(words[2]), Int32.Parse(numbers[1]), Int32.Parse(numbers[4]), Int32.Parse(numbers[5]), 0);

            // Add hours
            NumHours += ((endTime - startTime).TotalMinutes) / 60;

            // Add number of shifts once can recognise input string
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

            Console.WriteLine("Thank you! \n\nNow please enter the path to save the CSV file:");
            string path = Console.ReadLine();

            // Calculate number of shifts           NEED TO CHANGE
            p.NumShifts = input.Count() - 2;

            // Set up CSV
            StringBuilder roster = new StringBuilder();

            // Write headings in CSV
            var newLine = string.Format("{0},{1},{2},{3},{4}", "Subject", "Start Date", "Start Time", "End Date", "End Time");
            roster.AppendLine(newLine);

            // Add shifts to CSV
            for (int i = 1; i < (p.NumShifts + 1); i++)
            {
                // Get shift start and end time
                p.GetShiftDateTime(input[i]);

                /*
                // get month
                string[] words = input[i].Split(' ');
                p.GetMonthNum(words[2]);
                */

                // Append line of CSV
                newLine = string.Format("Ben Work,{0}/{1}/{2},{3}:{4},{0}/{1}/{2},{5}:{6}", p.startTime.Day, p.startTime.Month, DateTime.Today.Year, p.startTime.Hour, p.startTime.Minute, p.endTime.Hour, p.endTime.Minute);
                roster.AppendLine(newLine);
            }

            // Number of shifts and hours output to console
            Console.WriteLine("\nNumber of shifts this month: " + p.NumShifts);
            Console.WriteLine("Number of hours this month: " + p.NumHours);

            // Write to CSV
            File.WriteAllText(path + "\\Roster.csv", roster.ToString());
                       
            Console.WriteLine("\nDone! Your file has been saved to " + path + "\\Roster.csv");
            Console.ReadLine();
        }
    }
}
