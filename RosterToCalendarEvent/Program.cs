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
        public int NumShifts, NumPhoffS = 0;
        public double NumHours, NumPhoffH = 0;
        bool AddShift = true;
        bool PHOFF = false;
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

            Console.WriteLine("Thank you!\nNow please enter the path to save the CSV file:");
            string path = Console.ReadLine();

            // Set up CSV
            StringBuilder roster = new StringBuilder();

            // Write headings in CSV
            var newLine = string.Format("{0},{1},{2},{3},{4}", "Subject", "Start Date", "Start Time", "End Date", "End Time");
            roster.AppendLine(newLine);

            // Add shifts to CSV
            for (int i = 0; i < input.Count; i++)
            {
                // Get shift start and end time
                p.GetShiftDateTime(input[i]);

                if (p.AddShift) // Check if shift to add to CSV
                {
                    // Append line of CSV
                    newLine = string.Format("Work,{0}/{1}/{2},{3}:{4},{0}/{1}/{2},{5}:{6}", p.startTime.Day, p.startTime.Month, DateTime.Today.Year, p.startTime.Hour, p.startTime.Minute, p.endTime.Hour, p.endTime.Minute);
                    roster.AppendLine(newLine);
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

            // Write to CSV
            File.WriteAllText(path + "\\Roster.csv", roster.ToString());
                       
            Console.WriteLine("\nDone! Your file has been saved to " + path + "\\Roster.csv");
            Console.ReadLine();
        }
    }
}
