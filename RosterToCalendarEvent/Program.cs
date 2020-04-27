using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Globalization;
using Microsoft.Win32;

namespace RosterToCalendarEvent
{
    class Program
    {

        public int NumShifts = 0;
        public double NumHours = 0;
        public string DayDate, Month, StartHour, StartMinute, StartPrefix, FinishHour, FinishMinute;

        public void GetMonthNum(string month)
        {
            // Create array with months
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

            int index = Array.IndexOf(months, month) + 1;

            if (index < 10)
            {
                Month = ("0" + index.ToString());
            }
            else
            {
                Month = index.ToString();
            }
        } 
        public void GetShiftDateTime(string shift)
        {
            // Pass each line from message through to another statement and break up into numbers
            string[] numbers = Regex.Split(shift, @"\D+");

            // Assign to variables
            // Day date
            if (int.Parse(numbers[1]) < 10)
            {
                DayDate = ("0" + numbers[1]);
            }
            else
            {
                DayDate = numbers[1];
            }

            //Start hour and prefix
            if (int.Parse(numbers[2]) > 12)
            {
                StartHour = "0" + (int.Parse(numbers[2]) - 12).ToString();
                StartPrefix = "PM";
            }
            else if (int.Parse(numbers[2]) == 12)
            {
                StartHour = numbers[2];
                StartPrefix = "PM";
            }
            else
            {
                if (int.Parse(numbers[2]) < 10)
                {
                    StartHour = (/*"0" + */numbers[2]);
                }
                else
                {
                    StartHour = numbers[2];
                }
                StartPrefix = "AM";
            }

            // Start minute
            if (int.Parse(numbers[3]) == 0)
            {
                StartMinute = "00";
            }
            else
            {
                StartMinute = numbers[3];
            }

            // Finish hour
            FinishHour = "0" + (int.Parse(numbers[4]) - 12).ToString();

            // Finish minute
            if (int.Parse(numbers[5]) == 0)
            {
                FinishMinute = "00";
            }
            else
            {
                FinishMinute = numbers[5];
            }
        }

        public void CalculateHours(string startHour, string startMinute, string finishHour, string finishMinute, string day, string month, string prefix)
        {
            string adjustedHour;
            if (prefix == "AM")
            {
                adjustedHour = startHour;
            }
            else
            {
                if (Int32.Parse(startHour) == 12)
                {
                    adjustedHour = startHour;
                }
                else
                {
                    adjustedHour = (Int32.Parse(startHour) + 12).ToString();
                }
            }

            finishHour = (Int32.Parse(finishHour) + 12).ToString();

            // Start time
            String strDate = string.Format("{0}/{1}/{2} {3}:{4}:00", day, month, DateTime.Now.Year, adjustedHour, startMinute);
            DateTime StartDT = DateTime.ParseExact(strDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            // End time
            strDate = string.Format("{0}/{1}/{2} {3}:{4}:00", day, month, DateTime.Now.Year, finishHour, finishMinute);
            DateTime endDT = DateTime.ParseExact(strDate, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            // Add hours
            NumHours += ((endDT - StartDT).TotalMinutes) / 60;
        }


        static void Main(string[] args)
        {
            
            // Set up public variables
            Program p = new Program();

            // If want to use txt file:
            //string[] input = File.ReadAllLines("D:\\Desktop\\Shifts.txt");

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

            // Calculate number of shifts 
            p.NumShifts = input.Count() - 2;

            // Set up CSV
            StringBuilder roster = new StringBuilder();

            // Write headings in CSV
            var newLine = string.Format("{0},{1},{2},{3},{4}", "Subject", "Start Date", "Start Time", "End Date", "End Time");
            roster.AppendLine(newLine);

            // Add shifts to CSV
            for (int i = 1; i < (p.NumShifts + 1); i++)
            {
                p.GetShiftDateTime(input[i]);

                // get month
                string[] words = input[i].Split(' ');
                p.GetMonthNum(words[2]);

                // Append line of CSV
                newLine = string.Format("Work,{0}/{1}/{2},{3}:{4} {5},{0}/{1}/{2},{6}:{7} PM", p.DayDate, p.Month, DateTime.Today.Year, p.StartHour, p.StartMinute, p.StartPrefix, p.FinishHour, p.FinishMinute);
                roster.AppendLine(newLine);

                // Calculate hours
                p.CalculateHours(p.StartHour, p.StartMinute, p.FinishHour, p.FinishMinute, p.DayDate, p.Month, p.StartPrefix);
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
