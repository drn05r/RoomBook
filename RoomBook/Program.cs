using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace RoomBook
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 5)
            {
                Console.WriteLine("This application must be passed 4 arguments:");
                Console.WriteLine("  <username> - The username of the user requesting a calendar. (E.g. abc1d23).");
                Console.WriteLine("  <password> - The pass word of the user requesting a calendar.");
                Console.WriteLine("  <calendar> - The name of the calendar. (E.g. building 32 room 4049 would be b32r4049).");
                Console.WriteLine("  <starttime> - Unix time when booking should start.");
                Console.WriteLine("  <endtime> - Unix time when booking should end.");
                Environment.Exit(1);
            }
            // Connecting to Exchange
            Console.WriteLine("Connecting to Exchange...");
            ExchangeService WebService = ExchangeWebServiceWrapper.GetWebServiceInstance(args[0], args[1], "SOTON");

            // Getting calendar
            Console.WriteLine("Getting calendar for " + args[2] + "...");
            Mailbox mailbox = new Mailbox(args[2] + "@soton.ac.uk");
            FolderId calendarFolderId = new FolderId(WellKnownFolderName.Calendar, mailbox);
            CalendarFolder Calendar = CalendarFolder.Bind(WebService, calendarFolderId);
            if (Calendar == null)
            {
                throw new Exception("Could not find calendar folder.");
            }

            // Checking for exitsing appointments
            Console.WriteLine("Checking for existing appointments...");
            DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
            DateTime startTime = epoch.AddSeconds(Convert.ToDouble(args[3]));
            if (startTime.Minute > 30)
                startTime = new DateTime(startTime.Year, startTime.Month, startTime.Day, startTime.Hour, 30, 0);
            else
                startTime = new DateTime(startTime.Year, startTime.Month, startTime.Day, startTime.Hour, 0, 0);

            epoch = new DateTime(1970, 1, 1, 0, 0, 0, 0).ToLocalTime();
            DateTime endTime = epoch.AddSeconds(Convert.ToDouble(args[4]));
            if (endTime.Minute > 30)
                endTime = new DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, 30, 0);
            else
                endTime = new DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, 0, 0);
            FilterSetting FilterOptions = new FilterSetting();
            FilterOptions.SearchStartDate = startTime.AddSeconds(1);
            FilterOptions.SearchEndDate = endTime.AddSeconds(-1);
            FindItemsResults<Appointment> Appointments = Calendar.FindAppointments(new CalendarView(FilterOptions.SearchStartDate, FilterOptions.SearchEndDate));
            if (Appointments.TotalCount > 0)
            {
                Console.WriteLine("This room is already booked " + Appointments.TotalCount + " time(s) between " + startTime + " to " + endTime);
                Array AppArr = Appointments.ToArray();
                foreach (Appointment App in AppArr)
                    Console.WriteLine("Booking from " + App.Start.ToString() + " to " + App.End.ToString());
                Environment.Exit(2);
            }
            else
                Console.WriteLine("This room can be booked");           
                     
            // Adding appointment
            Console.WriteLine("Booking " + args[2] + " from " + startTime.ToString() + " to " + endTime.ToString());
            var appointment = new Appointment(WebService)
            {
                Subject = "RoomBook app (immediate booking)",
                Body = "This is an immediate booking from the RoomBook App ",
                Start = startTime,
                End = endTime
            };
            appointment.Save(calendarFolderId, SendInvitationsMode.SendToNone);
            Console.WriteLine("Room is now booked");
        }
    }
}
