using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;
using Microsoft.Exchange.WebServices.Data;
using System.Text.RegularExpressions;

namespace RoomBook
{   
    internal class iCalWrapper
    {
        #region Private Members

        private IICalendar miCalObject;

        #endregion

        /// <summary>
        /// Initialise a new wrapper object, with a blank iCal.
        /// </summary>
        public iCalWrapper()
        {
            miCalObject = new iCalendar();
        }

        /// <summary>
        /// Copy an Exchange appointment into an iCal event.
        /// </summary>
        /// <param name="ExchangeAppointment">Exchange appointment object to copy</param>
        public void AddEventToCalendar(Appointment ExchangeAppointment)
        {
            // Create a new event
            Event iCalEvent = miCalObject.Create<Event>();

            // Set the event properties
            iCalEvent.Summary = ExchangeAppointment.Subject.Replace("–", "-");      //Note: the replace replaces Word's "long dash"
            iCalEvent.Description = ExchangeAppointment.Subject.Replace("–", "-");  //      with a normal dash
            iCalEvent.Start = new iCalDateTime(ExchangeAppointment.Start);
            iCalEvent.End = new iCalDateTime(ExchangeAppointment.End);
            iCalEvent.Duration = ExchangeAppointment.Duration;
            iCalEvent.Location = ExchangeAppointment.Location;

            // TODO : Set Timezone

            // Output details for debug purposes
            //Console.WriteLine(string.Format("Subject: {0}\n Start Date: {1}\tEnd Date: {2}\nDuration: {3}\nLocation: {4}\n\n",
            //                ExchangeAppointment.Subject, ExchangeAppointment.Start, ExchangeAppointment.End, ExchangeAppointment.Duration,
            //                ExchangeAppointment.Location ?? "Not Set"));

            //Console.WriteLine();

        }

        /// <summary>
        /// Add a list of Exchange appointment objects to an iCal object, whilst applying a set of filters.
        /// </summary>
        /// <param name="ExchangeAppointments">List of Exchange Appointment objects to add</param>
        /// <param name="FilterOptions">Filter settings to apply</param>
        public void AddEventsToCalendar(IEnumerable<Appointment> ExchangeAppointments, FilterSetting FilterOptions)
        {
            Regex SubjectRegEx = null;
            Regex LocationRegEx = null;

            // If the subject filter has been set, initialise the subject regex
            if (FilterOptions.SubjectFilterSet)
            {
                SubjectRegEx = new Regex(FilterOptions.SubjectFilter);
            }

            // If the location filter has been set, initialise the location regex
            if (FilterOptions.LocationFilterSet)
            {
                LocationRegEx = new Regex(FilterOptions.LocationFilter);
            }

            // Iterate over the list of Exchange Appointment objects...
            foreach (Appointment ExchangeAppointment in ExchangeAppointments)
            {
                // If there is a subject filter...
                if (FilterOptions.SubjectFilterSet)
                {
                    // ... Check this appointment's subject matches it - if not ...
                    if (SubjectRegEx.IsMatch(ExchangeAppointment.Subject) == false)
                    {
                        // ... move on to the next appointment ...
                        continue;
                    }
                }

                // If there is a location filter...
                if (FilterOptions.LocationFilterSet)
                {
                    // ... Check this appointment's location matches it - if not ...
                    if (LocationRegEx.IsMatch(ExchangeAppointment.Location) == false)
                    {
                        // ... move on to the next appointment ...
                        continue;
                    }
                }

                // If there's a category filter ...
                if (FilterOptions.CategoryFilterFilterSet)
                {
                    bool FilterCategoryFound = false;

                    // Iterate over the categories to filter by ...
                    foreach (string CategoryFilter in FilterOptions.CategoryFilter)
                    {
                        // ... if one of the appointment's categories matches ...
                        if (ExchangeAppointment.Categories.Contains(CategoryFilter))
                        {
                            // ... mark this appointment as includable, and stop checking
                            FilterCategoryFound = true;
                            break;
                        }                        
                    }

                    // If we didn't find any category matches, move on to the next appointment ...
                    if (FilterCategoryFound == false)
                    {
                        continue;
                    }
                }

                // If we got here, all of the filter conditions passed, so add the appointment to the iCal object
                AddEventToCalendar(ExchangeAppointment);
            }
        }

        /// <summary>
        /// Serialises the underlying iCal object to a string.
        /// </summary>
        /// <returns>String representing the iCal object.</returns>
        public string GenerateString()
        {
            iCalendarSerializer Serializer = new iCalendarSerializer();

            return Serializer.SerializeToString(miCalObject);
        }

        /// <summary>
        /// Serialises the underlying iCal object to a file.
        /// </summary>
        /// <param name="Filename">Path of the file to write to.</param>
        public void GenerateFile(string Filename)
        { 
            iCalendarSerializer Serializer = new iCalendarSerializer();

            Serializer.Serialize(miCalObject, Filename);
        }
    }
}
