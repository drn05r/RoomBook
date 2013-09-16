using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RoomBook
{
    /// <summary>
    /// Class holding the set of selected filter options.
    /// </summary>
    internal class FilterSetting
    {
        /// <summary>
        /// RegEx pattern to apply as filter to the subject field.
        /// </summary>
        public string SubjectFilter
        {
            get;
            set;
        }

        /// <summary>
        /// RegEx pattern to apply as filter to the location field.
        /// </summary>
        public string LocationFilter
        {
            get;
            set;
        }

        /// <summary>
        /// Set of strings to match one or more of in the set of categories
        /// </summary>
        public IEnumerable<string> CategoryFilter
        {
            get;
            set;
        }

        /// <summary>
        /// Date of first appointment to find.
        /// </summary>
        public DateTime SearchStartDate
        {
            get;
            set;
        }

        /// <summary>
        /// Date of last appointment to find.
        /// </summary>
        public DateTime SearchEndDate
        {
            get;
            set;
        }

        /// <summary>
        /// Was the subject filter set?
        /// </summary>
        public bool SubjectFilterSet
        {
            get
            {
                return !String.IsNullOrEmpty(SubjectFilter);
            }
        }

        /// <summary>
        /// Was the location filter set?
        /// </summary>
        public bool LocationFilterSet
        {
            get
            {
                return !String.IsNullOrEmpty(LocationFilter);
            }
        }

        /// <summary>
        /// Was the category filter set?
        /// </summary>
        public bool CategoryFilterFilterSet
        {
            get
            {
                return (CategoryFilter != null);
            }
        }
    }
}
