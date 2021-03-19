using SQLite;
using System;

namespace Zeiterfassung
{
    class CheckinEntry
    {
        [PrimaryKey]
        private int ID { get; set; }
        public double hrs { get; set; }
        public string comment { get; set; }
        public string type { get; set; }



        public CheckinEntry()
        {
            ID = 1;
            hrs = -1;
        }

        public CheckinEntry(TimeSpan ts)
        {
            ID = 1;
            hrs = ts.TotalHours;
        }


    }





}
