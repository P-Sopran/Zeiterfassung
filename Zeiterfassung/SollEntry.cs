using DocumentFormat.OpenXml.Spreadsheet;
using SQLite;
using System;
using System.Collections.Generic;

namespace Zeiterfassung
{
    public class SollEntry
    {
        [PrimaryKey]
        public string Date { get; set; }
        public string SollString { get; set; }
        public double SollHours { get; set; }
        public uint SortDate { get; set; }
        public string Weekday { get; set; }

        public SollEntry(TimeSpan ts, DateTime dt)
        {
            SollHours = ts.TotalHours;
            SollString = ts.ToString(@"hh\:mm");
            Date = dt.ToString("dd.MM.yy");
            switch ((int)dt.DayOfWeek)
            {
                case 1:
                    Weekday = "Mäntig";
                    break;
                case 2:
                    Weekday = "Zischtig";
                    break;
                case 3:
                    Weekday = "Mittwuch";
                    break;
                case 4:
                    Weekday = "Donschtig";
                    break;
                case 5:
                    Weekday = "Friitig";
                    break;
                case 6:
                    Weekday = "Samschtig";
                    break;
                default:
                    Weekday = "Sunntig";
                    break;
            }




            SortDate = Convert.ToUInt32(dt.ToString("yyMMdd")) * 10000;

        }

        public SollEntry()
        { // required for saving to the database
        }

        // used to create a string of format HH:MM without converting the hours to days when there are more than 24 hours
        private string TimeSpanString(double nr)
        {
            string prezero;
            int PlusMinus = 1;
            if (nr < 0)
            {
                PlusMinus = -1;
            }

            double rnr = Math.Round(nr, 2);

            double absnr = Math.Abs(rnr);
            double abshr = Math.Floor(absnr);
            double hr = abshr * PlusMinus;
            double min = Math.Round((absnr - abshr) * 60);

            if (min < 10)
            {
                prezero = "0";
            }
            else
            {
                prezero = "";
            }
            return hr + ":" + prezero + min;
        }



        // method fetches data from a soll entry object and creates an excel row. then inserts the data in the right places
        public void export(SheetData dest, uint row)
        {
            Row data = new Row() { RowIndex = row };
            Cell data1 = new Cell() { CellReference = String.Concat("A", row), CellValue = new CellValue(Date), DataType = CellValues.String };
            data.Append(data1);
            Cell data2 = new Cell() { CellReference = String.Concat("B", row), CellValue = new CellValue(SollString), DataType = CellValues.String };
            data.Append(data2);


            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                List<TimeEntry> reportTimeEntries = conn.Table<TimeEntry>().Where(s => s.Date == Date).OrderByDescending(t => t.TimeIn).ToList();
                double time = 0;
                double overtime = 0;

                // as NF are not counted agains SollTime these will be filtered out
                foreach (var timeItem in reportTimeEntries)
                {

                    if (timeItem.Type != "NF")
                    {
                        time += timeItem.hours;

                    }
                }

                overtime = time - SollHours;
                string overTime = TimeSpanString(overtime);
                string TimeSum = TimeSpanString(time);
                Cell sumTime = new Cell { CellReference = String.Concat("E", row), CellValue = new CellValue(TimeSum), DataType = CellValues.String };
                Cell OverTime = new Cell { CellReference = String.Concat("f", row), CellValue = new CellValue(overTime), DataType = CellValues.String };

                data.Append(sumTime, OverTime);
            }
            dest.Append(data);
        }

    }
}
