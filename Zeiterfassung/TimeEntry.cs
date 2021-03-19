using DocumentFormat.OpenXml.Spreadsheet;
using SQLite;
using System;
using System.Diagnostics;

namespace Zeiterfassung
{
    public class TimeEntry
    {
        [PrimaryKey, AutoIncrement]
        public int ID { get; set; }
        public string Date { get; set; }
        public string TimeIn { get; set; }
        public string TimeOut { get; set; }
        public string Time { get; set; }
        public string Type { get; set; }
        public string Comment { get; set; }
        public double hours { get; set; }
        public uint SortDate { get; set; }

        public TimeEntry(DateTime dt, TimeSpan timeIn, TimeSpan timeOut, string comment)
        {
            Date = dt.ToString("dd.MM.yy");
            TimeIn = timeIn.ToString(@"hh\:mm");
            TimeOut = timeOut.ToString(@"hh\:mm");
            TimeSpan ts;
            if (timeIn.CompareTo(timeOut) == 1)
            {
                TimeSpan tfhours = TimeSpan.FromHours(24);
                ts = timeOut.Add(tfhours).Subtract(timeIn);
            }
            else
            {
                ts = timeOut.Subtract(timeIn);
            }
            Time = ts.ToString(@"hh\:mm");
            Comment = comment;
            hours = ts.TotalHours;
            SortDate = Convert.ToUInt32(dt.ToString("yyMMdd") + timeIn.ToString("hhmm"));
        }

        public TimeEntry()
        {
            // required for saving to the database
        }
        // method fetches data from a time entry object and creates an excel row. then inserts the data in the right places
        public void export(SheetData dest, uint row)
        {
            Row data = new Row() { RowIndex = row };

            if (Type == "NF" || Type == "ÜZ")
            {
                Cell data1 = new Cell() { CellReference = String.Concat("A", row), CellValue = new CellValue(Date), DataType = CellValues.String };
                data.Append(data1);
            }
            Cell data2 = new Cell() { CellReference = String.Concat("C", row), CellValue = new CellValue(TimeIn), DataType = CellValues.String };
            data.Append(data2);
            Cell data3 = new Cell() { CellReference = String.Concat("D", row), CellValue = new CellValue(TimeOut), DataType = CellValues.String };
            data.Append(data3);
            Cell data4 = new Cell() { CellReference = String.Concat("E", row), CellValue = new CellValue(Time), DataType = CellValues.String };
            data.Append(data4);
            Cell data5 = new Cell() { CellReference = String.Concat("F", row), CellValue = new CellValue(Type), DataType = CellValues.String };
            data.Append(data5);
            Cell data6 = new Cell() { CellReference = String.Concat("G", row), CellValue = new CellValue(Comment), DataType = CellValues.String };
            data.Append(data6);

            dest.Append(data);
        }







    }



}
