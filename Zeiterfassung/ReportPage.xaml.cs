
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using Xamarin.Essentials;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;

namespace Zeiterfassung
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class ReportPage : ContentPage
    {
        private List<TimeEntry> timeEntries = new List<TimeEntry>();
        private List<SollEntry> sollEntries = new List<SollEntry>();
        private uint startDate;
        private uint endDate;
        private string path => Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments).ToString();

        public ReportPage()
        {
            startDate = 1;
            endDate = Convert.ToUInt32(DateTime.Today.ToString("yyMMdd")) * 10000 + 9999;
            InitializeComponent();
        }
        // calculates difference between time worked and soll time everytime the page appears
        protected override void OnAppearing()
        {
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<SollEntry>();
                sollEntries = conn.Table<SollEntry>().Where(s => s.SortDate <= endDate && s.SortDate >= startDate).ToList();
                conn.CreateTable<TimeEntry>();
                timeEntries = conn.Table<TimeEntry>().Where(s => s.SortDate <= endDate && s.SortDate >= startDate).ToList();
            }

            double sumSoll = 0;
            foreach (var item in sollEntries)
            {
                sumSoll += item.SollHours;
            }

            double sumActual = 0;
            double sumNotfall = 0;
            foreach (var item in timeEntries)
            {
                if (item.Type == "NF")
                {
                    sumNotfall += item.hours;
                }
                else
                {
                    sumActual += item.hours;
                }
            }
            double sumDifference = sumActual - sumSoll;
            string Actual = TimeSpanString(sumActual);
            string Soll = TimeSpanString(sumSoll);
            string Difference = TimeSpanString(sumDifference);
            string Notfall = TimeSpanString(sumNotfall);

            Worked.Text = Actual;
            SollTime.Text = Soll;
            Diff.Text = Difference;
            Emergency.Text = Notfall;
            base.OnAppearing();
        }


        public void calculate(object sender, EventArgs e)
        {
            startDate = Convert.ToUInt32(StartDatePicked.Date.ToString("yyMMdd")) * 10000;
            endDate = Convert.ToUInt32(EndDatePicked.Date.ToString("yyMMdd")) * 10000 + 9999;
            OnAppearing();
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

        //thanks to https://stackoverflow.com/questions/44572018/using-open-xml-to-create-excel-file
        // and thanks to https://askxammy.com/getting-started-with-excel-files-creation-in-xamarin-forms/

        // creates an excel file with the data from selected time entries and soll entries
        public void report(object sender, EventArgs e)
        {
            calculate(null, null);
            string filepath = path + "/report.xlsx";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook))
            {
                //preparation of workbook
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.
                    GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "report"
                };


                //Create Report Summary
                Row Title = new Row { RowIndex = 1 };
                string sDate = StartDatePicked.Date.ToString("dd.MM.yy");
                string eDate = EndDatePicked.Date.ToString("dd.MM.yy");
                Cell titlecell = new Cell() { CellReference = "A1", CellValue = new CellValue("Zeiterfassung Caroline Koch von " + sDate + " bis " + eDate), DataType = CellValues.String };
                Title.Append(titlecell);

                Row HeaderTop = new Row() { RowIndex = 2 };
                Cell header1 = new Cell() { CellReference = "A2", CellValue = new CellValue("Total gearbeitet"), DataType = CellValues.String };
                HeaderTop.Append(header1);
                Cell header2 = new Cell() { CellReference = "B2", CellValue = new CellValue("Total Sollzeit"), DataType = CellValues.String };
                HeaderTop.Append(header2);
                Cell header3 = new Cell() { CellReference = "C2", CellValue = new CellValue("Überzeit"), DataType = CellValues.String };
                HeaderTop.Append(header3);
                Cell header4 = new Cell() { CellReference = "D2", CellValue = new CellValue("Total Notfälle"), DataType = CellValues.String };
                HeaderTop.Append(header4);

                Row ContentTop = new Row() { RowIndex = 3 };
                Cell content1 = new Cell() { CellReference = "A3", CellValue = new CellValue(Worked.Text), DataType = CellValues.String };
                ContentTop.Append(content1);
                Cell content2 = new Cell() { CellReference = "B3", CellValue = new CellValue(SollTime.Text), DataType = CellValues.String };
                ContentTop.Append(content2);
                Cell content3 = new Cell() { CellReference = "C3", CellValue = new CellValue(Diff.Text), DataType = CellValues.String };
                ContentTop.Append(content3);
                Cell content4 = new Cell() { CellReference = "D3", CellValue = new CellValue(Emergency.Text), DataType = CellValues.String };
                ContentTop.Append(content4);


                // Create Report Details Section for regular time entries
                Row Header = new Row() { RowIndex = 5 };
                Cell Header1 = new Cell() { CellReference = "A5", CellValue = new CellValue("Datum"), DataType = CellValues.String };
                Header.Append(Header1);
                Cell Header2 = new Cell() { CellReference = "B5", CellValue = new CellValue("Sollzeit"), DataType = CellValues.String };
                Header.Append(Header2);
                Cell Header3 = new Cell() { CellReference = "C5", CellValue = new CellValue("Von"), DataType = CellValues.String };
                Header.Append(Header3);
                Cell Header4 = new Cell() { CellReference = "D5", CellValue = new CellValue("Bis"), DataType = CellValues.String };
                Header.Append(Header4);
                Cell Header5 = new Cell() { CellReference = "E5", CellValue = new CellValue("Total"), DataType = CellValues.String };
                Header.Append(Header5);
                Cell Header6 = new Cell() { CellReference = "F5", CellValue = new CellValue("Typ / Überzeit"), DataType = CellValues.String };
                Header.Append(Header6);
                Cell Header7 = new Cell() { CellReference = "G5", CellValue = new CellValue("Kommentar"), DataType = CellValues.String };
                Header.Append(Header7);

                sheetData.Append(Title, HeaderTop, ContentTop, Header);


                uint rownr = 6;

                using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
                {
                    conn.CreateTable<SollEntry>();
                    sollEntries = conn.Table<SollEntry>().Where(s => s.SortDate <= endDate && s.SortDate >= startDate).OrderByDescending(t => t.SortDate).ToList();


                    foreach (var sollItem in sollEntries)
                    {
                        sollItem.export(sheetData, rownr);
                        uint lookuprow = rownr;
                        rownr++;


                        List<TimeEntry> reportTimeEntries = conn.Table<TimeEntry>().Where(s => s.Date == sollItem.Date).OrderByDescending(t => t.SortDate).ToList();

                        double time = 0;

                        foreach (var timeItem in reportTimeEntries)
                        {

                            if (timeItem.Type != "NF")
                            {
                                time += timeItem.hours;
                                timeItem.export(sheetData, rownr);
                                rownr++;
                            }
                        }
                        Row ts = new Row { RowIndex = lookuprow };

                        string TimeSum = TimeSpanString(time);
                        Cell sumTime = new Cell { CellReference = String.Concat("E", lookuprow), CellValue = new CellValue(TimeSum), DataType = CellValues.String };
                        sheetData.Append(sumTime);

                    }
                    rownr++;
                    // Create Report Details Section for NF time entries
                    Row Emergency = new Row() { RowIndex = rownr };
                    Cell emergencyTitle = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("Notfälle und Überzeit"), DataType = CellValues.String };
                    Emergency.Append(emergencyTitle);
                    rownr++;

                    Row Emergencies = new Row() { RowIndex = rownr };
                    Cell emerg1 = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("Datum"), DataType = CellValues.String };
                    Emergencies.Append(emerg1);
                    Cell emerg3 = new Cell() { CellReference = String.Concat("C", rownr), CellValue = new CellValue("Von"), DataType = CellValues.String };
                    Emergencies.Append(emerg3);
                    Cell emerg4 = new Cell() { CellReference = String.Concat("D", rownr), CellValue = new CellValue("Bis"), DataType = CellValues.String };
                    Emergencies.Append(emerg4);
                    Cell emerg5 = new Cell() { CellReference = String.Concat("E", rownr), CellValue = new CellValue("Total"), DataType = CellValues.String };
                    Emergencies.Append(emerg5);
                    Cell emerg6 = new Cell() { CellReference = String.Concat("F", rownr), CellValue = new CellValue("Typ"), DataType = CellValues.String };
                    Emergencies.Append(emerg6);
                    Cell emerg7 = new Cell() { CellReference = String.Concat("G", rownr), CellValue = new CellValue("Kommentar"), DataType = CellValues.String };
                    Emergencies.Append(emerg7);
                    sheetData.Append(Emergency, Emergencies);
                    rownr++;

                    var nfEntries = conn.Table<TimeEntry>().Where(nf => nf.SortDate <= endDate && nf.SortDate >= startDate && nf.Type == "NF").OrderByDescending(t => t.SortDate).ToList();
                    foreach (var nfItem in nfEntries)
                    {
                        nfItem.export(sheetData, rownr);
                        rownr++;
                    }
                    rownr++;

                    // Create Report Details Section for ÜZ time entries
                    var otEntries = conn.Table<TimeEntry>().Where(ot => ot.Type == "ÜZ").OrderByDescending(t => t.SortDate).ToList();
                    foreach (var otItem in otEntries)
                    {
                        otItem.export(sheetData, rownr);
                        rownr += 2;
                    }

                    // legend for abbreviations
                    Row Legend1 = new Row() { RowIndex = rownr };
                    Cell legend11 = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("NF"), DataType = CellValues.String };
                    Legend1.Append(legend11);
                    Cell legend12 = new Cell() { CellReference = String.Concat("B", rownr), CellValue = new CellValue("Notfall (eingetragene Zeit wird nicht zur Überzeit hinzugerechnet, da monetär abgegolten)"), DataType = CellValues.String };
                    Legend1.Append(legend12);

                    rownr++;
                    Row Legend2 = new Row() { RowIndex = rownr };
                    Cell legend21 = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("WB"), DataType = CellValues.String };
                    Legend2.Append(legend21);
                    Cell legend22 = new Cell() { CellReference = String.Concat("B", rownr), CellValue = new CellValue("Weiterbildung (zählt als Arbeitszeit)"), DataType = CellValues.String };
                    Legend2.Append(legend22);

                    rownr++;
                    Row Legend3 = new Row() { RowIndex = rownr };
                    Cell legend31 = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("M"), DataType = CellValues.String };
                    Legend3.Append(legend31);
                    Cell legend32 = new Cell() { CellReference = String.Concat("B", rownr), CellValue = new CellValue("Meeting (zählt als Arbeitszeit)"), DataType = CellValues.String };
                    Legend3.Append(legend32);

                    rownr++;
                    Row Legend4 = new Row() { RowIndex = rownr };
                    Cell legend41 = new Cell() { CellReference = String.Concat("A", rownr), CellValue = new CellValue("ÜZ"), DataType = CellValues.String };
                    Legend4.Append(legend41);
                    Cell legend42 = new Cell() { CellReference = String.Concat("B", rownr), CellValue = new CellValue("Überzeit aus gelöschten Einträgen"), DataType = CellValues.String };
                    Legend4.Append(legend42);

                    sheetData.Append(Legend1, Legend2, Legend3, Legend4);

                }
                sheets.Append(sheet);
                workbookpart.Workbook.Save();
                spreadsheetDocument.Close();

                Launcher.OpenAsync(new OpenFileRequest()
                {
                    File = new ReadOnlyFile(filepath)

                });
            }

        }
        // deletes a the time entries bewteen two dates and replaces them with a single time entry to keep overtime accounted for
        public async void deleteData(object sender, EventArgs e)
        {
            calculate(null, null);
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<SollEntry>();
                sollEntries = conn.Table<SollEntry>().Where(s => s.SortDate <= endDate && s.SortDate >= startDate).ToList();
                conn.CreateTable<TimeEntry>();
                timeEntries = conn.Table<TimeEntry>().Where(s => s.SortDate <= endDate && s.SortDate >= startDate).ToList();

                string sDate = StartDatePicked.Date.ToString("dd.MM.yy");
                string eDate = EndDatePicked.Date.ToString("dd.MM.yy");

                bool answer = await DisplayAlert("Eintrag Löschen", "Wetsch die Iitrag von " + sDate + " bis " + eDate + " würkli lösche?", "Ja", "Nei");

                if (answer)
                {
                    TimeEntry overtime = new TimeEntry();
                    overtime.SortDate = 2011010000;
                    overtime.Comment = "Überzeit von " + sDate + " bis " + eDate;
                    overtime.hours = timeEntries.Sum(item => item.hours) - sollEntries.Sum(item => item.SollHours);
                    overtime.Date = "01.11.20";
                    overtime.Type = "ÜZ";
                    overtime.Time = TimeSpanString(overtime.hours);
                    overtime.TimeIn = "";
                    overtime.TimeOut = "";

                    foreach (var te in timeEntries)
                    {
                        conn.Delete<TimeEntry>(te.ID);
                    }
                    foreach (var se in sollEntries)
                    {
                        conn.Delete<SollEntry>(se.Date);
                    }
                    conn.Insert(overtime);
                }
            }
        }
    }
}