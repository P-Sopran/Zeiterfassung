using System;
using System.Diagnostics;
using System.Collections.Generic;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Zeiterfassung
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class ManualEntryPage : ContentPage
    {
        private bool changeEntry = false;
        public ManualEntryPage(string date, string timeIn, string TimeOut, string Type, string comments)
        {
            changeEntry = true;
            InitializeComponent();
            string[] dateInfo = date.Split('.');

            int year = Convert.ToInt32("20" + dateInfo[2]);
            int month = Convert.ToInt32(dateInfo[1]);
            int day = Convert.ToInt32(dateInfo[0]);
            
            DateTime then = new DateTime(year, month, day);
            

            DatePicked.Date = then;
            TimeInPicked.Time = TimeSpan.Parse(timeIn);
            TimeOutPicked.Time = TimeSpan.Parse(TimeOut);
            comment.Text = comments;

            switch (Type)
            {
                case ("WB"):
                    type.SelectedIndex= 3;
                    break;
                case ("NF"):
                    type.SelectedIndex = 1;
                    break;
                case ("M"):
                    type.SelectedIndex = 2;
                    break;
                default:
                    type.SelectedIndex = 0;
                    break;
            }

        }
        public ManualEntryPage()
        {
            InitializeComponent();
        }
        // thanks to https://www.youtube.com/watch?v=JhWwBOoqXQ8&t=1801s&ab_channel=AltexSoft
        public async void enterToDB(object sender, EventArgs e)
        {
            TimeSpan TimeIn = TimeInPicked.Time;
            TimeSpan TimeOut = TimeOutPicked.Time;
            /*
            if (TimeIn > TimeOut)
            {
                Debug.WriteLine("------");
                bool answer = await DisplayAlert("Obacht!", "Hesch würkli über Mitternach gschaffet oder dich vertippt", "stimmt scho so", "ups...");
                
                if(!answer)
                {
                    return;
                }
            }
            */
            string Type;
            string Comment = comment.Text;
            string check = type.SelectedItem.ToString();
            switch (check)
            {
                case "Notfall":
                    Type = "NF";
                    break;
                case "Meeting":
                    Type = "M";
                    break;
                case "Wiiterbiudig":
                    Type = "WB";
                    break;
                default:
                    Type = " ";
                    break;
            }
            TimeEntry manualEntry = new TimeEntry(DatePicked.Date, TimeIn, TimeOut, Comment);
            manualEntry.Type = Type;

            var se = new SollEntry(SollTimePicked.Time, DatePicked.Date);
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<TimeEntry>();
                conn.Insert(manualEntry);
                if (Type != "NF")
                {
                    conn.CreateTable<SollEntry>();
                    conn.InsertOrReplace(se);
                }
            }

            if (changeEntry)
            {
                this.SendBackButtonPressed();
            }
        }
    }
}