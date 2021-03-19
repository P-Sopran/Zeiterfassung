using System;
using Xamarin.Forms;

namespace Zeiterfassung
{
    public partial class MainPage : ContentPage
    {
        private TimeSpan TimeIn;
        private TimeSpan TimeOut;

        public MainPage()
        {
            InitializeComponent();
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<CheckinEntry>();
                if (conn.Table<CheckinEntry>().Count() > 0)
                {
                    check_out.IsEnabled = true;
                    check_in.IsEnabled = false;

                }
            }
        }

        //sets checkin entry in preparation for saving a time entry to the database
        //inserts a soll entry to the database (overwrites if one already exists on that day)
        public void checkin(object sender, EventArgs e)
        {
            var ce = new CheckinEntry(DateTime.Now.TimeOfDay);
            ce.comment = comment.Text;


            string check = type.SelectedItem.ToString(); // save
            switch (check) // save
            {
                case "Notfall":
                    ce.type = "NF";
                    break;
                case "Meeting":
                    ce.type = "M";
                    break;
                case "Wiiterbildig":
                    ce.type = "WB";
                    break;
                default:
                    ce.type = " ";
                    break;
            }

            var se = new SollEntry(SollZeit.Time, DateTime.Now);
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<CheckinEntry>();
                conn.InsertOrReplace(ce);

                conn.CreateTable<SollEntry>();
                conn.InsertOrReplace(se);
            }

            check_out.IsEnabled = true;
            check_in.IsEnabled = false;

        }

        // gets info from checkin entry and saves new time entry to database
        public void checkout(object sender, EventArgs e)
        {
            TimeOut = DateTime.Now.TimeOfDay;
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {

                CheckinEntry checkinInfo = conn.Table<CheckinEntry>().ToList()[0];
                TimeIn = TimeSpan.FromHours(checkinInfo.hrs);

                DateTime checkindate = DateTime.Now;
                if (TimeIn.CompareTo(TimeOut) == 1)
                {
                    checkindate = DateTime.Today.AddDays(-1);
                }

                TimeEntry te = new TimeEntry(checkindate, TimeIn, TimeOut, checkinInfo.comment);
                te.Type = checkinInfo.type;

                conn.CreateTable<TimeEntry>();
                conn.Insert(te);


                conn.DropTable<CheckinEntry>();


            }

            check_out.IsEnabled = false;
            check_in.IsEnabled = true;


        }

    }
}