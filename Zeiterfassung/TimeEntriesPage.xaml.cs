using System.Collections.Generic;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Zeiterfassung
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class TimeEntriesPage : ContentPage
    {
        List<TimeEntry> timeEntries = new List<TimeEntry>();
        

        public TimeEntriesPage()
        {
            InitializeComponent();
        }
        //loads list of timeentries on appearing
        protected override void OnAppearing()
        {
            base.OnAppearing();
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<TimeEntry>();
                timeEntries = conn.Table<TimeEntry>().OrderByDescending(t => t.SortDate).ToList();
                TimeEntryList.ItemsSource = timeEntries;
            }
        }
        // deletes selected item. item is selected by clicking / touching it.
        // calls on appearing to immediately display changes
        async void delete(object sender, SelectedItemChangedEventArgs e)
        {
            if (e.SelectedItem == null)
            {
                return;
            }
            var choices = new[] { "iitrag lösche", "iitrag ändere" };
            var answer = await DisplayActionSheet("Was hesch vor?","abbreche","", choices);

            System.Diagnostics.Debug.WriteLine(answer);
            if (answer == "iitrag lösche")
            {
                int Index = e.SelectedItemIndex;
                TimeEntry selectedtimeEntry = timeEntries[Index];
                using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
                {
                    conn.Delete<TimeEntry>(selectedtimeEntry.ID);
                }
                OnAppearing();
            }
            else if (answer == "iitrag ändere")
            {
                int Index = e.SelectedItemIndex;
                TimeEntry selectedtimeEntry = timeEntries[Index];
                var timeInChange = selectedtimeEntry.TimeIn;
                var timeOutChange = selectedtimeEntry.TimeOut;
                string dateChange = selectedtimeEntry.Date;
                string typeChange = selectedtimeEntry.Type;
                string commentChange = selectedtimeEntry.Comment;
                await Navigation.PushModalAsync(new ManualEntryPage(dateChange, timeInChange,timeOutChange,typeChange,commentChange));

                using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
                {
                    conn.Delete<TimeEntry>(selectedtimeEntry.ID);
                }
                
            }



        }
    }
}