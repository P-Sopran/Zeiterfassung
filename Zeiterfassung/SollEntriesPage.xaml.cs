using System.Collections.Generic;
using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Zeiterfassung
{
    [XamlCompilation(XamlCompilationOptions.Compile)]
    public partial class SollEntriesPage : ContentPage
    {
        private bool answer = false;
        List<SollEntry> sollEntries = new List<SollEntry>();
        public SollEntriesPage()
        {
            InitializeComponent();
        }
        protected override void OnAppearing()
        {
            base.OnAppearing();
            using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
            {
                conn.CreateTable<SollEntry>();
                sollEntries = conn.Table<SollEntry>().OrderByDescending(t => t.SortDate).ToList();
                SollEntryList.ItemsSource = sollEntries;
            }
        }

        // same as delete of time entry
        async void delete(object sender, SelectedItemChangedEventArgs e)
        {
            if (e.SelectedItem == null)
            {
                return;
            }
            answer = await DisplayAlert("Eintrag Löschen", "Wetsch de Iitrag würkli lösche", "Ja", "Nei");

            if (answer)
            {
                int Index = e.SelectedItemIndex;
                SollEntry selectedSollEntry = sollEntries[Index];
                using (SQLite.SQLiteConnection conn = new SQLite.SQLiteConnection(App.DB_Path))
                {
                    conn.Delete<SollEntry>(selectedSollEntry.Date);
                }
                OnAppearing();
                answer = false;
            }
        }
    }
}