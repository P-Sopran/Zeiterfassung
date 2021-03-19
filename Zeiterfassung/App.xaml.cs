using Xamarin.Forms;

namespace Zeiterfassung
{

    public partial class App : Application
    {
        public static string DB_Path = string.Empty;
        public App()
        {
            InitializeComponent();
            MainPage = new NavigationPage(new MainCarouselPage());
        }

        public App(string Db_Path)
        {
            InitializeComponent();
            DB_Path = Db_Path;
            MainPage = new MainCarouselPage();
        }



        protected override void OnStart()
        {
        }

        protected override void OnSleep()
        {
        }

        protected override void OnResume()
        {
        }
    }
}
