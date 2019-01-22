using System;
using System.Windows.Forms;
using Google.Apis.Calendar.v3;
using Google.Apis.Sheets.v4;

namespace CalendarTest
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>

        //Google Calendar variables

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/calendar-dotnet-quickstart.json
        public static string[] GoogleScopes = { CalendarService.Scope.CalendarReadonly, CalendarService.Scope.CalendarEvents, SheetsService.Scope.Spreadsheets };
        public static string ApplicationName = "Google Calendar API .NET Quickstart";
        
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
