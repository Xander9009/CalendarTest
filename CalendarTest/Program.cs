using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Threading;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

using TSheets;

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
