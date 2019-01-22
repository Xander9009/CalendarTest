using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
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
using System.Text.RegularExpressions;

namespace CalendarTest
{
    public partial class Form1 : Form
    {
        private static List<User> Users = new List<User>();
        private static List<SavableUserData> SavableUsers = new List<SavableUserData>();
        public static User CurrentUser;

        public Form1()
        {
            InitializeComponent();
            notificationIcon.Visible = false;

            if (File.Exists("users.json"))
            {
                using (StreamReader file = File.OpenText("users.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    SavableUsers = (List<SavableUserData>)serializer.Deserialize(file, typeof(List<SavableUserData>));
                    if (SavableUsers.Count > 0)
                    {
                        UserListView.Items.Clear();
                        foreach (SavableUserData user in SavableUsers)
                        {
                            UserListView.Items.Add(user.DisplayName);
                            Users.Add(new User(user.DisplayName, user.ID));
                        }
                        UserListView.Select();
                        UserListView.Items[0].Selected = true;
                    }
                }
            }
        }

        private void btnNewUser_Click(object sender, EventArgs e)
        {
            NewUser();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadUser();
        }

        private void btnSync_Click(object sender, EventArgs e)
        {
            Sync(CurrentUser);
        }

        private void btnCreateTimesheet_Click(object sender, EventArgs e)
        {
            //string StartDate = Prompt.ShowDialog("Please enter the first date of the timesheet in YYYY-MM-DD format.", "Timehseet start date.");
            var StartDate = "2019-01-15";
            if (StartDate.Length != 10) { MessageBox.Show("Date entered is invalid."); return; }
            List<Event> TimesheetEvents = TSheetsListTimesheetEvents(CurrentUser, StartDate);
            List<Event> MergedEvents = new List<Event>();
            
            Event LatestEvent = null;
            foreach (var TimesheetEvent in TimesheetEvents)
            {
                if (LatestEvent != null)
                {
                    if (LatestEvent.End == TimesheetEvent.Start)
                    {
                        LatestEvent.End = TimesheetEvent.End;
                    }
                    else
                    {
                        if (LatestEvent != null)
                        {
                            MergedEvents.Add(LatestEvent);
                        }
                        LatestEvent = TimesheetEvent;
                    }
                }
                else
                {
                    LatestEvent = TimesheetEvent;
                }
            }
            if (LatestEvent != null)
            {
                MergedEvents.Add(LatestEvent);
            }


            var getRequest = CurrentUser.GoogleSheetsService.Spreadsheets.Get("1uD1eclhcxHiWGj3raNbuLiTdI47cWoHV8AMJoikBx-8");
            getRequest.Ranges = "A1:Z50";
            getRequest.IncludeGridData = true;
            var Template = getRequest.Execute();

            Template.SpreadsheetId = null;
            Template.SpreadsheetUrl = null;

            SpreadsheetsResource.CreateRequest request = new SpreadsheetsResource.CreateRequest(
                CurrentUser.GoogleSheetsService,
                Template
                );
            var NewSpreadsheet = request.Execute();

            var ID = NewSpreadsheet.SpreadsheetId;

            var reqs = new BatchUpdateSpreadsheetRequest();
            reqs.Requests = new List<Request>();

            reqs.Requests.Add(UpdateCellValue("B2", new string[,] { { "Test" } }, Template));
            
            

            
            // Execute request
            var response = CurrentUser.GoogleSheetsService.Spreadsheets.BatchUpdate(reqs, ID).Execute(); // Replace Spreadsheet.SpreadsheetId with your recently created spreadsheet ID

            System.Diagnostics.Process.Start(NewSpreadsheet.SpreadsheetUrl);
        }

        private Request UpdateCellValue(string Range, string[,] Value, Spreadsheet Template)
        {
            // Create starting coordinate where data would be written to

            var StartRow = RowFromCellString(Range);
            var StartColumn = ColumnFromCellString(Range);
            var EndRow = RowFromCellString(Range);
            var EndColumn = ColumnFromCellString(Range);

            GridCoordinate gridCoordinate = new GridCoordinate();
            gridCoordinate.ColumnIndex = StartColumn;
            gridCoordinate.RowIndex = StartRow;
            gridCoordinate.SheetId = 542233618; // Your specific sheet ID here

            var request = new Request();
            request.UpdateCells = new UpdateCellsRequest();
            request.UpdateCells.Start = gridCoordinate;
            request.UpdateCells.Fields = "*"; // needed by API, throws error if null

            foreach (var Row in Value)
            {
                RowData rowData = new RowData();
                List<CellData> listCellData = new List<CellData>();

                foreach (var Cell in Row)
                {
                    ExtendedValue extendedValue = new ExtendedValue();
                    extendedValue.StringValue = Cell;

                    CellData cellData = new CellData();
                    cellData.UserEnteredValue = extendedValue;
                    cellData.EffectiveFormat = Template.Sheets[0].Data[0].RowData[Row].Values[Column].EffectiveFormat;
                    listCellData.Add(cellData);
                }

                // Assigning data to cells
                
                

                rowData.Values = listCellData;

                // Put cell data into a row
                List<RowData> listRowData = new List<RowData>();
                listRowData.Add(rowData);
                request.UpdateCells.Rows = listRowData;
            }
            
            return request;
        }

        private int ColumnFromCellString(string Cell)
        {
            Regex rx = new Regex(@"([a-z]+)[0-9]+", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var Column = rx.Match(Cell).Groups[1].Value;
            Column = Column.ToUpperInvariant();
            int sum = 0;
            for (int i = 0; i < Column.Length; i++)
            {
                sum *= 26;
                sum += (Column[i] - 'A' + 1);
            }
            return sum;
        }

        private int RowFromCellString(string Cell)
        {
            Regex rx = new Regex(@"[a-z]+([0-9]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            return int.Parse(rx.Match(Cell).Groups[1].Value);
        }

        private bool maxedWhenTrayed = false;
        private void trayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            maxedWhenTrayed = this.WindowState == FormWindowState.Maximized;
            this.WindowState = FormWindowState.Minimized;
            notificationIcon.Visible = true;
            this.ShowInTaskbar = false;
        }

        private void notificationIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (maxedWhenTrayed)
            {
                this.WindowState = FormWindowState.Maximized;
            }
            else
            {
                this.WindowState = FormWindowState.Normal;
            }
            notificationIcon.Visible = false;
            this.ShowInTaskbar = true;
        }

        private void UserListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadUser();
        }

        private void lblURL_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start((sender as Control).Text);
        }

        private void lblCalendar_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://calendar.google.com/calendar/embed?src=" + CurrentUser.Calendar);
        }

        internal void Sync(User user)
        {
            List<Event> TSheetsScheduleEvents = TSheetsListScheduleEvents(user);
            List<Event> GoogleCalendarEvents = GoogleCalendarListEvents(user);
            bool EventsChanged = false;
            foreach (var t in TSheetsScheduleEvents)
            {
                bool Found = false;
                foreach (var g in GoogleCalendarEvents)
                {
                    if (g.Start.Equals(t.Start) && g.End.Equals(t.End) && g.Job == t.Job + " (TS)")
                    {
                        Found = true;
                        break;
                    }
                }
                if (!Found)
                {
                    EventsResource.InsertRequest request = new EventsResource.InsertRequest(
                        user.GoogleCalendarService,
                        new Google.Apis.Calendar.v3.Data.Event()
                        {
                            Summary = t.Job + " (TS)",
                            Start = new EventDateTime() { DateTime = t.Start },
                            End = new EventDateTime() { DateTime = t.End },
                            Source = new Google.Apis.Calendar.v3.Data.Event.SourceData() { Title = "TSheets", Url = user.URL }
                        },
                        user.Calendar
                    );
                    request.Execute();
                    EventsChanged = true;
                }
            }

            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);
            foreach (var g in GoogleCalendarEvents)
            {
                bool Found = false;
                foreach (var t in TSheetsScheduleEvents)
                {
                    if (t.Start.Equals(g.Start) && t.End.Equals(g.End) && g.Job == t.Job + " (TS)")
                    {
                        Found = true;
                        break;
                    }
                }
                if (!Found)
                {
                    EventsResource.DeleteRequest request = new EventsResource.DeleteRequest(
                        user.GoogleCalendarService,
                        user.Calendar,
                        g.GoogleCalendarID
                        );
                    request.Execute();
                    EventsChanged = true;
                }
            }

            if (EventsChanged) GoogleCalendarListUpcomingEvents(user);
        }

        //Google Calendar methods
        public static void NewEvent(User user)
        {
            EventsResource.InsertRequest request = new EventsResource.InsertRequest(
                user.GoogleCalendarService,
                new Google.Apis.Calendar.v3.Data.Event()
                {
                    Description = "New test event",
                    Start = new EventDateTime() { DateTime = new DateTime(2019, 1, 9, 5, 0, 0) },
                    End = new EventDateTime() { DateTime = new DateTime(2019, 1, 10, 5, 0, 0) }
                },
                user.Calendar
                );
            request.Execute();
        }


        //TSheets methods

        /// <summary>
        /// Shows how to set up authentication to authenticate the user in an embedded browser form
        /// and get an OAuth2 token by prompting the user for credentials.
        /// </summary>
        private static void AuthenticateWithBrowser(User user)
        {
            //First, check if there's a saved token that we can use.

            string savedToken;
            UserAuthentication userAuthProvider;
            if (File.Exists(user.ID + "TStoken.json"))
            {
                using (StreamReader file = File.OpenText(user.ID + "TStoken.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    savedToken = (string)serializer.Deserialize(file, typeof(string));

                    // This can be restored into a UserAuthentication object later to reuse:
                    OAuthToken restoredToken = OAuthToken.FromJson(savedToken);
                    userAuthProvider = new UserAuthentication(user.TSheetsConnection, restoredToken);
                }
            }
            else
            {
                // The UserAuthentication class will handle the OAuth2 desktop
                // authentication flow using an embedded WebBrowser form, 
                // cache the returned token for later API usage, and handle token refreshes.
                userAuthProvider = new UserAuthentication(user.TSheetsConnection);
            }

            user.TSheetsAuthProvider = userAuthProvider;

            // optionally register an event handler to be notified if/when the auth
            // token changes
            userAuthProvider.TokenChanged += userAuthProvider_TokenChanged;

            // Retrieve a token from the server
            // Note: the RestApi class will call this as needed so it isn't required
            // to call it before accessing the API. However, manually calling GetToken first
            // is recommended so the app can more gracefully handle authentication errors
            OAuthToken authToken = userAuthProvider.GetToken();


            // OAuth2 tokens can and should be cached across application uses so users
            // don't need to grant access every time they run the application.
            // To do this, call OAuthToken.ToJSon to get a serialized version of
            // the token that can be used later. Be sure to treat this string as a 
            // user password and store it securely!
            // Note that this token will potentially be refreshed during API usage
            // using the OAuth2 token refresh protocol.  If that happens, your application
            // should overwrite the previously saved token with the new token value.
            // You can register for the TokenChanged event to be notified of any new/changed tokens
            // or you can call UserAuthentication.GetToken().ToJson() after using the API 
            // to manually retrieve the most current token.
            savedToken = authToken.ToJson();

            using (StreamWriter file = File.CreateText(user.ID + "TStoken.json"))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(file, savedToken);
            }
        }

        /// <summary>
        /// Event handler that will be called when the UserAuthentication OAuthToken changes
        /// </summary>
        static void userAuthProvider_TokenChanged(object sender, TokenChangedEventArgs e)
        {
            if (e.CurrentToken != null)
            {
                Console.WriteLine("Received new auth token:");
                Console.WriteLine(e.CurrentToken.ToJson());
            }
            else
            {
                Console.WriteLine("Token no longer valid");
            }
        }

        /// <summary>
        /// Lists the next 10 upcoming events from Google Calendar.
        /// </summary>
        private void GoogleCalendarListUpcomingEvents(User user)
        {
            try
            {
                // Define parameters of request.
                EventsResource.ListRequest request = user.GoogleCalendarService.Events.List(user.Calendar);
                request.TimeMin = DateTime.Now.AddDays(-7);
                request.TimeMax = DateTime.Now.AddDays(14);
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                // List events. Place event details in the listview.
                Events events = request.Execute();

                GoogleEventsList.Items.Clear();
                if (events.Items != null && events.Items.Count > 0)
                {
                    foreach (var eventItem in events.Items)
                    {
                        try
                        {
                            string when = eventItem.Start.DateTime.ToString();
                            if (when != "")
                            {
                                when = when.Substring(0, when.Length - 6) + " " + when.Substring(when.Length - 2, 2);
                            }
                            if (when == "")
                            {
                                when = eventItem.Start.Date.ToString();
                            }
                            try
                            {
                                var Hour = eventItem.End.DateTime.Value.Hour;
                                var Minute = eventItem.End.DateTime.Value.Minute;
                                var Period = "AM";
                                if (Hour > 12)
                                {
                                    Hour -= 12;
                                    Period = "PM";
                                }
                                when += string.Format("-{0}:{1:00} {2}", Hour, Minute, Period);
                            }
                            catch { }
                            var eventListViewItem = new ListViewItem(eventItem.Summary);
                            eventListViewItem.SubItems.Add(when);
                            GoogleEventsList.Items.Add(eventListViewItem);
                        }
                        catch { }
                    }
                }
                else
                {
                    var eventListViewItem = new ListViewItem(new string[] { "No upcoming events found", "" });
                    GoogleEventsList.Items.Add(eventListViewItem);
                }
            }
            catch { }
        }

        public static int CountPrimaryGoogleCalendarTSheetsEvents(User user)
        {
            try
            {
                // Define parameters of request.
                EventsResource.ListRequest request = user.GoogleCalendarService.Events.List("primary");
                request.ShowDeleted = false;
                request.SingleEvents = true;
                request.Q = "(TS)";
                request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

                // List events. Place event details in the listview.
                Events events = request.Execute();
                return events.Items.Count;
            }
            catch { }
            return 0;
        }

        public static void MigrateCalendarEvents(User user)
        {
            EventsResource.ListRequest request = user.GoogleCalendarService.Events.List("primary");
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.Q = "(TS)";
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events. Place event details in the listview.
            Events events = request.Execute();
            foreach (var eventItem in events.Items)
            {
                EventsResource.MoveRequest moveRequest = new EventsResource.MoveRequest(user.GoogleCalendarService, "primary", eventItem.Id, user.Calendar);
                moveRequest.Execute();
            }
        }

        internal void LoadUser()
        {
            if (UserListView.SelectedItems != null && UserListView.SelectedItems.Count > 0)
            {
                string DisplayName = UserListView.SelectedItems[0].Text;
                if (DisplayName != "")
                {
                    foreach (User user in Users)
                    {
                        if (user.DisplayName == DisplayName)
                        {
                            CurrentUser = user;
                            LoadUser(user);
                            break;
                        }
                    }
                }
            }
        }

        internal void LoadUser(User user)
        {
            GoogleAuthenticate(user);
            TSheetsAuthenticate(user);

            if (user.Calendar == "primary")
            {
                GoogleCalendarFindTSCalendar(user, out user.Calendar, out user.CalendarName);
                if (user.Calendar != "primary" && CountPrimaryGoogleCalendarTSheetsEvents(user) != 0)
                {
                    DialogResult dialogResult = MessageBox.Show("A Google Calendar specifically for TSheets has been found. Would you like to migrate all TSheets events from your primary calendar to the new calendar?", "Delete events?", MessageBoxButtons.YesNo);
                    if (dialogResult == DialogResult.Yes)
                    {
                        MigrateCalendarEvents(user);
                    }
                }
            }

            LoadUserLabelInfo(user);

            GoogleCalendarListUpcomingEvents(user);
            TSheetsListScheduleEvents(user);
        }

        private void LoadUserLabelInfo(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);
            var userData = tsheetsApi.Get(ObjectType.CurrentUser);
            var responseObject = JObject.Parse(userData);
            var userObject = responseObject.SelectToken("results.users.*");

            lblCurrentUser.Text = string.Format("{0} {1}", userObject["first_name"], userObject["last_name"]);
            lblEmail.Text = userObject["email"].ToString();
            user.URL = "https://" + userObject["client_url"].ToString() + ".tsheets.com";
            lblURL.Text = user.URL;
            lblCalendar.Text = user.CalendarName;
        }

        private List<Event> TSheetsListScheduleEvents(User user)
        {
            TSheetsEventsList.Items.Clear();

            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var filters = new Dictionary<string, string>();
            var N = DateTime.Now.AddDays(-7);
            filters.Add("start", N.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'") + "00+00:00");
            N = DateTime.Now.AddDays(14);
            filters.Add("end", N.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'") + "00+00:00");
            filters.Add("schedule_calendar_ids", "173108");
            var ScheduleData = tsheetsApi.Get(ObjectType.ScheduleEvents, filters);
            var ScheduleEventsObject = JObject.Parse(ScheduleData);
            var allScheduleEvents = ScheduleEventsObject.SelectTokens("results.schedule_events.*");

            var jobcodesData = tsheetsApi.Get(ObjectType.Jobcodes);
            var jobcodesObject = JObject.Parse(jobcodesData)["results"]["jobcodes"];

            var Events = new List<Event>();

            foreach (var ScheduleEvent in allScheduleEvents)
            {
                var tsUser = ScheduleEvent.SelectToken("supplemental_data.users." + ScheduleEvent["user_id"]);

                var JobCode = ScheduleEvent["jobcode_id"].ToString();
                JobCode = JobCode == "0" ? "Untitled" : jobcodesObject[JobCode]["name"].ToString();
                var eventListViewItem = new ListViewItem(string.Format("{0}", JobCode));
                eventListViewItem.SubItems.Add(string.Format("{0:g}-{1:t}", ScheduleEvent["start"], ScheduleEvent["end"]));
                TSheetsEventsList.Items.Add(eventListViewItem);

                Event E = new Event
                {
                    Type = Event.EventType.TSheets,
                    Start = (DateTime)ScheduleEvent["start"],
                    End = (DateTime)ScheduleEvent["end"],
                    Job = JobCode,
                    TSheetsID = (string)ScheduleEvent["id"],
                    URL = user.URL
                };
                Events.Add(E);
            }
            return Events;
        }

        private List<Event> TSheetsListTimesheetEvents(User user, string StartDate)
        {
            TSheetsEventsList.Items.Clear();

            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var filters = new Dictionary<string, string>();
            var N = new DateTime(int.Parse(StartDate.Substring(0, 4)), int.Parse(StartDate.Substring(5, 2)), int.Parse(StartDate.Substring(8, 2)));
            filters.Add("start_date", StartDate);
            N = DateTime.Now.AddDays(14);
            filters.Add("end_date", N.ToString("yyyy'-'MM'-'dd"));
            var ScheduleData = tsheetsApi.Get(ObjectType.Timesheets, filters);
            var ScheduleEventsObject = JObject.Parse(ScheduleData);
            var allScheduleEvents = ScheduleEventsObject.SelectTokens("results.timesheets.*");

            var jobcodesData = tsheetsApi.Get(ObjectType.Jobcodes);
            var jobcodesObject = JObject.Parse(jobcodesData)["results"]["jobcodes"];

            var Events = new List<Event>();

            foreach (var TimesheetEvent in allScheduleEvents)
            {
                var tsUser = TimesheetEvent.SelectToken("supplemental_data.users." + TimesheetEvent["user_id"]);

                var JobCode = TimesheetEvent["jobcode_id"].ToString();
                JobCode = JobCode == "0" ? "Untitled" : jobcodesObject[JobCode]["name"].ToString();
                var eventListViewItem = new ListViewItem(string.Format("{0}", JobCode));
                eventListViewItem.SubItems.Add(string.Format("{0:g}-{1:t}", TimesheetEvent["start"], TimesheetEvent["end"]));
                TSheetsEventsList.Items.Add(eventListViewItem);

                Event E = new Event
                {
                    Type = Event.EventType.TSheets,
                    Start = (DateTime)TimesheetEvent["start"],
                    End = (DateTime)TimesheetEvent["end"],
                    Job = JobCode,
                    TSheetsID = (string)TimesheetEvent["id"],
                    URL = user.URL
                };
                Events.Add(E);
            }
            //Events.Sort((x, y) => x.Start.CompareTo(y.Start));
            return Events;
        }

        private static List<Event> GoogleCalendarListEvents(User user)
        {
            List<Event> Events = new List<Event>();

            EventsResource.ListRequest request = user.GoogleCalendarService.Events.List(user.Calendar);
            request.TimeMin = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day - 7);
            request.TimeMax = request.TimeMin.Value.AddDays(21);
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            Events GCEvents = request.Execute();

            if (GCEvents.Items != null && GCEvents.Items.Count > 0)
            {
                foreach (var eventItem in GCEvents.Items)
                {
                    try
                    {
                        Event E = new Event
                        {
                            Type = Event.EventType.GoogleCalendar,
                            Start = (DateTime)eventItem.Start.DateTime,
                            End = (DateTime)eventItem.End.DateTime,
                            Job = (string)eventItem.Summary,
                            GoogleCalendarID = (string)eventItem.Id
                        };
                        Events.Add(E);
                    }
                    catch { }
                }
            }
            return Events;
        }

        internal void NewUser()
        {
            string userName = Prompt.ShowDialog("Please enter new user's name.", "Name to be displayed.");
            if (userName != "")
            {
                User user = new User(userName);
                Users.Add(user);
                CurrentUser = user;
                using (StreamWriter file = File.CreateText("users.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    serializer.Serialize(file, Users);
                }
                LoadUser(user);

                UserListView.Items.Add(user.DisplayName);
                UserListView.Items[UserListView.Items.Count - 1].Selected = true;
            }
        }

        internal static void GoogleAuthenticate(User user)
        {
            //This loads user token from file, but isn't user-specific, yet.
            using (var stream =
                new FileStream("googlecredentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = user.ID + "GCtoken.json";
                user.GoogleCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Program.GoogleScopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);

                
            }

            // Create Google Calendar API service.
            user.GoogleCalendarService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = user.GoogleCredential,
                ApplicationName = Program.ApplicationName,
            });

            // Create Google Sheets API service.
            user.GoogleSheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = user.GoogleCredential,
                ApplicationName = Program.ApplicationName,
            });
        }

        internal static void TSheetsAuthenticate(User user)
        {
            TSheetsCredentials TSCreds = null;
            if (File.Exists("tsheetscredentials.json"))
            {
                using (StreamReader file = File.OpenText("tsheetscredentials.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    TSCreds = (TSheetsCredentials)serializer.Deserialize(file, typeof(TSheetsCredentials));
                }
            }
            if (TSCreds != null)
            {
                // set up the ConnectionInfo object which tells the API how to connect to the server
                user.TSheetsConnection = new ConnectionInfo(TSCreds._baseUri, TSCreds._clientId, TSCreds._redirectUri, TSCreds._clientSecret);

                // AuthenticateWithBrowser will do a full OAuth2 forms based authentication in a
                //web browser form and prompt the user for credentials.
                AuthenticateWithBrowser(user);
            }
        }

        public static void GoogleCalendarFindTSCalendar(User user, out string calendar, out string calendarName)
        {
            string Calendar = "primary";
            string CalendarName = "Primary";

            var request = new CalendarListResource.ListRequest(user.GoogleCalendarService);
            var response = request.Execute();
            foreach (var cal in response.Items)
            {
                if (cal.Summary.ToLower() == "ts" || cal.Summary.ToLower() == "tsheets")
                {
                    Calendar = cal.Id;
                    CalendarName = cal.Summary;
                    break;
                }
            }
            calendar = Calendar;
            calendarName = CalendarName;
        }

        public class User
        {
            public string DisplayName;
            public int ID;
            public string URL;
            public string Calendar;
            public string CalendarName;

            public int testnum = 100;

            //Google Calendar variables
            public UserCredential GoogleCredential;
            public CalendarService GoogleCalendarService;
            public SheetsService GoogleSheetsService;

            //TSheets variables
            public ConnectionInfo TSheetsConnection;
            public IOAuth2 TSheetsAuthProvider;

            public User(string displayName)
            {
                DisplayName = displayName;
                ID = (new Random()).Next(1, 999999999);
                InitializeAuths(this);
            }

            [JsonConstructor]
            public User(string displayName, int id)
            {
                DisplayName = displayName;
                ID = id;
                InitializeAuths(this);
            }

            private void InitializeAuths(User user)
            {
                GoogleAuthenticate(user);
                TSheetsAuthenticate(user);
                GoogleCalendarFindTSCalendar(user, out user.Calendar, out user.CalendarName);
            }

            public SavableUserData GetSavableData()
            {
                var SavableData = new SavableUserData();
                SavableData.ID = ID;
                SavableData.DisplayName = DisplayName;
                SavableData.URL = URL;
                SavableData.Calendar = Calendar;
                SavableData.CalendarName = CalendarName;
                return SavableData;
            }
        }

        public class SavableUserData
        {
            public string DisplayName;
            public int ID;
            public string URL;
            public string Calendar;
            public string CalendarName;
        }

        public class TSheetsCredentials
        {
            public string _clientId = "6d46f274ed05492f9d46eb6b60dc0cf6";
            public string _redirectUri = "http://localhost";
            public string _clientSecret = "f91f618aca664c9487d6904ec3282dfd";
            public string _baseUri = "https://rest.tsheets.com/api/v1";
        }

        public class Event
        {
            public enum EventType { GoogleCalendar, TSheets, Both };
            public EventType Type;
            public DateTime Start;
            public DateTime End;
            public string Job;
            public string GoogleCalendarID;
            public string TSheetsID;
            public string URL;
        }

        public static class Prompt
        {
            public static string ShowDialog(string text, string caption)
            {
                Form prompt = new Form()
                {
                    Width = 500,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterScreen
                };
                Label textLabel = new Label() { Left = 50, Top = 20, Text = text };
                TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
                Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };
                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            }
        }
    }
}
