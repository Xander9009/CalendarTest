﻿using System;
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

using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;

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
            dtpPayDate.Value = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
            notificationIcon.Visible = false;

            if (System.IO.File.Exists("users.json"))
            {
                using (StreamReader file = System.IO.File.OpenText("users.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    SavableUsers = (List<SavableUserData>)serializer.Deserialize(file, typeof(List<SavableUserData>));
                    if (SavableUsers.Count > 0)
                    {
                        UserListView.Items.Clear();
                        foreach (SavableUserData user in SavableUsers)
                        {
                            UserListView.Items.Add(user.DisplayName);
                            Users.Add(new User(user.DisplayName, user.EmployeeNumber, user.ID));
                        }
                        UserListView.Select();
                        UserListView.Items[0].Selected = true;
                    }
                }
            }
            foreach (var file in Directory.EnumerateFiles(Environment.CurrentDirectory, "*TSToken.json"))
            {
                Regex rx = new Regex(@"\\([0-9]+)TStoken.json");
                var match = rx.Match(file);
                if (match.Success)
                {
                    var found = false;
                    foreach (User user in Users)
                    {
                        if (user.ID.ToString() == match.Groups[1].Value)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        System.IO.File.Delete(file);
                    }
                }
            }
            foreach (var directory in Directory.EnumerateDirectories(Environment.CurrentDirectory, "*GCToken.json"))
            {
                Regex rx = new Regex(@"\\([0-9]+)GCtoken.json");
                var match = rx.Match(directory);
                if (match.Success)
                {
                    var found = false;
                    foreach (User user in Users)
                    {
                        if (user.ID.ToString() == match.Groups[1].Value)
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found)
                    {
                        Directory.Delete(directory, true);
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
            List<Event> TimesheetEvents = TSheetsListTimesheetEvents(CurrentUser, dtpPayDate.Value.AddDays(-20));
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

            var templateSearchRequest =  CurrentUser.GoogleDriveService.Files.List();
            templateSearchRequest.Q = "name='CDC Timesheet Template'";
            var templateSearchResponse = templateSearchRequest.Execute();
            if (templateSearchResponse.Files.Count == 0)
            {
                MessageBox.Show("Template file not found. Please make a copy of the template file which this user owns.");
                return;
            }
            var TemplateID = templateSearchResponse.Files[0].Id;

            var getRequest = CurrentUser.GoogleSheetsService.Spreadsheets.Get(TemplateID);
            getRequest.Ranges = "A1:Z50";
            getRequest.IncludeGridData = true;
            var Template = getRequest.Execute();

            Template.SpreadsheetId = null;
            Template.SpreadsheetUrl = null;
            Template.Properties.Title = "CDC Timesheet (" + dtpPayDate.Value.ToString("yyyy-MM-dd") + ")";
            
            SpreadsheetsResource.CreateRequest request = new SpreadsheetsResource.CreateRequest(
                CurrentUser.GoogleSheetsService,
                Template
                );
            var NewSpreadsheet = request.Execute();

            var ID = NewSpreadsheet.SpreadsheetId;

            var reqs = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
            };

            reqs.Requests.Add(UpdateCellValue("C3", CurrentUser.EmployeeNumber, Template));
            reqs.Requests.Add(UpdateCellValue("C6", CurrentUser.DisplayName, Template));
            reqs.Requests.Add(UpdateCellValue("V2", dtpPayDate.Value.ToString("MM/dd/yyyy"), Template));
            var Week1 = new string[7,12];
            double WeekTotal1 = 0;
            double RegTimeTotal1 = 0;
            double OverTimeTotal1 = 0;
            double PTOTimeTotal1 = 0;
            double HDPTimeTotal1 = 0;
            for (int i = 0; i < 7; i++)
            {
                int Shifts = 0;
                var Date = dtpPayDate.Value.AddDays(i - 20);
                double PTOTime = 0;
                double HDPTime = 0;
                double OverTime = 0;
                double RegTime = 0;
                double DayTotal = 0;
                foreach (var Event in MergedEvents)
                {
                    if (Date.Month == Event.Start.Month && Date.Day == Event.Start.Day)
                    {
                        if (Event.PTO > 0)
                        {
                            PTOTime += Event.PTO;
                            PTOTimeTotal1 += Event.PTO;
                            DayTotal += Event.PTO;
                            WeekTotal1 += Event.PTO;
                        }
                        else if (Event.HDP > 0)
                        {
                            HDPTime += Event.HDP;
                            HDPTimeTotal1 += Event.HDP;
                            DayTotal += Event.HDP;
                            WeekTotal1 += Event.HDP;
                        }
                        else
                        {
                            Week1[i, Shifts] = Event.Start.ToString("h:mm tt");
                            if (Week1[i, Shifts] == "00:00 AM")
                            {
                                Week1[i, Shifts] = "12:00 PM";
                            }
                            Week1[i, Shifts + 1] = Event.End.ToString("h:mm tt");
                            if (Week1[i, Shifts + 1] == "00:00 AM")
                            {
                                Week1[i, Shifts + 1] = "12:00 PM";
                            }
                            DayTotal += (Event.End - Event.Start).TotalHours;
                            RegTime += (Event.End - Event.Start).TotalHours;
                            Shifts++;
                        }
                    }
                }
                if (RegTime > 0)
                {
                    var RegWeekTimeLeft = 40 - RegTimeTotal1;
                    if (RegWeekTimeLeft > 0)
                    {
                        if (RegWeekTimeLeft < RegTime)
                        {
                            OverTime += RegTime - RegWeekTimeLeft;
                            RegTime = RegWeekTimeLeft;
                        }
                    }
                    else
                    {
                        OverTime += RegTime;
                        RegTime = 0;
                    }
                    WeekTotal1 += RegTime + OverTime;
                    RegTimeTotal1 += RegTime;
                    OverTimeTotal1 += OverTime;
                    if (RegTime > 0)
                    {
                        Week1[i, 6] = RegTime.ToString();
                    }
                }
                if (OverTime > 0)
                {
                    Week1[i, 7] = OverTime.ToString();
                }
                if (PTOTime > 0)
                {
                    Week1[i, 8] = PTOTime.ToString();
                }
                if (HDPTime > 0)
                {
                    Week1[i, 9] = HDPTime.ToString();
                }
                Week1[i, 11] = DayTotal.ToString();
            }
            reqs.Requests.Add(UpdateCellValue("C12", Week1, Template));
            reqs.Requests.Add(UpdateCellValue("I19", new string[,] { { RegTimeTotal1.ToString(), OverTimeTotal1.ToString(), PTOTimeTotal1.ToString(), HDPTimeTotal1.ToString(), "", WeekTotal1.ToString() } }, Template));

            var Week2 = new string[7, 12];
            double WeekTotal2 = 0;
            double RegTimeTotal2 = 0;
            double OverTimeTotal2 = 0;
            double PTOTimeTotal2 = 0;
            double HDPTimeTotal2 = 0;
            for (int i = 0; i < 7; i++)
            {
                int Shifts = 0;
                var Date = dtpPayDate.Value.AddDays(i - 20);
                double PTOTime = 0;
                double HDPTime = 0;
                double OverTime = 0;
                double RegTime = 0;
                double DayTotal = 0;
                foreach (var Event in MergedEvents)
                {
                    if (Date.Month == Event.Start.Month && Date.Day == Event.Start.Day)
                    {
                        if (Event.PTO > 0)
                        {
                            PTOTime += Event.PTO;
                            PTOTimeTotal2 += Event.PTO;
                            DayTotal += Event.PTO;
                            WeekTotal2 += Event.PTO;
                        }
                        else if (Event.HDP > 0)
                        {
                            HDPTime += Event.HDP;
                            HDPTimeTotal2 += Event.HDP;
                            DayTotal += Event.HDP;
                            WeekTotal2 += Event.HDP;
                        }
                        else
                        {
                            Week2[i, Shifts] = Event.Start.ToString("h:mm tt");
                            if (Week2[i, Shifts] == "00:00 AM")
                            {
                                Week2[i, Shifts] = "12:00 PM";
                            }
                            Week2[i, Shifts + 1] = Event.End.ToString("h:mm tt");
                            if (Week2[i, Shifts + 1] == "00:00 AM")
                            {
                                Week2[i, Shifts + 1] = "12:00 PM";
                            }
                            DayTotal += (Event.End - Event.Start).TotalHours;
                            RegTime += (Event.End - Event.Start).TotalHours;
                            Shifts++;
                        }
                    }
                }
                if (RegTime > 0)
                {
                    var RegWeekTimeLeft = 40 - RegTimeTotal2;
                    if (RegWeekTimeLeft > 0)
                    {
                        if (RegWeekTimeLeft < RegTime)
                        {
                            OverTime += RegTime - RegWeekTimeLeft;
                            RegTime = RegWeekTimeLeft;
                        }
                    }
                    else
                    {
                        OverTime += RegTime;
                        RegTime = 0;
                    }
                    WeekTotal2 += RegTime + OverTime;
                    RegTimeTotal2 += RegTime;
                    OverTimeTotal2 += OverTime;
                    if (RegTime > 0)
                    {
                        Week2[i, 6] = RegTime.ToString();
                    }
                }
                if (OverTime > 0)
                {
                    Week2[i, 7] = OverTime.ToString();
                }
                if (PTOTime > 0)
                {
                    Week2[i, 8] = PTOTime.ToString();
                }
                if (HDPTime > 0)
                {
                    Week2[i, 9] = HDPTime.ToString();
                }
                Week2[i, 11] = DayTotal.ToString();
            }
            reqs.Requests.Add(UpdateCellValue("C20", Week2, Template));
            reqs.Requests.Add(UpdateCellValue("I27", new string[,] { { RegTimeTotal2.ToString(), OverTimeTotal2.ToString(), PTOTimeTotal2.ToString(), HDPTimeTotal2.ToString(), "", WeekTotal2.ToString() } }, Template));


            // Execute request
            var response = CurrentUser.GoogleSheetsService.Spreadsheets.BatchUpdate(reqs, ID).Execute();

            System.Diagnostics.Process.Start(NewSpreadsheet.SpreadsheetUrl);
        }

        private void dtpPayDate_ValueChanged(object sender, EventArgs e)
        {
            LoadUser();
        }

        private Request UpdateCellValue(string StartCell, int Value, Spreadsheet Template)
        {
            return UpdateCellValue(StartCell, new int[,] { { Value } }, Template);
        }

        private Request UpdateCellValue(string StartCell, string Value, Spreadsheet Template)
        {
            return UpdateCellValue(StartCell, new string[,] { { Value } }, Template);
        }

        private Request UpdateCellValue(string StartCell, string[,] Values, Spreadsheet Template)
        {
            // Create starting coordinate where data would be written to

            GridCoordinate gridCoordinate = new GridCoordinate();
            gridCoordinate.ColumnIndex = ColumnFromCellString(StartCell);
            gridCoordinate.RowIndex = RowFromCellString(StartCell);
            gridCoordinate.SheetId = Template.Sheets[0].Properties.SheetId;

            var request = new Request();
            request.UpdateCells = new UpdateCellsRequest();
            request.UpdateCells.Start = gridCoordinate;
            request.UpdateCells.Fields = "*"; // needed by API, throws error if null

            List<RowData> listRowData = new List<RowData>();

            for (int Row = 0; Row < Values.GetLength(0); Row += 1)
            {
                List<CellData> listCellData = new List<CellData>();

                for (int Column = 0; Column < Values.GetLength(1); Column += 1)
                {
                    CellData cellData = new CellData();
                    cellData.UserEnteredValue = new ExtendedValue() { StringValue = Values[Row, Column] };
                    cellData.EffectiveFormat = Template.Sheets[0].Data[0].RowData[(int)gridCoordinate.RowIndex + Row].Values[(int)gridCoordinate.ColumnIndex + Column].EffectiveFormat;
                    cellData.UserEnteredFormat = Template.Sheets[0].Data[0].RowData[(int)gridCoordinate.RowIndex + Row].Values[(int)gridCoordinate.ColumnIndex + Column].UserEnteredFormat;

                    listCellData.Add(cellData);
                }

                RowData rowData = new RowData() { Values = listCellData };

                // Put cell data into a row
                
                listRowData.Add(rowData);
            }

            request.UpdateCells.Rows = listRowData;
            return request;
        }

        private Request UpdateCellValue(string StartCell, int[,] Values, Spreadsheet Template)
        {
            // Create starting coordinate where data would be written to

            GridCoordinate gridCoordinate = new GridCoordinate();
            gridCoordinate.ColumnIndex = ColumnFromCellString(StartCell);
            gridCoordinate.RowIndex = RowFromCellString(StartCell);
            gridCoordinate.SheetId = Template.Sheets[0].Properties.SheetId;

            var request = new Request();
            request.UpdateCells = new UpdateCellsRequest();
            request.UpdateCells.Start = gridCoordinate;
            request.UpdateCells.Fields = "*"; // needed by API, throws error if null

            List<RowData> listRowData = new List<RowData>();

            for (int Row = 0; Row < Values.GetLength(0); Row += 1)
            {
                List<CellData> listCellData = new List<CellData>();

                for (int Column = 0; Column < Values.GetLength(1); Column += 1)
                {
                    CellData cellData = new CellData();
                    cellData.UserEnteredValue = new ExtendedValue() { NumberValue = Values[Row, Column] };
                    cellData.EffectiveFormat = Template.Sheets[0].Data[0].RowData[(int)gridCoordinate.RowIndex + Row].Values[(int)gridCoordinate.ColumnIndex + Column].EffectiveFormat;
                    cellData.UserEnteredFormat = Template.Sheets[0].Data[0].RowData[(int)gridCoordinate.RowIndex + Row].Values[(int)gridCoordinate.ColumnIndex + Column].UserEnteredFormat;

                    listCellData.Add(cellData);
                }

                RowData rowData = new RowData() { Values = listCellData };

                // Put cell data into a row

                listRowData.Add(rowData);
            }

            request.UpdateCells.Rows = listRowData;
            return request;
        }

        private int ColumnFromCellString(string Cell, bool End = false)
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
            return sum - 1;
        }

        private int RowFromCellString(string Cell, bool End = false)
        {
            Regex rx = new Regex(@"[a-z]+([0-9]+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            return int.Parse(rx.Match(Cell).Groups[1].Value) - 1;
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
            if (System.IO.File.Exists(user.ID + "TStoken.json"))
            {
                using (StreamReader file = System.IO.File.OpenText(user.ID + "TStoken.json"))
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

            using (StreamWriter file = System.IO.File.CreateText(user.ID + "TStoken.json"))
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

                request.TimeMin = dtpPayDate.Value.AddDays(-20);
                request.TimeMax = dtpPayDate.Value.AddDays(-6);
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
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var ScheduleID = (int)JObject.Parse(tsheetsApi.Get(ObjectType.ScheduleCalendars)).SelectTokens("results.schedule_calendars.*").First()["id"];
            var filters = new Dictionary<string, string>();
            var N = dtpPayDate.Value.AddDays(-20);
            filters.Add("start", N.ToString("yyyy'-'MM'-'dd'T'") + "00:00:00+00:00");
            N = dtpPayDate.Value.AddDays(-6);
            filters.Add("end", N.ToString("yyyy'-'MM'-'dd'T'") + "00:00:00+00:00");
            filters.Add("schedule_calendar_ids", ScheduleID.ToString());
            var ScheduleData = JObject.Parse(tsheetsApi.Get(ObjectType.ScheduleEvents, filters));
            var allScheduleEvents = ScheduleData.SelectTokens("results.schedule_events.*");

            Dictionary<int, string> Jobs = new Dictionary<int, string>() { { 0, "Untitled" } };

            var jobcodesObject = JObject.Parse(tsheetsApi.Get(ObjectType.Jobcodes))["results"]["jobcodes"];
            foreach (var Jobcode in jobcodesObject)
            {
                foreach (var J in Jobcode)
                {
                    Jobs.Add((int)J["id"], (string)J["name"]);
                }
            }

            Dictionary<int, string> PTOCodes = new Dictionary<int, string>();
            jobcodesObject = JObject.Parse(tsheetsApi.Get(ObjectType.Jobcodes, new Dictionary<string, string>() { { "type", "pto" } }))["results"]["jobcodes"];
            foreach (var Jobcode in jobcodesObject)
            {
                foreach (var J in Jobcode)
                {
                    PTOCodes.Add((int)J["id"], (string)J["name"]);
                }
            }

            var Events = new List<Event>();

            foreach (var ScheduleEvent in allScheduleEvents)
            {
                var tsUser = ScheduleEvent.SelectToken("supplemental_data.users." + ScheduleEvent["user_id"]);

                var JobcodeID = (int)ScheduleEvent["jobcode_id"];
                string JobName = "Untitled";
                bool PTO = false;
                if (Jobs.ContainsKey(JobcodeID))
                {
                    JobName = Jobs[JobcodeID];
                }
                else if (PTOCodes.ContainsKey(JobcodeID))
                {
                    JobName = PTOCodes[JobcodeID];
                    PTO = true;
                }

                Event E = new Event
                {
                    Type = Event.EventType.TSheets,
                    Job = JobName,
                    TSheetsID = (string)ScheduleEvent["id"],
                    URL = user.URL,
                    Start = (DateTime)ScheduleEvent["start"],
                    End = (DateTime)ScheduleEvent["end"]
                };
                
                Events.Add(E);
            }
            Events.Sort((x, y) => x.Start.CompareTo(y.Start));
            TSheetsEventsList.Items.Clear();
            foreach (var Event in Events)
            {
                var eventListViewItem = new ListViewItem(Event.Job);
                eventListViewItem.SubItems.Add(string.Format("{0}-{1}", Event.Start.ToString("MM/dd/yyyy HH:mm tt"), Event.End.ToString("HH:mm tt")));
                TSheetsEventsList.Items.Add(eventListViewItem);
            }
            return Events;
        }

        private List<Event> TSheetsListTimesheetEvents(User user, DateTime StartDate)
        {
            TSheetsEventsList.Items.Clear();

            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var filters = new Dictionary<string, string>();
            filters.Add("start_date", StartDate.ToString("yyyy'-'MM'-'dd"));
            filters.Add("end_date", StartDate.AddDays(14).ToString("yyyy'-'MM'-'dd"));
            var ScheduleData = tsheetsApi.Get(ObjectType.Timesheets, filters);
            var ScheduleEventsObject = JObject.Parse(ScheduleData);
            var allScheduleEvents = ScheduleEventsObject.SelectTokens("results.timesheets.*");

            Dictionary<int, string> Jobs = new Dictionary<int, string>() { { 0, "Untitled" } };

            var jobcodesObject = JObject.Parse(tsheetsApi.Get(ObjectType.Jobcodes))["results"]["jobcodes"];
            foreach (var Jobcode in jobcodesObject)
            {
                foreach (var J in Jobcode)
                {
                    Jobs.Add((int)J["id"], (string)J["name"]);
                }
            }

            Dictionary<int, string> PTOCodes = new Dictionary<int, string>();
            jobcodesObject = JObject.Parse(tsheetsApi.Get(ObjectType.Jobcodes, new Dictionary<string, string>() { { "type", "pto" } }))["results"]["jobcodes"];
            foreach (var Jobcode in jobcodesObject)
            {
                foreach (var J in Jobcode)
                {
                    PTOCodes.Add((int)J["id"], (string)J["name"]);
                }
            }

            var Events = new List<Event>();

            foreach (var TimesheetEvent in allScheduleEvents)
            {
                var tsUser = TimesheetEvent.SelectToken("supplemental_data.users." + TimesheetEvent["user_id"]);
                var JobcodeID = (int)TimesheetEvent["jobcode_id"];
                string JobName = "Untitled";
                int Type = 0;
                if (Jobs.ContainsKey(JobcodeID))
                {
                    JobName = Jobs[JobcodeID];
                }
                else if (PTOCodes.ContainsKey(JobcodeID))
                {
                    JobName = PTOCodes[JobcodeID];
                    if (JobName == "Holiday")
                    {
                        Type = 1;
                    }
                    else if (JobName == "PTO")
                    {
                        Type = 2;
                    }
                }

                Event E = new Event
                {
                    Type = Event.EventType.TSheets,
                    Job = JobName,
                    TSheetsID = (string)TimesheetEvent["id"],
                    URL = user.URL
                };
                if (Type == 0)
                {
                    E.Start = (DateTime)TimesheetEvent["start"];
                    E.End = (DateTime)TimesheetEvent["end"];
                }
                else
                {
                    var Date = (string)TimesheetEvent["date"];
                    E.Start = new DateTime(int.Parse(Date.Substring(0, 4)), int.Parse(Date.Substring(5, 2)), int.Parse(Date.Substring(8, 2)));
                    E.End = E.Start;
                    if (Type == 1)
                    {
                        E.HDP = (double)TimesheetEvent["duration"] / 3600;
                    }
                    else if (Type == 2)
                    {
                        E.PTO = (double)TimesheetEvent["duration"] / 3600;
                    }
                }
                Events.Add(E);
            }
            return Events;
        }

        private List<Event> GoogleCalendarListEvents(User user)
        {
            List<Event> Events = new List<Event>();

            EventsResource.ListRequest request = user.GoogleCalendarService.Events.List(user.Calendar);
            request.TimeMin = dtpPayDate.Value.AddDays(-20);
            request.TimeMax = dtpPayDate.Value.AddDays(-6);
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
                int employeeNumber;
                if (int.TryParse(Prompt.ShowDialog("Please enter " + userName + "'s employee number.", "Emplyee Number"), out employeeNumber))
                {
                    User user = new User(userName, employeeNumber);
                    Users.Add(user);
                    CurrentUser = user;
                    using (StreamWriter file = System.IO.File.CreateText("users.json"))
                    {
                        JsonSerializer serializer = new JsonSerializer();
                        serializer.Serialize(file, Users);
                    }
                    LoadUser(user);

                    UserListView.Items.Add(user.DisplayName);
                    UserListView.Items[UserListView.Items.Count - 1].Selected = true;
                }
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

            // Create Google Drive API service.
            user.GoogleDriveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = user.GoogleCredential,
                ApplicationName = Program.ApplicationName,
            });
        }

        internal static void TSheetsAuthenticate(User user)
        {
            TSheetsCredentials TSCreds = null;
            if (System.IO.File.Exists("tsheetscredentials.json"))
            {
                using (StreamReader file = System.IO.File.OpenText("tsheetscredentials.json"))
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
            public int EmployeeNumber;
            public int ID;
            public string URL;
            public string Calendar;
            public string CalendarName;

            public int testnum = 100;

            //Google Calendar variables
            public UserCredential GoogleCredential;
            public CalendarService GoogleCalendarService;
            public SheetsService GoogleSheetsService;
            public DriveService GoogleDriveService;

            //TSheets variables
            public ConnectionInfo TSheetsConnection;
            public IOAuth2 TSheetsAuthProvider;

            public User(string displayName, int employeeNumber)
            {
                DisplayName = displayName;
                ID = (new Random()).Next(1, 999999999);
                EmployeeNumber = employeeNumber;
                InitializeAuths(this);
            }

            [JsonConstructor]
            public User(string displayName, int employeeNumber, int id)
            {
                DisplayName = displayName;
                ID = id;
                EmployeeNumber = employeeNumber;
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
                SavableData.EmployeeNumber = EmployeeNumber;
                SavableData.URL = URL;
                SavableData.Calendar = Calendar;
                SavableData.CalendarName = CalendarName;
                return SavableData;
            }
        }

        public class SavableUserData
        {
            public string DisplayName;
            public int EmployeeNumber;
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
            public double PTO = 0;
            public double HDP = 0;
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
                Label textLabel = new Label() { Left = 50, Top = 20, Text = text, Width = 400 };
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
