using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Text;
using System.Threading;

using Newtonsoft.Json.Linq;
using TSheets;
using Newtonsoft.Json;

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
        static string[] Scopes = { CalendarService.Scope.CalendarEvents };
        static string ApplicationName = "Google Calendar API .NET Quickstart";
        
        //Other variables
        private static List<User> Users = new List<User>();
        private static List<SavableUserData> SavableUsers = new List<SavableUserData>();
        private static Form1 form;
        public static User CurrentUser;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            form = new Form1();

            //Theoretically needed steps:
            //Step 1: Get a list of saved users. -Check
            //Step 2a: Choose an existing user or choose to add a new one. -Check
            //Step 2b: If adding a new user, add user to list. -Check
            //Step 3a: Load user's tokens, if they exist. -Check
            //Step 3b: If tokens don't exist, authorize new tokens. -Check
            //Step 4: Use tokens to access TSheets and compare to Google Calendar.
            //Step 5: Mirror changes to Google Calendar.

            //Step 1: Get a list of saved users.
            if (File.Exists("users.json"))
            {
                using (StreamReader file = File.OpenText("users.json"))
                {
                    JsonSerializer serializer = new JsonSerializer();
                    SavableUsers = (List<SavableUserData>)serializer.Deserialize(file, typeof(List<SavableUserData>));
                    if (SavableUsers.Count > 0)
                    {
                        ListView UserListView = form.Controls.Find("UserListView", true)[0] as ListView;
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

            Application.Run(form);
        }

        //Google Calendar methods
        static public void NewEvent(User user)
        {
            EventsResource.InsertRequest request = new EventsResource.InsertRequest(user.GoogleService,
                new Event()
                {
                    Description = "New test event",
                    Start = new EventDateTime() { DateTime = new DateTime(2019, 1, 9, 5, 0, 0) },
                    End = new EventDateTime() { DateTime = new DateTime(2019, 1, 10, 5, 0, 0) }
                },
                "primary"
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
        /// Shows how to get all users for the company
        /// </summary>
        private static void GetUsersSample(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);
            var userData = tsheetsApi.Get(ObjectType.Users);
            var responseObject = JObject.Parse(userData);

            var users = responseObject.SelectTokens("results.users.*");
            foreach (var userObject in users)
            {
                Console.WriteLine(string.Format("Current User: {0} {1}, email = {2}, client url = {3}",
                    userObject["first_name"],
                    userObject["last_name"],
                    userObject["email"],
                    userObject["client_url"]));
            }
        }

        /// <summary>
        /// Shows how to receive all timesheets for a given timeframe by using filter arguments
        /// and how to access the supplemental data in the response.
        /// Supplemental data will contain all of the user/jobcode/etc related data
        /// about the selected timesheets. API users should use the supplemental data when available
        /// rather than making additional calls to the server to receive that information.
        /// </summary>
        private static void GetTimesheetsSample(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var filters = new Dictionary<string, string>();
            filters.Add("start_date", "2014-01-01");
            filters.Add("end_date", "2020-01-01");
            var timesheetData = tsheetsApi.Get(ObjectType.Timesheets, filters);

            var timesheetsObject = JObject.Parse(timesheetData);
            var allTimeSheets = timesheetsObject.SelectTokens("results.timesheets.*");
            foreach (var timesheet in allTimeSheets)
            {
                Console.WriteLine(string.Format("Timesheet: ID={0}, Duration={1}, Data={2}, tz={3}",
                    timesheet["id"], timesheet["duration"], timesheet["date"], timesheet["tz"]));

                // get the associated user for this timesheet
                var tsUser = timesheetsObject.SelectToken("supplemental_data.users." + timesheet["user_id"]);
                Console.WriteLine(string.Format("\tUser: {0} {1}", tsUser["first_name"], tsUser["last_name"]));
            }
        }

        /// <summary>
        /// The Get api calls can potentially return many records from the server. The TSheets rest APIs
        /// support a paging request model so API clients can request records in smaller chunks.
        /// This sample shows how to request all available jobcodes using paging filters to retrieve
        /// the records this way.
        /// </summary>
        private static void GetJobcodesByPageSample(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);
            var filters = new Dictionary<string, string>();

            // start by requesting the first page
            int currentPage = 1;

            // and set our items per page to be 2
            // Note: 50 is the recommended per_page value for normal usage. This sample
            // is using a smaller number to make the sample more clear. Be sure to 
            // manually create >2 jobcodes in your account to see the paging happen.
            filters["per_page"] = "2";

            bool moreData = true;
            while (moreData)
            {
                filters["page"] = currentPage.ToString();

                var getResponse = tsheetsApi.Get(ObjectType.Jobcodes, filters);
                var responseObject = JObject.Parse(getResponse);

                // see if we have more pages to retrieve
                moreData = bool.Parse(responseObject.SelectToken("more").ToString());

                // increment to the next page
                currentPage++;

                var jobcodes = responseObject.SelectTokens("results.jobcodes.*");
                foreach (var jobcode in jobcodes)
                {
                    Console.WriteLine(string.Format("Jobcode Name: {0}, type = {1}, shortcode = {2}",
                        jobcode["name"],
                        jobcode["type"],
                        jobcode["short_code"]));
                }
            }
        }


        /// <summary>
        /// Shows how to create a user, create a jobcode, log time against it, and then run a project report
        /// that shows them
        /// </summary>
        public static void ProjectReportSample(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            DateTime today = DateTime.Now;
            string todayString = today.ToString("yyyy-MM-dd");
            DateTime tomorrow = today + new TimeSpan(1, 0, 0, 0);
            string tomorrowString = tomorrow.ToString("yyyy-MM-dd");

            // create a user
            int userId = CreateUser(tsheetsApi);

            // now create a jobcode we can log time against
            int jobCodeId = CreateJobCode(tsheetsApi);

            // log some time
            {
                var timesheetObjects = new List<JObject>();
                dynamic timesheet = new JObject();
                timesheet.user_id = userId;
                timesheet.jobcode_id = jobCodeId;
                timesheet.type = "manual";
                timesheet.duration = 3600;
                timesheet.date = todayString;
                timesheetObjects.Add(timesheet);

                timesheet = new JObject();
                timesheet.user_id = userId;
                timesheet.jobcode_id = jobCodeId;
                timesheet.type = "manual";
                timesheet.duration = 7200;
                timesheet.date = todayString;
                timesheetObjects.Add(timesheet);

                var addTimesheetResponse = tsheetsApi.Add(ObjectType.Timesheets, timesheetObjects);
                Console.WriteLine(addTimesheetResponse);

                var addedTimesheets = JObject.Parse(addTimesheetResponse);
            }

            // and run the report
            {
                dynamic reportOptions = new JObject();
                reportOptions.data = new JObject();
                reportOptions.data.start_date = todayString;
                reportOptions.data.end_date = tomorrowString;

                var projectReport = tsheetsApi.GetReport(ReportType.Project, reportOptions.ToString());

                Console.WriteLine(projectReport);
            }
        }

        /// <summary>
        /// Shows how to add, edit, and delete a timesheet
        /// </summary>
        private static void AddEditDeleteTimesheetSample(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            DateTime today = DateTime.Now;
            string todayString = today.ToString("yyyy-MM-dd");
            DateTime yesterday = today - new TimeSpan(1, 0, 0, 0);
            string yesterdayString = yesterday.ToString("yyyy-MM-dd");

            // create a user
            int userId = CreateUser(tsheetsApi);

            // now create a jobcode we can log time against
            int jobCodeId = CreateJobCode(tsheetsApi);

            // add a couple of timesheets
            var timesheetsToAdd = new List<JObject>();
            dynamic timesheet = new JObject();
            timesheet.user_id = userId;
            timesheet.jobcode_id = jobCodeId;
            timesheet.type = "manual";
            timesheet.duration = 3600;
            timesheet.date = todayString;
            timesheetsToAdd.Add(timesheet);

            timesheet = new JObject();
            timesheet.user_id = userId;
            timesheet.jobcode_id = jobCodeId;
            timesheet.type = "manual";
            timesheet.duration = 7200;
            timesheet.date = todayString;
            timesheetsToAdd.Add(timesheet);

            var result = tsheetsApi.Add(ObjectType.Timesheets, timesheetsToAdd);
            Console.WriteLine(result);

            // pull out the ids of the new timesheets
            var addedTimesheets = JObject.Parse(result).SelectTokens("results.timesheets.*");
            var timesheetIds = new List<int>();
            foreach (var ts in addedTimesheets)
            {
                timesheetIds.Add((int)ts["id"]);
            }

            // make some edits
            var timesheetsToEdit = new List<JObject>();
            timesheet = new JObject();
            timesheet.id = timesheetIds[0];
            timesheet.date = yesterdayString;
            timesheetsToEdit.Add(timesheet);
            timesheet = new JObject();
            timesheet.id = timesheetIds[1];
            timesheet.date = yesterdayString;
            timesheetsToEdit.Add(timesheet);
            result = tsheetsApi.Edit(ObjectType.Timesheets, timesheetsToEdit);
            Console.WriteLine(result);

            // and delete them
            result = tsheetsApi.Delete(ObjectType.Timesheets, timesheetIds);
            Console.WriteLine(result);
        }

        /// <summary>
        /// Helper to create random string so we don't get duplicate name conflicts if samples
        /// are run multiple times
        /// </summary>
        private static string CreateRandomString()
        {
            var randData = new Random().Next(1000).ToString();
            return randData;
        }

        /// <summary>
        /// Helper to create a random job code
        /// </summary>
        private static int CreateJobCode(RestClient tsheetsApi)
        {
            var jobCodeObjects = new List<JObject>();
            dynamic jobCode = new JObject();
            jobCode.name = "jc" + CreateRandomString();
            jobCode.assigned_to_all = true;
            jobCodeObjects.Add(jobCode);

            var addJobCodeResponse = tsheetsApi.Add(ObjectType.Jobcodes, jobCodeObjects);
            Console.WriteLine(addJobCodeResponse);

            // get the job code ID so we can use it later
            var addedJobCode = JObject.Parse(addJobCodeResponse).SelectToken("results.jobcodes.1");
            return (int)addedJobCode["id"];
        }

        /// <summary>
        /// Helper to create a random user
        /// </summary>
        private static int CreateUser(RestClient tsheetsApi)
        {
            var userObjects = new List<JObject>();
            dynamic user = new JObject();
            user.username = "user" + CreateRandomString();
            user.password = "Pa$$W0rd";
            user.first_name = "first";
            user.last_name = "last";
            userObjects.Add(user);

            var addUserResponse = tsheetsApi.Add(ObjectType.Users, userObjects);
            Console.WriteLine(addUserResponse);

            // get the user ID so we can use it later
            var addedUser = JObject.Parse(addUserResponse).SelectToken("results.users.1");
            return (int)addedUser["id"];
        }

        //Google Calendar methods

        /// <summary>
        /// Lists the next 10 upcoming events from Google Calendar.
        /// </summary>
        private static void GoogleCalendarListEvents(User user)
        {
            //Example calendar event request. First, we need to get a list of events to see which TSheets events have been mirrored already.
            // Define parameters of request.
            EventsResource.ListRequest request = user.GoogleService.Events.List("primary");
            request.TimeMin = DateTime.Now;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            request.MaxResults = 10;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            // List events. Place event details in the listview.
            Events events = request.Execute();

            var eventsListView = (ListView)form.Controls.Find("GoogleEventsList", true)[0];
            eventsListView.Items.Clear();
            if (events.Items != null && events.Items.Count > 0)
            {
                foreach (var eventItem in events.Items)
                {
                    string when = eventItem.Start.DateTime.ToString();
                    if (String.IsNullOrEmpty(when))
                    {
                        when = eventItem.Start.Date;
                    }
                    var eventListViewItem = new ListViewItem(eventItem.Summary);
                    eventListViewItem.SubItems.Add(when);
                    eventsListView.Items.Add(eventListViewItem);
                }
            }
            else
            {
                var eventListViewItem = new ListViewItem(new string[] { "No upcoming events found", "" });
                eventsListView.Items.Add(eventListViewItem);
            }
        }

        internal static void LoadUser()
        {
            ListView UserListView = form.Controls.Find("UserListView", true)[0] as ListView;
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

        internal static void LoadUser(User user)
        {
            
            GoogleAuthenticate(user);
            TSheetsAuthenticate(user);

            LoadUserLabelInfo(user);

            //Google Calendar samples
            GoogleCalendarListEvents(user);

            //TSheets samples
            TSheetsListTimesheets(user);
            //Example methods that pull in information. Only posts results to console.
            //GetUsersSample(user);
            //GetTimesheetsSample(user);

            //These two samples create users and job codes which then have to be cleaned
            //up manually (and multiple users aren't available for free accounts, anyway).
            //So, they've been disabled. However, they're left intact so they can be examined.
            //ProjectReportSample(user);
            //AddEditDeleteTimesheetSample(user);

            //GetJobcodesByPageSample(user);
        }

        private static void LoadUserLabelInfo(User user)
        {
            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);
            var userData = tsheetsApi.Get(ObjectType.CurrentUser);
            var responseObject = JObject.Parse(userData);
            var userObject = responseObject.SelectToken("results.users.*");

            form.Controls.Find("lblCurrentUser", true)[0].Text = string.Format("{0} {1}", userObject["first_name"], userObject["last_name"]);
            form.Controls.Find("lblEmail", true)[0].Text = userObject["email"].ToString();
            form.Controls.Find("lblURL", true)[0].Text = userObject["client_url"].ToString() + ".tsheets.com";
            LinkLabel L = form.Controls.Find("lblURL", true)[0] as LinkLabel;
            L.LinkArea = new LinkArea(0, L.Text.Length);
        }

        private static void TSheetsListTimesheets(User user)
        {
            ListView TSheetsEventList = form.Controls.Find("TSheetsEventsList", true)[0] as ListView;
            TSheetsEventList.Items.Clear();

            var tsheetsApi = new RestClient(user.TSheetsConnection, user.TSheetsAuthProvider);

            var filters = new Dictionary<string, string>();
            filters.Add("start_date", "2019-1-12");
            filters.Add("end_date", "2019-2-1");
            var timesheetData = tsheetsApi.Get(ObjectType.Timesheets, filters);
            var timesheetsObject = JObject.Parse(timesheetData);
            var allTimeSheets = timesheetsObject.SelectTokens("results.timesheets.*");

            var jobcodesData = tsheetsApi.Get(ObjectType.Jobcodes);
            var jobcodesObject = JObject.Parse(jobcodesData)["results"]["jobcodes"];
            
            foreach (var timesheet in allTimeSheets)
            {
                var tsUser = timesheetsObject.SelectToken("supplemental_data.users." + timesheet["user_id"]);

                var eventListViewItem = new ListViewItem(string.Format("{0} {1}", tsUser["first_name"], tsUser["last_name"]));
                eventListViewItem.SubItems.Add(string.Format("{0:g}-{1:t}", timesheet["start"], timesheet["end"]));
                eventListViewItem.SubItems.Add(string.Format("{0}", jobcodesObject[timesheet["jobcode_id"].ToString()]["name"]));
                TSheetsEventList.Items.Add(eventListViewItem);
            }
        }

        internal static void NewUser()
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

                ListView UserListView = form.Controls.Find("UserListView", true)[0] as ListView;
                UserListView.Items.Add(user.DisplayName);
                UserListView.Items[UserListView.Items.Count - 1].Selected = true;
            }
        }

        static internal void GoogleAuthenticate(User user)
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
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Calendar API service.
            user.GoogleService = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = user.GoogleCredential,
                ApplicationName = ApplicationName,
            });
        }

        static internal void TSheetsAuthenticate(User user)
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
    }

    internal class User
    {
        public string DisplayName;
        public int ID;

        //Google Calendar variables
        public UserCredential GoogleCredential;
        public CalendarService GoogleService;

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
            Program.GoogleAuthenticate(user);
            Program.TSheetsAuthenticate(user);
        }

        public SavableUserData GetSavableData()
        {
            var SavableData = new SavableUserData();
            SavableData.ID = ID;
            SavableData.DisplayName = DisplayName;
            return SavableData;
        }
    }

    internal class SavableUserData
    {
        public string DisplayName;
        public int ID;
    }

    internal class TSheetsCredentials
    {
        public string _clientId = "6d46f274ed05492f9d46eb6b60dc0cf6";
        public string _redirectUri = "http://localhost";
        public string _clientSecret = "f91f618aca664c9487d6904ec3282dfd";
        public string _baseUri = "https://rest.tsheets.com/api/v1";
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
