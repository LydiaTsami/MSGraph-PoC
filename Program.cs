using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
using Microsoft.Graph.CallRecords;
using Microsoft.Identity.Client;
using System;
using System.Linq;
using System.Net.Http.Headers;

namespace GraphTutorial
{
    class Program
    {

        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["tenantId"]) ||
                string.IsNullOrEmpty(appConfig["clientSecret"]))
            {
                return null;
            }

            return appConfig;
        }

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            Console.WriteLine(".NET Core Graph Tutorial\n");

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var tenantId = appConfig["tenantId"];
            var clientSecret = appConfig["clientSecret"];

            // Initialize the auth provider with values from appsettings.json
            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();


            //Install-Package Microsoft.Graph.Auth -PreRelease
            ClientCredentialProvider authProvider = new ClientCredentialProvider(confidentialClientApplication);

            // Initialize Graph client
            GraphHelper.Initialize(authProvider);

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Get meeting's members");
                Console.WriteLine("2. Get call's attendance");

                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch (choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;
                    case 1:
                        Console.WriteLine("Please input the meetingId");
                        string meetingId = Console.ReadLine();
                        var members = await GraphHelper.GetTeamMembers(meetingId);
                        foreach (User member in members.ToList())
                        {
                            Console.WriteLine("Found member: " + member.DisplayName);
                        }
                        //var meeting = await GraphHelper.GetMeeting("https://teams.microsoft.com/l/meetup-join/19%3ameeting_NTg3ZGQ5YjYtNjI1Ny00ZTQ4LTg0ZWMtMmI4ZjZkYjJkNGRj%40thread.v2/0?context=%7b%22Tid%22%3a%221aed5afa-c363-478e-ae14-b73b6949addb%22%2c%22Oid%22%3a%228b406a47-ed00-45cb-ad12-cd01d6143bbb%22%7d");
                        break;
                    case 2:
                        Console.WriteLine("Please input the callId");
                        string callId = Console.ReadLine();
                        //var callRecord = await GraphHelper.GetCallRecord(callId != "" ? callId : "f4ea5721-a7b5-44ee-8ceb-9dfa7a6dd41e");
                        var callRecord = await GraphHelper.GetCallRecordSessions(callId != "" ? callId : "f4ea5721-a7b5-44ee-8ceb-9dfa7a6dd41e");
                        var joinWebUrl = callRecord.JoinWebUrl;
                        if (joinWebUrl == null) break;
                        foreach (Session session in callRecord.Sessions)
                        {
                            ParticipantEndpoint caller = (ParticipantEndpoint)session.Caller;
                            var user = await GraphHelper.GetUserAsync(caller.Identity.User.Id);
                            //TODO - Find the mapping between this userId and the university's student ID.
                            StudentEvent studentEvent = new StudentEvent
                            {
                                //TODO - Find course ID based on joinWebUrl.
                                CourseID = "COMP0088", // Course ID Upper case.
                                Timestamp = ((DateTimeOffset)session.StartDateTime).UtcDateTime,
                                EventType = EventType.Attendance,
                                ActivityType = "Meeting",
                                ActivityName = "Weekly Lecture",
                                Student = new Student
                                {
                                    Email = user.Mail,
                                    FirstName = user.GivenName,
                                    LastName = user.Surname,
                                    ID = user.Id
                                }
                            };
                            Console.WriteLine(studentEvent.ToString());
                            //_eventAggregator.ProcessEvent(studentEvent);
                        }
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }
    }
}