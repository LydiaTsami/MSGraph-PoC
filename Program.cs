using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Auth;
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
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }

            return appConfig;
        }

        static string FormatDateTimeTimeZone(
            Microsoft.Graph.DateTimeTimeZone value,
            string dateTimeFormat)
        {
            // Parse the date/time string from Graph into a DateTime
            var dateTime = DateTime.Parse(value.DateTime);

            return dateTime.ToString(dateTimeFormat);
        }

        static void ListCalendarEvents(string userTimeZone, string dateTimeFormat)
        {
            var events = GraphHelper
                .GetCurrentWeekCalendarViewAsync(DateTime.Today, userTimeZone)
                .Result;

            Console.WriteLine("Events:");

            foreach (var calendarEvent in events)
            {
                Console.WriteLine($"Subject: {calendarEvent.Subject}");
                Console.WriteLine($"  Organizer: {calendarEvent.Organizer.EmailAddress.Name}");
                Console.WriteLine($"  Start: {FormatDateTimeTimeZone(calendarEvent.Start, dateTimeFormat)}");
                Console.WriteLine($"  End: {FormatDateTimeTimeZone(calendarEvent.End, dateTimeFormat)}");
            }
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
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

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
                        break;
                    case 2:
                        Console.WriteLine("Please input the callId");
                        string callId = Console.ReadLine();
                        var callRecord = await GraphHelper.GetCallRecordSessions(callId);
                        //foreach (Participant participant in callRecord.Callee.ToList())
                        //{
                        //    Console.WriteLine("Found participant: " + participant.DisplayName);
                        //}
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }
    }
}