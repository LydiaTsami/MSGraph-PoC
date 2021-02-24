using Microsoft.Graph;
using Microsoft.Graph.CallRecords;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TimeZoneConverter;

namespace GraphTutorial
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        public static async Task<User> GetUserAsync(string userId)
        {
            try
            {
                // GET /users/{id}
                return await graphClient.Users[userId]
                    .Request()
                    .Select(e => new
                    {
                        e.Mail,
                        e.GivenName,
                        e.Surname,
                        e.Id
                    })
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting user: {ex.Message}");
                return null;
            }
        }

        public static async Task<IGroupMembersCollectionWithReferencesPage> GetTeamMembers(string teamId)
        {
            try
            {
                // GET /groups/{groupId}/Members
                return await graphClient.Groups[teamId]
                    .Members
                    .Request()
                    //.Select(u => new
                    //{
                    //    u.DisplayName,
                    //    u.id
                    //})
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting team members: {ex.Message}");
                return null;
            }
        }

        public static async Task<CallRecord> GetCallRecordSessions(string callId)
        {
            try
            {
                // GET /groups/{groupId}/Members
                return await graphClient.Communications
                    .CallRecords[callId]
                    .Request()
                    .Expand("sessions")
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting call participants: {ex.Message}");
                return null;
            }
        }

        public static async Task<CallRecord> GetCallRecord(string callId)
        {
            try
            {
                // GET /communications/callRecords/{callId}
                return await graphClient.Communications
                    .CallRecords[callId]
                    .Request()
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting call participants: {ex.Message}");
                return null;
            }
        }



        public static async Task<IUserOnlineMeetingsCollectionPage> GetMeeting(string callId)
        {
            try
            {
                // GET /communications/callRecords/{callId}
                return await graphClient.Users["8b406a47-ed00-45cb-ad12-cd01d6143bbb"]
                    .OnlineMeetings
                    .Request()
                    .Filter("JoinWebUrl eq '" + callId + "'")
                    .GetAsync();
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting call participants: {ex.Message}");
                return null;
            }
        }

    }
}