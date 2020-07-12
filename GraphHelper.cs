using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace GraphMessagesSample
{
    public class GraphHelper
    {
        private static GraphServiceClient graphClient;
        public static void Initialize(IAuthenticationProvider authProvider)
        {
            graphClient = new GraphServiceClient(authProvider);
        }

        public static async Task<IEnumerable<Group>> GetGroupsAsync()
        {
            try
            {
                List<Group> groups = new List<Group>();
                var resultPage = await graphClient.Groups.Request().GetAsync();
                while (true)
                {
                    groups.AddRange(resultPage.CurrentPage);
                    if (resultPage.NextPageRequest == null)
                    {
                        break;
                    }
                    resultPage = resultPage.NextPageRequest.GetAsync().Result;
                }
                return groups;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }
        public static async Task<IEnumerable<DirectoryObject>> GetGroupMembersAsync(string groupId)
        {
            try
            {
                List<DirectoryObject> dirobjects = new List<DirectoryObject>();
                var resultPage = await graphClient.Groups[groupId].Members.Request().GetAsync();
                while (true)
                {
                    dirobjects.AddRange(resultPage.CurrentPage);
                    if (resultPage.NextPageRequest == null)
                    {
                        break;
                    }
                    resultPage = resultPage.NextPageRequest.GetAsync().Result;
                }
                return dirobjects;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }
        public static async Task<IEnumerable<Message>> GetMessagesAsync(string userId)
        {
            try
            {
                List<Message> dirobjects = new List<Message>();
                var resultPage = await graphClient.Users[userId].Messages.Request().GetAsync();
                while (true)
                {
                    dirobjects.AddRange(resultPage.CurrentPage);
                    if (resultPage.NextPageRequest == null)
                    {
                        break;
                    }
                    resultPage = resultPage.NextPageRequest.GetAsync().Result;
                }
                return dirobjects;
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return null;
            }
        }
    }
}