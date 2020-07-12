using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;

namespace GraphMessagesSample
{
    class Program
    {
        static Dictionary<string, string> LoadClientSecretAppSettings()
        {
            Dictionary<string, string> result = null;
            // Get config ftom AppSettings
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();
            var appId = appConfig["appId"];
            var scopes = appConfig["scopes"];
            var tenantId = appConfig["tenantId"];
            var clientSecret = appConfig["clientSecret"];
            if (string.IsNullOrEmpty(appId) == false &&
                string.IsNullOrEmpty(scopes) == false &&
                string.IsNullOrEmpty(tenantId) == false &&
                string.IsNullOrEmpty(clientSecret) == false)
            {
                result = new Dictionary<string, string>()
                {
                    {"appId", appId},
                    {"scopes", scopes},
                    {"tenantId", tenantId},
                    {"clientSecret", clientSecret}
                };
            }
            return result;
        }

        static void GetMessages()
        {
            var groups = GraphHelper.GetGroupsAsync().Result;
            foreach (var group in groups)
            {
                Console.WriteLine($"group.Id: {group.Id}");
                Console.WriteLine($"group.DisplayName: {group.DisplayName}");
                var members = GraphHelper.GetGroupMembersAsync(group.Id).Result;
                foreach (Microsoft.Graph.User member in members)
                {
                    Console.WriteLine($"member.Id: {member.Id}");
                    Console.WriteLine($"member.DisplayName: {member.DisplayName}");
                    Console.WriteLine($"member.UserPrincipalName: {member.UserPrincipalName}");
                    var messages = GraphHelper.GetMessagesAsync(member.Id).Result;
                    foreach (Microsoft.Graph.Message message in messages)
                    {
                        Console.WriteLine($"message.Subject: {message.Subject}");
                    }
                }
            }
        }

        static void Main(string[] args)
        {
            IAuthenticationProvider authProvider = null;

            var appConfig = LoadClientSecretAppSettings();
            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid AppSettings");
                return;
            }
            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');
            var tenantId = appConfig["tenantId"];
            var clientSecret = appConfig["clientSecret"];
            // Initialize the auth provider
            authProvider = new ClientSecretAuthProvider(appId, scopes, tenantId, clientSecret);
            // Initialize Graph client
            GraphHelper.Initialize(authProvider);
            // Get messages
            GetMessages();
        }
    }
}
