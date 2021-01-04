
using System;
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Newtonsoft.Json;

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
                Console.WriteLine($"group.DisplayName: {group.DisplayName}");
                var members = GraphHelper.GetGroupMembersAsync(group.Id).Result;
                foreach (Microsoft.Graph.User member in members)
                {
                    // Console.WriteLine($"member.Id: {member.Id}");
                    // Console.WriteLine($"member.DisplayName: {member.DisplayName}");
                    // Console.WriteLine($"member.UserPrincipalName: {member.UserPrincipalName}");
                    Console.WriteLine(JsonConvert.SerializeObject(member));
                    Console.WriteLine("-----------------------");

                    var messages = GraphHelper.GetMessagesAsync(member.Id).Result;
                    foreach (Microsoft.Graph.Message message in messages)
                    {
                        // Console.WriteLine($"message.Subject: {message.Subject}");
                        Console.WriteLine(JsonConvert.SerializeObject(message));
                    }
                    Console.WriteLine("=======================");
                }
            }
        }

        static void GetEvents()
        {
            var groups = GraphHelper.GetGroupsAsync().Result;
            foreach (var group in groups)
            {
                Console.WriteLine($"group.DisplayName: {group.DisplayName}");
                var members = GraphHelper.GetGroupMembersAsync(group.Id).Result;
                foreach (Microsoft.Graph.User member in members)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(member));
                    Console.WriteLine("-----------------------");

                    var calendars = GraphHelper.GetCalendarAsync(member.Id).Result;
                    foreach (Microsoft.Graph.Calendar calendar in calendars)
                    {
                        var events = GraphHelper.GetEventsAsync(member.Id, calendar.Id).Result;
                        foreach (Microsoft.Graph.Event e in events)
                        {
                            Console.WriteLine(JsonConvert.SerializeObject(e));
                            Console.WriteLine("+++++");
                        }
                        Console.WriteLine("=======================");
                    }
                }
            }
        }
        static void GetContacts()
        {
            var groups = GraphHelper.GetGroupsAsync().Result;
            foreach (var group in groups)
            {
                Console.WriteLine($"group.DisplayName: {group.DisplayName}");
                var members = GraphHelper.GetGroupMembersAsync(group.Id).Result;
                foreach (Microsoft.Graph.User member in members)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(member));
                    Console.WriteLine("-----------------------");

                    var contacts = GraphHelper.GetContactsAsync(member.Id).Result;
                    foreach (Microsoft.Graph.Contact c in contacts)
                    {
                        Console.WriteLine(JsonConvert.SerializeObject(c));
                        Console.WriteLine("=======================");
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

            // GetMessages();
            // GetEvents();
            GetContacts();
        }
    }
}
