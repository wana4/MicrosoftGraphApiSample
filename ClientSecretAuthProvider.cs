using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http;
using System.Threading.Tasks;

namespace GraphMessagesSample
{
    public class ClientSecretAuthProvider : IAuthenticationProvider
    {
        private IConfidentialClientApplication msalClient;
        private string appId;
        private string clientSecret;
        private string[] scopes;
        private string tenantId;

        public ClientSecretAuthProvider(string appId, string[] scopes, string tenantId, string clientSecret)
        {
            this.appId = appId;
            this.clientSecret = clientSecret;
            this.scopes = scopes;
            this.tenantId = tenantId;

            this.msalClient = ConfidentialClientApplicationBuilder.Create(this.appId)
                .WithAuthority(AzureCloudInstance.AzurePublic, this.tenantId)
                .WithClientSecret(this.clientSecret)
                .Build();
        }

        public async Task<string> GetAuthorizationHeader()
        {
            try
            {
                var result = await this.msalClient
                    .AcquireTokenForClient(this.scopes)
                    .ExecuteAsync();
                return result.CreateAuthorizationHeader();
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Error getting access token: {exception.Message}");
                return null;
            }
        }

        public async Task AuthenticateRequestAsync(HttpRequestMessage requestMessage)
        {
            requestMessage.Headers.Add("Authorization", await GetAuthorizationHeader());
        }
    }
}