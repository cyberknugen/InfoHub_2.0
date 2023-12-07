using Azure.Core;
using static System.Diagnostics.Debug;
using InfoHub_2._0.Components;
using Microsoft.Graph;
using Microsoft.TeamsFx;
using static System.Formats.Asn1.AsnWriter;
using Azure.Identity;
using InfoHub_2._0;
using Microsoft.Graph.Models.ExternalConnectors;
using Microsoft.Extensions.Configuration;


namespace InfoHub.GraphUtils
{
	public class CommonHelper
	{	//TODO fixa enligt nya hasPermission
		

        public static async Task<bool> HasPermission(string scope, TeamsUserCredential teamsUserCredential, IConfiguration configuration)
        {
            try
            {
                var tokenCredential = await GetOnBehalfOfCredential(teamsUserCredential, configuration);
                await tokenCredential.GetTokenAsync(new TokenRequestContext(new string[] { scope }), new CancellationToken());

                return true;
            }
            catch (Exception e)
            {
                WriteLine("Does not have permission: " + e.Message);

            } 
            return false;
        }

        public static async Task<OnBehalfOfCredential> GetOnBehalfOfCredential(TeamsUserCredential teamsUserCredential, IConfiguration configuration)
        {
            var config = configuration.Get<ConfigOptions>();
            var tenantId = SharePointID.TenantId;
            AccessToken ssoToken = await teamsUserCredential.GetTokenAsync(new TokenRequestContext(null), new CancellationToken());
            return new OnBehalfOfCredential(
                tenantId,
                config.TeamsFx.Authentication.ClientId,
                config.TeamsFx.Authentication.ClientSecret,
                ssoToken.Token
            );
        }

        public static GraphServiceClient GetGraphServiceClient(TokenCredential tokenCredential, string scope)
        {
            var client = new GraphServiceClient(tokenCredential, new string[] { scope });
            return client;
        }
    }
}
