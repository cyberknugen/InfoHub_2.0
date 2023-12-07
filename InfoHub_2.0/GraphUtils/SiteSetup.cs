using InfoHub.GraphUtils;
using InfoHub_2._0.Interop.TeamsSDK;
using Microsoft.AspNetCore.Components;
using Microsoft.Graph;
using Microsoft.TeamsFx;

namespace InfoHub.GraphUtils
{
    public class SiteSetup
    {


        public async Task<bool> SetUpSite(GraphServiceClient graphClient, MicrosoftTeams microsoftTeams)
        {
            try
            {
                var teamcontext = await microsoftTeams.GetTeamsContextAsync();
                var theChannel = teamcontext.Channel;
                var channelURL = theChannel.RelativeUrl;
                string theName = ExtractSegmentAfterSites(channelURL);
                

                //string theName = "InfoHub";
                var siteId = await FindSiteIDByName(theName, graphClient);

                return await SetAllListIDs(graphClient, siteId);
            }

            catch (Exception e)
            {
                Console.WriteLine($"Message: {e.Message}");
            }

            return false;
        }


        private async Task<bool> SetAllListIDs(GraphServiceClient graphClient, string siteId)
        {
            var lists = await graphClient.Sites[siteId].Lists.GetAsync();

            foreach (var list in lists.Value)
            {
                var name = list.DisplayName;
                var listId = list.Id;

                switch (GetSharePointModule(name))
                {
                    case SharePointModules.FAQ:
                        SharePointID.ListId_FAQ = listId;
                        break;

                    case SharePointModules.System:
                        SharePointID.ListId_System = listId;
                        break;


                    case SharePointModules.Embedded:
                        SharePointID.ListId_Embedded = listId;
                        break;

                    case SharePointModules.Custom:
                        SharePointID.ListId_CustomContent = listId;
                        break;

                    case SharePointModules.Kontaktpersoner:
                        SharePointID.ListId_ContactPersons = listId;
                        break;
                    /*
                case SharePointModules.Log:
                    SharePointID.Log = listId;
                    break;

                case SharePointModules.Known_issues:
                    SharePointID.Known_issues = listId;
                    break;
                    */
                    default:
                        break;
                }

            }


            return SharePointID.GetStatus();
        }


        static string ExtractSegmentAfterSites(string input)
        {
            string[] parts = input.Split('/');

            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] == "sites" && i < parts.Length - 1)
                {
                    return parts[i + 1];
                }
            }

            return null;
        }


        private SharePointModules? GetSharePointModule(string name)
        {
            if (Enum.TryParse<SharePointModules>(name, out var module))
            {
                return module;
            }

            return null;
        }


        public async Task<string> FindSiteIDByName(string name, GraphServiceClient graphClient)
        {

            try
            {
                var result = await graphClient.Sites.GetAsync((requestConfiguration) =>
                {
                    requestConfiguration.QueryParameters.Search = name;
                });

                return result.Value.ElementAt(0).Id;

            }
            catch (Exception e)
            {
                Console.WriteLine($"Message: {e.Message}");
            }
            return "";
        }

    }
}
