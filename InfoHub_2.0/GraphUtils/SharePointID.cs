namespace InfoHub.GraphUtils
{
    public class SharePointID
    {
        public static string SiteId_InfoHub =
            "livelabsmah.sharepoint.com,06e83ef0-42c7-41c5-a2ee-98e8fbf048ae,f00c9244-4a5e-4a46-b882-452b4f61bbbc";

        public static string ListId_ContactPersons = "795be841-c6dc-418a-b690-add878fd8ad6";
        public static string ListId_System = "6a90f083-7a54-4020-858d-8f4ea3da8cb8";
        public static string ListId_FAQ = "8322e7ec-d76d-4aaf-882b-c0f771a822dd";
        public static string ListId_Embedded = "d1737c38-c9e1-42fd-b191-eb6cab5b25c4";
        public static string ListId_CustomContent = "8f21d027-a7ab-4541-bc2a-b9b96313154e";
        public static string ListId_KnownIssues = "c7e35205-a1ce-44a4-a4f4-11bc0e4290bb";
        public static string ListId_LogEntries = "6a36ca0f-cc15-4236-9053-1e18dfb25eff";
        public static readonly string TenantId = "e0808df3-c991-424f-9cd6-50f43dee5689";

        public static bool GetStatus()
        {
            if (string.IsNullOrEmpty(ListId_System) ||
                string.IsNullOrEmpty(ListId_ContactPersons) ||
                string.IsNullOrEmpty(ListId_FAQ) ||
                string.IsNullOrEmpty(ListId_Embedded) ||
                string.IsNullOrEmpty(ListId_CustomContent) ||
                string.IsNullOrEmpty(ListId_KnownIssues) ||
                string.IsNullOrEmpty(ListId_LogEntries))
            {
                return false;
            }

            return true;
        }
    }
}