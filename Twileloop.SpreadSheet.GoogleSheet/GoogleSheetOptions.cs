using System;

namespace Twileloop.SpreadSheet.GoogleSheet
{

    public class GoogleSheetOptions
    {
        public GoogleSheetOptions(Uri sheetsURI, string applicationName, string jsonCredentialContent, bool bulkUpdate)
        {
            SheetsURI = sheetsURI;
            ApplicationName = applicationName;
            JsonCredentialContent = jsonCredentialContent;
            BulkUpdate = bulkUpdate;
        }

        public Uri SheetsURI { get; set; }
        public string ApplicationName { get; set; }
        public string JsonCredentialContent { get; set; }
        public bool BulkUpdate { get; set; }
    }
}