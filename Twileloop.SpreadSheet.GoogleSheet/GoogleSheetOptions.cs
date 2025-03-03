using System;

namespace Twileloop.SpreadSheet.GoogleSheet
{

    public class GoogleSheetOptions
    {
        public GoogleSheetOptions(Uri sheetsURI, string applicationName, string jsonCredentialContent)
        {
            SheetsURI = sheetsURI;
            ApplicationName = applicationName;
            JsonCredentialContent = jsonCredentialContent;
        }

        public Uri SheetsURI { get; set; }
        public string ApplicationName { get; set; }
        public string JsonCredentialContent { get; set; }
    }
}