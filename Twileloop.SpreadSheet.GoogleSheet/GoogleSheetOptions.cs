using System;

namespace Twileloop.SpreadSheet.GoogleSheet
{

    public class GoogleSheetOptions
    {
        public GoogleSheetOptions(Uri sheetsURI, string applicationName, string credential)
        {
            SheetsURI = sheetsURI;
            ApplicationName = applicationName;
            Credential = credential;
        }

        public Uri SheetsURI { get; set; }
        public string ApplicationName { get; set; }
        public string Credential { get; set; }
    }
}