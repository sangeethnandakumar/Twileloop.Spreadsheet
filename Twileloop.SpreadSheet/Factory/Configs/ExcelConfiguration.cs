using System;
using System.Linq;
using System.Security.Policy;

namespace Twileloop.SpreadSheet.Factory.Configs
{
    public class ServiceConfiguration
    {
    }

    public class MicrosoftExcelConfiguration : ServiceConfiguration
    {
        public string FileLocation { get; set; }
    }

    public class GoogleSheetConfiguration : ServiceConfiguration
    {
        public Uri SheetsURI { get; set; }
        public string ApplicationName { get; set; }
        public string Credential { get; set; }
    }
}