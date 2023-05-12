using System;
using Twileloop.SpreadSheet.Factory.Configs;
using Twileloop.SpreadSheet.Factory.Services;

namespace Twileloop.SpreadSheet.Factory
{
    public static class SpreadSheetFactory
    {
        public static SpreadSheetService CreateSpreadSheetService(SpreadSheetKind spreadsheetKind, ServiceConfiguration configuration)
        {
            var spreadsheetService = new SpreadSheetService();
            switch (spreadsheetKind)
            {
                case SpreadSheetKind.MicrosoftExcel:
                    var excel = new MicrosoftExcelService(configuration as MicrosoftExcelConfiguration);
                    spreadsheetService.Reader = excel;
                    spreadsheetService.Writer = excel;
                    spreadsheetService.Controller = excel;
                    break;
                case SpreadSheetKind.GoogleSheet:
                    var sheet = new GoogleSheetService(configuration as GoogleSheetConfiguration);
                    spreadsheetService.Reader = sheet;
                    spreadsheetService.Writer = sheet;
                    spreadsheetService.Controller = sheet;
                    break;
                default:
                    throw new NotSupportedException("Unsupported service");
            }
            return spreadsheetService;
        }
    }
}
