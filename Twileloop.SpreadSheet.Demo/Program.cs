using System.Data;
using System.Diagnostics;
using System.Drawing;
using Twileloop.SpreadSheet.Factory;
using Twileloop.SpreadSheet.GoogleSheet;
using Twileloop.SpreadSheet.MicrosoftExcel;
using Twileloop.SpreadSheet.Styling;

namespace Twileloop.SpreadSheet.Demo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // Excel
            string filePath = @"Demo.xlsx";
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions(filePath));
            ISpreadSheetAdapter excelAdapter = SpreadSheetFactory.CreateAdapter(excelDriver);

            // Google Sheet
            var sheetsURI = new Uri("https://docs.google.com/spreadsheets/d/1YWqL4_jmGhtpj--ZBLRe598w7IXDCvzL0UWHU_wZMqU/edit?gid=0#gid=0");
            var sheetName = "MySheet";
            var jsonCredentialContent = File.ReadAllText("secrets.json");
            var googleSheet = new GoogleSheetDriver(new GoogleSheetOptions(sheetsURI, sheetName, jsonCredentialContent));
            ISpreadSheetAdapter gsheetAdapter = SpreadSheetFactory.CreateAdapter(googleSheet);

            // Common API
            Parallel.Invoke(() => NewSpreadsheet(excelAdapter), () => NewSpreadsheet(gsheetAdapter));

            // Open Excel file
            OpenExcelFile(filePath);

            // Open Google Sheet in browser
            OpenGoogleSheetInBrowser(sheetsURI);
        }

        private static void OpenExcelFile(string filePath)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            });
        }

        private static void OpenGoogleSheetInBrowser(Uri sheetsURI)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = sheetsURI.ToString(),
                UseShellExecute = true
            });
        }


        private static void NewSpreadsheet(ISpreadSheetAdapter adapter)
        {
            adapter.Controller.InitialiseWorkbook();
            adapter.Controller.CreateSheets("A", "B");
            adapter.Controller.OpenSheet("Sheet 1");

            var headingStyle = new StyleBuilder()
                .Bold()
                .WithFontSize(18)
                .WithFont("Arial")
                .WithTextColor(Color.Blue)
                .WithHorizontalAlignment(HorizontalTxtAlignment.CENTER)
                .WithVerticalAlignment(VerticalTxtAlignment.MIDDLE)
                .WithBackgroundColor(Color.LightBlue)
                .Build();

            var myStyle = new StyleBuilder()
                .Italic()
                .Underline()
                .WithTextColor(Color.White)
                .WithBackgroundColor(Color.Green)
                .Build();

            var table = new DataTable();
            table.Columns.Add("ID");
            table.Columns.Add("Name");
            table.Columns.Add("Age");
            table.Columns.Add("City");
            table.Columns.Add("Salary");
            table.Rows.Add(1, "John Doe", 28, "New York", 55000);
            table.Rows.Add(2, "Alice Smith", 34, "Los Angeles", 62000);
            table.Rows.Add(3, "Bob Johnson", 41, "Chicago", 72000);
            table.Rows.Add(4, "Emily Davis", 25, "Houston", 48000);
            table.Rows.Add(5, "Michael Brown", 37, "Phoenix", 67000);
            table.Rows.Add(6, "Sarah Wilson", 30, "Philadelphia", 59000);
            table.Rows.Add(7, "David Lee", 45, "San Antonio", 75000);
            table.Rows.Add(8, "Laura White", 27, "San Diego", 51000);
            table.Rows.Add(9, "James Green", 33, "Dallas", 60000);
            table.Rows.Add(10, "Emma Harris", 29, "San Francisco", 68000);

            adapter.Writer.WriteTable("A1", table, myStyle);
            adapter.Writer.ApplyBorder("A1", "E11", new BorderStyling
            {
                TopBorder = true,
                LeftBorder = true,
                RightBorder = true,
                BottomBorder = true,
                BorderType = BorderType.SOLID,
                BorderColor = Color.OrangeRed,
                Thickness = BorderThickness.Thick
            });

            adapter.Writer.ApplyStyling("A1", "E1", headingStyle);
            adapter.Writer.MergeCells("A2", "E2");
            adapter.Writer.WriteCell("A2", "This is just a very very long description as an example", new StyleBuilder().Bold().Italic().Underline().Build());
            adapter.Writer.ResizeRow("A1", 50);
            adapter.Writer.ResizeColumn("E1", 50);
            adapter.Controller.SaveWorkbook();
        }
    }
}
