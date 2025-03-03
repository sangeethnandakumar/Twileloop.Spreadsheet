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
            if(File.Exists(filePath))
                File.Delete(filePath);
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions(filePath));
            ISpreadSheetAdapter excelAdapter = SpreadSheetFactory.CreateAdapter(excelDriver);

            // Google Sheet
            var sheetsURI = new Uri("https://docs.google.com/spreadsheets/d/1YWqL4_jmGhtpj--ZBLRe598w7IXDCvzL0UWHU_wZMqU/edit?gid=0#gid=0");
            var sheetName = "MySheet";
            var jsonCredentialContent = File.ReadAllText("secrets.json");
            var bulkUpdate = true;
            var googleSheet = new GoogleSheetDriver(new GoogleSheetOptions(sheetsURI, sheetName, jsonCredentialContent, bulkUpdate));
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
            //Process.Start(new ProcessStartInfo
            //{
            //    FileName = sheetsURI.ToString(),
            //    UseShellExecute = true
            //});
        }


        private static void NewSpreadsheet(ISpreadSheetAdapter adapter)
        {
            using (adapter)
            {
                Console.WriteLine($"Driver: {adapter.DriverName} | Initialising");
                adapter.Controller.InitialiseWorkbook();

                Console.WriteLine($"Driver: {adapter.DriverName} | Creating sheets");
                adapter.Controller.CreateSheets("A");

                Console.WriteLine($"Driver: {adapter.DriverName} | Opening sheet");
                adapter.Controller.OpenSheet("A");

                Console.WriteLine($"Driver: {adapter.DriverName} | Writing cells");
                adapter.Writer.WriteCell("A1", "Write");
                adapter.Writer.WriteCell((1, 2), "Individual");
                adapter.Writer.WriteCell(("C1"), "Cells");

                var impactStyle = new StyleBuilder()
                    .WithFont("Impact")
                    .WithTextColor(Color.AliceBlue)
                    .WithTextAllignment(HorizontalTxtAlignment.CENTER, VerticalTxtAlignment.BOTTOM)
                    .WithBackgroundColor(Color.Black)
                    .Build();

                Console.WriteLine($"Driver: {adapter.DriverName} | Merging cells");
                adapter.Writer.MergeCells("A2", "E2"); //A2:E2

                Console.WriteLine($"Driver: {adapter.DriverName} | Writing cell");
                adapter.Writer.WriteCell("A2", "This is just a very very long description as an example", new StyleBuilder().Bold().Italic().Underline().Build());

                
                Console.WriteLine($"Driver: {adapter.DriverName} | Writing rows");
                adapter.Writer.WriteRow("A3", new[] { "Col 1", "Col 2", "Col 3", "Col 4" }, impactStyle);
                adapter.Writer.WriteRow("A4", new[] { "Col 1", "Col 2", "Col 3", "Col 4" }, impactStyle);
                adapter.Writer.WriteRow("A5", new[] { "Col 1", "Col 2", "Col 3", "Col 4" }, impactStyle);

                Console.WriteLine($"Driver: {adapter.DriverName} | Writing columns");
                adapter.Writer.WriteColumn("A7", new[] { "Row 1", "Row 2", "Row 3", "Row 4" }, impactStyle);
                adapter.Writer.WriteColumn("B7", new[] { "Row 1", "Row 2", "Row 3", "Row 4" }, impactStyle);
                adapter.Writer.WriteColumn("C7", new[] { "Row 1", "Row 2", "Row 3", "Row 4" }, impactStyle);
                adapter.Writer.WriteColumn("D7", new[] { "Row 1", "Row 2", "Row 3", "Row 4" }, impactStyle);

                Console.WriteLine($"Driver: {adapter.DriverName} | Creating styles");
                var headingStyle = new StyleBuilder()
                    .Bold()
                    .WithFontSize(18)
                    .WithFont("Arial")
                    .WithTextColor(Color.Blue)
                    .WithTextAllignment(HorizontalTxtAlignment.LEFT, VerticalTxtAlignment.TOP)
                    .WithBackgroundColor(Color.LightBlue)
                    .Build();

                var myStyle = new StyleBuilder()
                    .Italic()
                    .Underline()
                    .WithTextColor(Color.White)
                    .WithTextAllignment(HorizontalTxtAlignment.RIGHT, VerticalTxtAlignment.MIDDLE)
                    .WithBackgroundColor(Color.DeepPink)
                    .Build();
                

                Console.WriteLine($"Driver: {adapter.DriverName} | Creating table");
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

                Console.WriteLine($"Driver: {adapter.DriverName} | Writing table");
                adapter.Writer.WriteTable("A12", table, myStyle);

                Console.WriteLine($"Driver: {adapter.DriverName} | Applying border");
                adapter.Writer.ApplyBorder("A12", "E21", new BorderStyling
                {
                    TopBorder = true,
                    LeftBorder = true,
                    RightBorder = true,
                    BottomBorder = true,
                    BorderType = BorderType.SOLID,
                    BorderColor = Color.OrangeRed,
                    Thickness = BorderThickness.Thick
                });

                Console.WriteLine($"Driver: {adapter.DriverName} | Applying styling");
                adapter.Writer.ApplyStyling("A1", "E1", headingStyle);
                adapter.Writer.ApplyStyling("A3", "D10", new StyleBuilder().WithTextAllignment(HorizontalTxtAlignment.CENTER, VerticalTxtAlignment.MIDDLE).Build());

                Console.WriteLine($"Driver: {adapter.DriverName} | Resizing row");
                adapter.Writer.ResizeRow("A1", 50);

                Console.WriteLine($"Driver: {adapter.DriverName} | Resizing column");
                adapter.Writer.ResizeColumn("A1", 40);
                adapter.Writer.ResizeColumn("B1", 40);
                adapter.Writer.ResizeColumn("C1", 40);
                adapter.Writer.ResizeColumn("D1", 40);
                adapter.Writer.ResizeColumn("E1", 40);

                Console.WriteLine($"Driver: {adapter.DriverName} | Saving workbook");
                adapter.Controller.SaveWorkbook();
            }
        }


    }
}
