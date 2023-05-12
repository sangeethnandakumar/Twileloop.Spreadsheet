using Spectre.Console;
using System.Data;
using Twileloop.SpreadSheet.Factory;
using Twileloop.SpreadSheet.GoogleSheet;
using Twileloop.SpreadSheet.MicrosoftExcel;

namespace Twileloop.SpreadSheet.Demo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //Biggest economies in the world
            var countries = new Dictionary<int, string>();
            countries.Add(1, "United States of America");
            countries.Add(2, "China");
            countries.Add(3, "Japan");
            countries.Add(4, "Germany");
            countries.Add(5, "India");
            countries.Add(6, "United Kingdom");
            countries.Add(7, "France");
            countries.Add(8, "Italy");
            countries.Add(9, "Canada");
            countries.Add(10, "South Korea");


            //Step 1: Initialize your spreadsheet drivers
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions
            {
                FileLocation = "<YOUR_EXCEL_FILE_LOCATION>"
            });

            var googleSheetsDriver = new GoogleSheetDriver(new GoogleSheetOptions
            {
                SheetsURI = new Uri("<YOUR_GOOGLE_SHEETS_URL>"),
                Credential = "secrets.json"
            });

            //Step 2: Use that driver to build a spreadsheet accessor
            var excelAccessor = SpreadSheetFactory.CreateAccessor(excelDriver);
            var googleSheetAccessor = SpreadSheetFactory.CreateAccessor(googleSheetsDriver);

            //Step 3: Now this accessor can Read/Write and Control spreadsheet. Let's open Sheet1
           
            
            //Step 4: Different Ways To Read Data
            using (excelAccessor)
            {
                excelAccessor.Controller.LoadSheet("Sheet1");

                //Read a single cell
                string data1 = excelAccessor.Reader.ReadCell(1, 1);
                string data2 = excelAccessor.Reader.ReadCell("A10");

                //Read a full row in bulk
                string[] data3 = excelAccessor.Reader.ReadRow(1);
                string[] data4 = excelAccessor.Reader.ReadRow("C9");

                //Read a full column in bulk
                string[] data5 = excelAccessor.Reader.ReadColumn(1);
                string[] data6 = excelAccessor.Reader.ReadColumn("D20");

                //Select an area and extract data in bulk
                DataTable data7 = excelAccessor.Reader.ReadSelection(1, 1, 10, 10);
                DataTable data8 = excelAccessor.Reader.ReadSelection("A1", "J10");
            }



        }


        public static void DrawDataTable(DataTable dataTable)
        {
            var table = new Spectre.Console.Table();
            foreach (DataColumn column in dataTable.Columns)
            {
                table.AddColumn(column.ColumnName);
            }
            foreach (DataRow row in dataTable.Rows)
            {
                var rowData = row.ItemArray.Select(cell => cell.ToString()).ToArray();
                table.AddRow(rowData);
            }
            Spectre.Console.AnsiConsole.Render(table);
        }

    }
}
