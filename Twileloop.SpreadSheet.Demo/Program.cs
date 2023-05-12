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


            //Step 4: Different Ways To Write Data
            using (googleSheetAccessor)
            {
                googleSheetAccessor.Controller.LoadSheet("Sheet1");

                //Write a single cell
                googleSheetAccessor.Writer.WriteCell(1, 1, "Country");
                googleSheetAccessor.Writer.WriteCell("C17", "Country");

                //Write a full row in bulk
                googleSheetAccessor.Writer.WriteRow(1, new string[] { "USA", "China", "Russia", "India" });
                googleSheetAccessor.Writer.WriteRow("A1", new string[] { "USA", "China", "Russia", "India" });

                //Write a full column in bulk
                googleSheetAccessor.Writer.WriteColumn(1, new string[] { "USA", "China", "Russia", "India" });
                googleSheetAccessor.Writer.WriteColumn("B22", new string[] { "USA", "China", "Russia", "India" });

                //Select an area and write a grid in bulk
                DataTable grid = new DataTable();
                grid.Columns.Add("Rank");
                grid.Columns.Add("Powerfull Militaries");

                grid.Rows.Add(1, "USA");
                grid.Rows.Add(2, "China");
                grid.Rows.Add(3, "Russia");
                grid.Rows.Add(4, "India");
                grid.Rows.Add(5, "France");

                googleSheetAccessor.Writer.WriteSelection(1, 1, grid);
                googleSheetAccessor.Writer.WriteSelection("D20", grid);
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
