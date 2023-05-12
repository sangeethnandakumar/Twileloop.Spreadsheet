using Spectre.Console;
using System.Data;
using Twileloop.SpreadSheet.Factory;
using Twileloop.SpreadSheet.Factory.Configs;

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


            //Make Microsoft Excel and GoogleSheet services
            var excel = SpreadSheetFactory.CreateSpreadSheetService(SpreadSheetKind.MicrosoftExcel, new MicrosoftExcelConfiguration
            {
                FileLocation = @"Demo.xlsx",
            });

            //Operate Microsoft Excel
            using (excel)
            {
                excel.Controller.LoadWorkbook("Sheet1");
                //Read a cell by address or row+column
                excel.Reader.ReadCell(1, 1);
                excel.Reader.ReadCell("A1");
                //Read full rows
                var excelRows = excel.Reader.ReadRow(1);
                //Read full columns
                var excelColumns = excel.Reader.ReadColumn(1);
                //Read a selection as grid
                var excelGrid = excel.Reader.ReadSelection(1, 1, 5, 5);
                excelGrid = excel.Reader.ReadSelection("C7", "G9");
                DrawDataTable(excelGrid);

                for (var row = 1; row <= countries.Count; row++)
                {
                    excel.Writer.WriteRow(row, row.ToString(), countries[row]);
                }
            }






            var googleSheets = SpreadSheetFactory.CreateSpreadSheetService(SpreadSheetKind.GoogleSheet, new GoogleSheetConfiguration
            {
                SheetsURI = new Uri("https://docs.google.com/spreadsheets/d/1V0w0bECUI4c0bUgyz11RLrIpkhoxlXhPtkw6mbNqws8/edit#gid=1048112514"),
                Credential = "secret.json",
            });

            //Operate Microsoft Excel
            using (googleSheets)
            {
                googleSheets.Controller.LoadWorkbook("<SHEET_NAME>");
                for (var row = 1; row <= countries.Count; row++)
                {
                    googleSheets.Writer.WriteRow(row, row.ToString(), countries[row]);
                }
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
