using Spectre.Console;
using System.Data;
using Twileloop.SpreadSheet.Factory;
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


            //Step 1: Create a driver
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions
            {
                FileLocation = @"Demo.xlsx"
            });

            //Step 2: Use that driver to build a spreadsheet accessor
            var accessor = SpreadSheetFactory.CreateAccessor(excelDriver);

            //Step 3: Now this accessor can Read/Write and Control spreadsheet. Let's open Sheet1
            using (accessor)
            {
                //Control spreadsheet
                accessor.Controller.LoadSheet("Sheet1");

                //Read something
                var seventhRow = accessor.Reader.ReadRow(7);

                //Write it as column downwards
                accessor.Writer.WriteColumn(2, seventhRow);
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
