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
                FileLocation = @"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\Demo.xlsx"
            });

            var googleSheetsDriver = new GoogleSheetDriver(new GoogleSheetOptions
            {
                SheetsURI = new Uri("https://docs.google.com/spreadsheets/d/18roEDKYpgYfKDj0rQlnt7QC3b31Eb24DAoH0S4CiALQ/edit#gid=0"),
                Credential = @"D:\secrets.json"
            });

            //Step 2: Use that driver to build a spreadsheet accessor
            var excelAccessor = SpreadSheetFactory.CreateAccessor(excelDriver);
            var googleSheetAccessor = SpreadSheetFactory.CreateAccessor(googleSheetsDriver);




            //Read and write both spreadsheets at once
            using (excelAccessor)
            {
                using (googleSheetAccessor)
                {
                    //Step 1: Open both spreadsheets
                    excelAccessor.Controller.LoadSheet("Sheet1");
                    googleSheetAccessor.Controller.LoadSheet("Sheet1");

                    //Step 2: Read from excel
                    DataTable excelData = excelAccessor.Reader.ReadSelection("A1", "D10");

                    //Step 3: Then write it to Google Sheet
                    googleSheetAccessor.Writer.WriteSelection("C1", excelData);                    
                }
            }





            //Step 4: Different Ways To Write Data
            using (excelAccessor)
            {
                using (googleSheetAccessor)
                {
                    //Load prefered sheet
                    excelAccessor.Controller.LoadSheet("Major");

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
                    DataTable data7 = excelAccessor.Reader.ReadSelection(1, 1, 10, 2);
                    DataTable data8 = excelAccessor.Reader.ReadSelection("A1", "D10");

                    Console.Clear();

                    DrawDataTable(data7);
                    DrawDataTable(data8);

                    excelAccessor.Writer.WriteSelection("C1", data7);


                    //Load sheet
                    googleSheetAccessor.Controller.LoadSheet("Major");

                    //Read a single cell
                    data1 = googleSheetAccessor.Reader.ReadCell(1, 1);
                    data2 = googleSheetAccessor.Reader.ReadCell("A10");

                    //Read a full row in bulk
                    data3 = googleSheetAccessor.Reader.ReadRow(1);
                    data4 = googleSheetAccessor.Reader.ReadRow("C9");

                    //Read a full column in bulk
                    data5 = googleSheetAccessor.Reader.ReadColumn(1);
                    data6 = googleSheetAccessor.Reader.ReadColumn("D20");

                    //Select an area and extract data in bulk
                    data7 = googleSheetAccessor.Reader.ReadSelection(1, 1, 10, 2);
                    data8 = googleSheetAccessor.Reader.ReadSelection("A1", "D10");

                    Console.Clear();

                    DrawDataTable(data7);
                    DrawDataTable(data8);

                    googleSheetAccessor.Writer.WriteSelection("C1", data7);
                }
            }


           




            //Step 4: Different Ways To Write Data
            using (googleSheetAccessor)
            {
                googleSheetAccessor.Controller.LoadSheet("Major");

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
