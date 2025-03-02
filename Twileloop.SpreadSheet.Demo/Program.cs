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
            string filePath = @"Demo.xlsx";
            File.Delete(filePath);

            //GoogleDriveDemo().Wait();

            // ----| MICROSOFT EXCEL |------------------------------------------------------------------

            SpreadSheetAccessor excelAccessor;
            SpreadsheetStyling headingStyle, myStyle;
            DataTable table;
            //WorkingWithExcel(out excelAccessor, out headingStyle, out myStyle, out table);

            // ----| GOOGLE SHEET |------------------------------------------------------------------

            //Microsoft Excel

            var sheetsURI = new Uri("https://docs.google.com/spreadsheets/d/1YWqL4_jmGhtpj--ZBLRe598w7IXDCvzL0UWHU_wZMqU/edit?gid=0#gid=0");
            var sheetName = "MySheet";
            var credential = @"secrets.json";

            var googleSheet = new GoogleSheetDriver(new GoogleSheetOptions(sheetsURI, sheetName, credential));
            var googleSheetAccessor = SpreadSheetFactory.CreateAccessor(googleSheet);

            //Creates new or loads an existsing workbook
            googleSheetAccessor.Controller.InitialiseWorkbook();

            //Create new sheets
            googleSheetAccessor.Controller.CreateSheets("Sheet 1", "Sheet 2");

            //Open 1st sheet
            googleSheetAccessor.Controller.OpenSheet("Sheet 1");

            //Write to Cell
            googleSheetAccessor.Writer.WriteCell("A1", "Write");
            googleSheetAccessor.Writer.WriteCell((1, 2), "Individual"); //Supports row, col numbers also
            googleSheetAccessor.Writer.WriteCell(("C1"), "Cells");

            //Write as Rows
            googleSheetAccessor.Writer.WriteRow("A3", ["Col 1", "Col 2", "Col 3", "Col 4"]);
            googleSheetAccessor.Writer.WriteRow("A4", ["Col 1", "Col 2", "Col 3", "Col 4"]);
            googleSheetAccessor.Writer.WriteRow("A5", ["Col 1", "Col 2", "Col 3", "Col 4"]);


            //Write as Cols
            googleSheetAccessor.Writer.WriteColumn("A7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            googleSheetAccessor.Writer.WriteColumn("B7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            googleSheetAccessor.Writer.WriteColumn("C7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            googleSheetAccessor.Writer.WriteColumn("D7", ["Row 1", "Row 2", "Row 3", "Row 4"]);

            //Create a style
            var headingStyle1 = new StyleBuilder()
                .Bold()
                .WithFontSize(18)
                .WithFont("Arial")
                .WithTextColor(Color.Blue)
                .WithHorizontalAlignment(HorizontalTxtAlignment.CENTER)
                .WithVerticalAlignment(VerticalTxtAlignment.MIDDLE)
                .WithBackgroundColor(Color.LightBlue)
                .Build();

            var myStyle1 = new StyleBuilder()
                .Italic()
                .Underline()
                .WithTextColor(Color.White)
                .WithBackgroundColor(Color.Green)
                .Build();

            //Table
            var table1 = new DataTable();
            table1.Columns.Add("ID");
            table1.Columns.Add("Name");
            table1.Columns.Add("Age");
            table1.Columns.Add("City");
            table1.Columns.Add("Salary");
            table1.Rows.Add(1, "John Doe", 28, "New York", 55000);
            table1.Rows.Add(2, "Alice Smith", 34, "Los Angeles", 62000);
            table1.Rows.Add(3, "Bob Johnson", 41, "Chicago", 72000);
            table1.Rows.Add(4, "Emily Davis", 25, "Houston", 48000);
            table1.Rows.Add(5, "Michael Brown", 37, "Phoenix", 67000);
            table1.Rows.Add(6, "Sarah Wilson", 30, "Philadelphia", 59000);
            table1.Rows.Add(7, "David Lee", 45, "San Antonio", 75000);
            table1.Rows.Add(8, "Laura White", 27, "San Diego", 51000);
            table1.Rows.Add(9, "James Green", 33, "Dallas", 60000);
            table1.Rows.Add(10, "Emma Harris", 29, "San Francisco", 68000);

            googleSheetAccessor.Writer.WriteTable("A12", table1, myStyle1);

            googleSheetAccessor.Writer.ApplyBorder("A12", "E21", new BorderStyling
            {
                TopBorder = true,
                LeftBorder = true,
                RightBorder = true,
                BottomBorder = true,
                BorderType = BorderType.SOLID,
                BorderColor = Color.OrangeRed,
                Thickness = BorderThickness.Thick
            });

            //More borders
            googleSheetAccessor.Writer.ApplyBorder("A3", "D5", new BorderStyling
            {
                TopBorder = true,
                LeftBorder = true,
                RightBorder = true,
                BottomBorder = true,
                BorderType = BorderType.DASHED,
                BorderColor = Color.OrangeRed,
                Thickness = BorderThickness.DoubleLined
            });

            //Apply selection styles
            googleSheetAccessor.Writer.ApplyStyling("A1", "C1", headingStyle1);

            //Merge few cells
            googleSheetAccessor.Writer.MergeCells("A2", "E2");
            googleSheetAccessor.Writer.WriteCell("A2", "This is just a very very long description as an example", new StyleBuilder()
                .Bold()
                .Italic()
                .Underline()
                .Build());

            //Resize a specific column
            googleSheetAccessor.Writer.ResizeRow("A1", 50);
            googleSheetAccessor.Writer.ResizeColumn("D1", 50);

            //Save file
            googleSheetAccessor.Controller.SaveWorkbook();

            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", $"\"{filePath}\"");
        }

        private static void WorkingWithExcel(out SpreadSheetAccessor excelAccessor, out SpreadsheetStyling headingStyle, out SpreadsheetStyling myStyle, out DataTable table)
        {
            //Microsoft Excel
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions(@"Demo.xlsx"));
            excelAccessor = SpreadSheetFactory.CreateAccessor(excelDriver);

            //Creates new or loads an existsing workbook
            excelAccessor.Controller.InitialiseWorkbook();

            //Create new sheets
            excelAccessor.Controller.CreateSheets("Sheet 1", "Sheet 2");

            //Open 1st sheet
            excelAccessor.Controller.OpenSheet("Sheet 1");

            //Write to Cell
            excelAccessor.Writer.WriteCell("A1", "Write");
            excelAccessor.Writer.WriteCell((1, 2), "Individual"); //Supports row, col numbers also
            excelAccessor.Writer.WriteCell(("C1"), "Cells");

            //Write as Rows
            excelAccessor.Writer.WriteRow("A3", ["Col 1", "Col 2", "Col 3", "Col 4"]);
            excelAccessor.Writer.WriteRow("A4", ["Col 1", "Col 2", "Col 3", "Col 4"]);
            excelAccessor.Writer.WriteRow("A5", ["Col 1", "Col 2", "Col 3", "Col 4"]);


            //Write as Cols
            excelAccessor.Writer.WriteColumn("A7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            excelAccessor.Writer.WriteColumn("B7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            excelAccessor.Writer.WriteColumn("C7", ["Row 1", "Row 2", "Row 3", "Row 4"]);
            excelAccessor.Writer.WriteColumn("D7", ["Row 1", "Row 2", "Row 3", "Row 4"]);

            //Create a style
            headingStyle = new StyleBuilder()
                .Bold()
                .WithFontSize(18)
                .WithFont("Arial")
                .WithTextColor(Color.Blue)
                .WithHorizontalAlignment(HorizontalTxtAlignment.CENTER)
                .WithVerticalAlignment(VerticalTxtAlignment.MIDDLE)
                .WithBackgroundColor(Color.LightBlue)
                .Build();
            myStyle = new StyleBuilder()
                .Italic()
                .Underline()
                .WithTextColor(Color.White)
                .WithBackgroundColor(Color.Green)
                .Build();

            //Table
            table = new DataTable();
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

            excelAccessor.Writer.WriteTable("A12", table, myStyle);

            excelAccessor.Writer.ApplyBorder("A12", "E21", new BorderStyling
            {
                TopBorder = true,
                LeftBorder = true,
                RightBorder = true,
                BottomBorder = true,
                BorderType = BorderType.SOLID,
                BorderColor = Color.OrangeRed,
                Thickness = BorderThickness.Thick
            });

            //More borders
            excelAccessor.Writer.ApplyBorder("A3", "D5", new BorderStyling
            {
                TopBorder = true,
                LeftBorder = true,
                RightBorder = true,
                BottomBorder = true,
                BorderType = BorderType.DASHED,
                BorderColor = Color.OrangeRed,
                Thickness = BorderThickness.DoubleLined
            });

            //Apply selection styles
            excelAccessor.Writer.ApplyStyling("A1", "C1", headingStyle);

            //Merge few cells
            excelAccessor.Writer.MergeCells("A2", "E2");
            excelAccessor.Writer.WriteCell("A2", "This is just a very very long description as an example", new StyleBuilder()
                .Bold()
                .Italic()
                .Underline()
                .Build());

            //Resize a specific column
            excelAccessor.Writer.ResizeRow("A1", 50);
            excelAccessor.Writer.ResizeColumn("D1", 50);

            //Save file
            excelAccessor.Controller.SaveWorkbook();
        }
    }
}
