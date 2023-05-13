using System.Diagnostics;
using Twileloop.SpreadSheet.Constructs;
using Twileloop.SpreadSheet.Factory;
using Twileloop.SpreadSheet.Formating;
using Twileloop.SpreadSheet.GoogleSheet;
using Twileloop.SpreadSheet.MicrosoftExcel;

namespace Twileloop.SpreadSheet.Demo
{
    public class Program
    {
        public static void Main(string[] args)
        {

            //Initialize your spreadsheet drivers
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions
            {
                FileLocation = @"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\Demo.xlsx"
            });

            var googleSheetsDriver = new GoogleSheetDriver(new GoogleSheetOptions
            {
                SheetsURI = new Uri("https://docs.google.com/spreadsheets/d/18roEDKYpgYfKDj0rQlnt7QC3b31Eb24DAoH0S4CiALQ/edit#gid=0"),
                Credential = @"D:\secrets.json"
            });


            //Use that driver to build a spreadsheet accessor
            var excelAccessor = SpreadSheetFactory.CreateAccessor(excelDriver);
            var googleSheetAccessor = SpreadSheetFactory.CreateAccessor(googleSheetsDriver);

            using (excelAccessor)
            {
                using (googleSheetAccessor)
                {
                    excelAccessor.Controller.LoadSheet("Major");
                    googleSheetAccessor.Controller.LoadSheet("Major");


                    var titleFormat = new Formatting
                    {
                        //Text formatting
                        TextFormating = new TextFormating
                        {
                            Bold = false,
                            Italic = true,
                            Underline = false,
                            Size = 15,
                            HorizontalAlignment = HorizontalAllignment.RIGHT,
                            VerticalAlignment = VerticalAllignment.BOTTOM,
                            Font = "Impact",
                            Color = System.Drawing.Color.White,
                        },
                        //Cell formatting
                        CellFormating = new CellFormating
                        {
                            BackgroundColor = System.Drawing.Color.IndianRed
                        },
                        //Border formatting
                        BorderFormating = new BorderFormating
                        {
                            TopBorder = true,
                            LeftBorder = true,
                            RightBorder = true,
                            BottomBorder = true,
                            BorderType = BorderType.SOLID,
                            Thickness = 5
                        }
                    };



                    excelAccessor.Writer.ApplyFormatting(1, 1, 10, 4, titleFormat);
                    googleSheetAccessor.Writer.ApplyFormatting(1, 1, 10, 4, titleFormat);



                    //var activeExcelSheet = excelAccessor.Controller.GetActiveSheet();
                    //var googleSheetSheet = googleSheetAccessor.Controller.GetActiveSheet();

                    //var allExcelSheets = excelAccessor.Controller.GetSheets();
                    //var allGoogleSheetSheet = googleSheetAccessor.Controller.GetSheets();

                    //excelAccessor.Controller.CreateSheets("Sheet1", "Sheet2", "Sheet3");
                    //googleSheetAccessor.Controller.CreateSheets("Sheet1", "Sheet2", "Sheet3");
                }
            }

            string filePath = @"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\Demo.xlsx";

            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", $"\"{filePath}\"");

        }
    }
}
