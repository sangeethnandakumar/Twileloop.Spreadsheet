using System.Diagnostics;
using Twileloop.SpreadSheet.Constructs;
using Twileloop.SpreadSheet.Factory;
using Twileloop.SpreadSheet.Formating;
using Twileloop.SpreadSheet.GoogleSheet;
using Twileloop.SpreadSheet.MicrosoftExcel;
using Twileloop.Storage.GoogleDrive;

namespace Twileloop.SpreadSheet.Demo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            GoogleDriveDemo().Wait();


            //Initialize your spreadsheet drivers
            var excelDriver = new MicrosoftExcelDriver(new MicrosoftExcelOptions
            {
                FileLocation = @"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\Demo.xlsx"
            });

            var googleSheetsDriver = new GoogleSheetDriver(new GoogleSheetOptions
            {
                SheetsURI = new Uri("https://docs.google.com/spreadsheets/d/1BZTxOnwBMwcPXSJp4etagUsbkwr7N8K6T5g98V9wlvA/edit#gid=1354136408"),
                Credential = @"secrets.json"
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



                    var activeExcelSheet = excelAccessor.Controller.GetActiveSheet();
                    var googleSheetSheet = googleSheetAccessor.Controller.GetActiveSheet();

                    var allExcelSheets = excelAccessor.Controller.GetSheets();
                    var allGoogleSheetSheet = googleSheetAccessor.Controller.GetSheets();

                    excelAccessor.Controller.CreateSheets("Sheet1", "Sheet2", "Sheet3");
                    googleSheetAccessor.Controller.CreateSheets("Sheet1", "Sheet2", "Sheet3");
                }
            }

            string filePath = @"C:\Users\Sangeeth Nandakumar\OneDrive\Desktop\Demo.xlsx";

            Process.Start(@"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE", $"\"{filePath}\"");

        }

        private static async Task GoogleDriveDemo()
        {
            // Initialize Google Drive service
            var googleDriveService = new GoogleDriveService("drive.json", "BestSellerScrapper");

            // List all files and directories
            var items = googleDriveService.GetAllFilesAndDirectoriesAsync().Result.ToList();
            foreach (var item in items)
            {
                Console.WriteLine($"{item.Name} ({item.Id})");
            }

            //Make them listed for my account in GoogleDrive
            //foreach (var item in items)
            //{
            //    Console.WriteLine("Sharing with another user...");
            //    //await googleDriveService.ShareFileWithSpecificUsers(item.Id, new List<string> { "sangeethnandakumarofficial@gmail.com" });

            //    Console.WriteLine("Generating a sharable link");
            //    var link = await googleDriveService.GenerateShareableLink(item.Id);
            //}

            //// Create a new directory
            //await googleDriveService.CreateDirectoryAsync("NewDirectory");

            // Upload a file
            await googleDriveService.UploadFileAsync("demofile.pdf", "application/pdf", progress: (total, downloaded) =>
            {
                decimal percentage = Math.Round((decimal)downloaded/(decimal)total * 100,1);
                Console.WriteLine($"[{percentage}%] Uploading {downloaded} of {total} bytes...");
            }, chunkSizeInMB: 1);

            items = googleDriveService.GetAllFilesAndDirectoriesAsync().Result.ToList();

            // Rename the file
            var file = items.First(i => i.Name == "demofile.pdf");
            await googleDriveService.RenameFileAsync(file.Id, "renamed_demofile.jpg");

            // Move the file
            var directory = items.First(i => i.Name == "NewDirectory");
            await googleDriveService.MoveFileAsync(file.Id, directory.Id);

            // Copy the file
            await googleDriveService.CopyFileAsync(file.Id, directory.Id);

            // Download the file
            await googleDriveService.DownloadFileAsync(file.Id, "renamed_demofile.txt", progress: (total, downloaded) =>
            {
                Console.WriteLine($"Downloaded {downloaded} of {total} bytes");
            });

            // Delete the file
            await googleDriveService.DeleteFileAsync(file.Id);
        }
    }
}
