using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.IO;
using Twileloop.SpreadSheet.Factory.Base;

namespace Twileloop.SpreadSheet.MicrosoftExcel
{
    public partial class MicrosoftExcelDriver : ISpreadSheetDriver
    {
        private readonly MicrosoftExcelOptions config;
        private IWorkbook workbook;
        private ISheet sheet;

        public string DriverName => "MicrosoftExcel";

        public MicrosoftExcelDriver(MicrosoftExcelOptions config)
        {
            this.config = config;
        }

        public void InitialiseWorkbook()
        {
            if (File.Exists(config.FileLocation))
            {
                using (FileStream fileStream = new FileStream(config.FileLocation, FileMode.Open, FileAccess.ReadWrite))
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
            }
            else
            {
                workbook = new XSSFWorkbook();
            }
        }

        public void Dispose()
        {
            var xfile = new FileStream(config.FileLocation, FileMode.Create, FileAccess.Write);
            workbook.Write(xfile, false);
            workbook.Close();
            xfile.Close();
            GC.SuppressFinalize(this);
        }

        public string[] GetSheets()
        {
            var sheetTitles = new string[workbook.NumberOfSheets];
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                sheetTitles[i] = workbook.GetSheetName(i);
            }
            return sheetTitles;
        }

        public void OpenSheet(string sheetName)
        {
            sheet = workbook.GetSheet(sheetName);
        }

        public string GetActiveSheet()
        {
            var activeSheetIndex = workbook.ActiveSheetIndex;
            var activeSheetTitle = workbook.GetSheetName(activeSheetIndex);
            return activeSheetTitle;
        }

        public void CreateSheets(params string[] sheetNames)
        {
            foreach (string sheetName in sheetNames)
            {
                if (workbook.GetSheetIndex(sheetName) == -1)
                {
                    workbook.CreateSheet(sheetName);
                }
            }
        }

        public void SaveWorkbook()
        {
            using (FileStream fileStream = new FileStream(config.FileLocation, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream); // Write the workbook to the file
            }
        }
    }
}