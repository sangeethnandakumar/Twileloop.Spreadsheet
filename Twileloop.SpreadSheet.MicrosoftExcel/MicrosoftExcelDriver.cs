using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Twileloop.SpreadSheet.Constructs;
using Twileloop.SpreadSheet.Factory.Abstractions;
using Twileloop.SpreadSheet.Factory.Base;
using Twileloop.SpreadSheet.Formating;

namespace Twileloop.SpreadSheet.MicrosoftExcel
{
    public class MicrosoftExcelDriver : ISpreadSheetDriver
    {
        private readonly MicrosoftExcelOptions config;
        private IWorkbook workbook;
        private ISheet sheet;

        public MicrosoftExcelDriver(MicrosoftExcelOptions config)
        {
            this.config = config;
        }

        private void ValidatePrerequisites()
        {
            if (workbook is null)
            {
                throw new IOException($"Failed to load the workbook at '{config.FileLocation}'");
            }
            if (sheet is null)
            {
                throw new IOException($"Failed to load the required sheet. First sheet was '{workbook.GetSheetName(workbook.ActiveSheetIndex)}'");
            }
        }

        public string ReadCell(int row, int column)
        {
            ValidatePrerequisites();
            IRow excelRow = sheet.GetRow(row - 1);
            if (excelRow is not null)
            {
                ICell cell = excelRow.GetCell(column - 1);
                return cell?.ToString();
            }
            return null;
        }

        public string ReadCell(string address)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(address);
            return ReadCell(cellReference.Row + 1, cellReference.Col + 1); // Adjust row and column index
        }

        public string[] ReadColumn(int columnIndex)
        {
            ValidatePrerequisites();
            var columnData = new List<string>();
            for (int rowIndex = 0; ; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                    break;
                ICell cell = row.GetCell(columnIndex - 1); // Adjust column index
                if (cell != null)
                    columnData.Add(cell.ToString());
            }
            return columnData.ToArray();
        }

        public string[] ReadColumn(string address)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(address);
            return ReadColumn(cellReference.Col + 1); // Adjust column index
        }

        public string[] ReadRow(int rowIndex)
        {
            ValidatePrerequisites();
            List<string> rowData = new List<string>();
            IRow row = sheet.GetRow(rowIndex - 1); // Adjust row index
            if (row != null)
            {
                int cellIndex = 0;
                while (true)
                {
                    ICell cell = row.GetCell(cellIndex);
                    if (cell == null)
                        break;

                    rowData.Add(cell.ToString());
                    cellIndex++;
                }
            }
            return rowData.ToArray();
        }

        public string[] ReadRow(string address)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(address);
            return ReadRow(cellReference.Row + 1); // Adjust row index
        }

        public DataTable ReadSelection(int startRow, int startColumn, int endRow, int endColumn)
        {
            ValidatePrerequisites();

            if (startRow <= 0 || startColumn <= 0 || endRow <= 0 || endColumn <= 0) // Update the condition for index check
                throw new ArgumentException("Cell index must be > 0");

            DataTable dataTable = new DataTable();
            for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
            {
                string columnName = ToColumnName(columnIndex);
                dataTable.Columns.Add(columnName);
            }

            for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex - 1); // Adjust row index
                if (row != null)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
                    {
                        ICell cell = row.GetCell(columnIndex - 1); // Adjust column index
                        if (cell != null)
                        {
                            int dataTableColumnIndex = columnIndex - startColumn; // Adjust column index
                            if (dataTableColumnIndex >= dataTable.Columns.Count)
                            {
                                DataColumn dataColumn = new DataColumn();
                                dataTable.Columns.Add(dataColumn);
                            }
                            dataRow[dataTableColumnIndex] = cell.ToString();
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            return dataTable;
        }

        private string ToColumnName(int column)
        {
            const int lettersCount = 26;
            string columnName = "";
            while (column > 0)
            {
                column--;
                columnName = Convert.ToChar('A' + column % lettersCount) + columnName;
                column /= lettersCount;
            }
            return columnName;
        }

        public DataTable ReadSelection(string startAddress, string endAddress)
        {
            CellReference startReference = new CellReference(startAddress);
            CellReference endReference = new CellReference(endAddress);

            int startRow = startReference.Row + 1;
            int startColumn = startReference.Col + 1;
            int endRow = endReference.Row + 1;
            int endColumn = endReference.Col + 1;

            return ReadSelection(startRow, startColumn, endRow, endColumn);
        }

        public void LoadSheet(string sheetName)
        {
            using (FileStream fileStream = new FileStream(config.FileLocation, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(fileStream);
                sheet = workbook.GetSheet(sheetName);
            }
        }

        public void WriteCell(int row, int column, string data)
        {
            ValidatePrerequisites();
            IRow excelRow = sheet.GetRow(row);
            if (excelRow is null)
            {
                excelRow = sheet.CreateRow(row);
            }

            ICell cell = excelRow.GetCell(column);
            if (cell is null)
            {
                cell = excelRow.CreateCell(column);
            }
            cell.SetCellValue(data);
        }

        public void WriteCell(string address, string data)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(address);
            WriteCell(cellReference.Row, cellReference.Col, data);
        }

        public void WriteColumn(int column, string[] data)
        {
            ValidatePrerequisites();
            for (int rowIndex = 0; rowIndex < data.Length; rowIndex++)
            {
                string cellValue = data[rowIndex];
                WriteCell(rowIndex, column, cellValue);
            }
        }

        public void WriteColumn(string column, string[] data)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(column);
            WriteColumn(cellReference.Col + 1, data);
        }

        public void WriteRow(int row, string[] data)
        {
            row -= 1;
            ValidatePrerequisites();
            IRow excelRow = sheet.GetRow(row);
            if (excelRow == null)
                excelRow = sheet.CreateRow(row);
            for (int columnIndex = 0; columnIndex < data.Length; columnIndex++)
            {
                string cellValue = data[columnIndex];
                WriteCell(row, columnIndex, cellValue);
            }
        }

        public void WriteRow(string address, string[] data)
        {
            ValidatePrerequisites();
            CellReference cellReference = new CellReference(address);
            WriteRow(cellReference.Row, data);
        }

        public void WriteSelection(int startRow, int startColumn, DataTable data)
        {
            ValidatePrerequisites();
            int numRows = startRow + data.Rows.Count;
            int numCols = startColumn + data.Columns.Count;

            for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
            {
                DataRow dataRow = data.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < data.Columns.Count; columnIndex++)
                {
                    string cellValue = dataRow[columnIndex].ToString();
                    WriteCell(startRow + rowIndex, startColumn + columnIndex, cellValue);
                }
            }
        }

        public void WriteSelection(string startAddress, DataTable data)
        {
            ValidatePrerequisites();
            CellReference startReference = new CellReference(startAddress);
            WriteSelection(startReference.Row, startReference.Col, data);
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

        public string GetActiveSheet()
        {
            var activeSheetIndex = workbook.ActiveSheetIndex;
            var activeSheetTitle = workbook.GetSheetName(activeSheetIndex);
            return activeSheetTitle;
        }

        public void CreateSheets(params string[] sheetNames)
        {
            ValidatePrerequisites();
            foreach (string sheetName in sheetNames)
            {
                if (workbook.GetSheetIndex(sheetName) == -1)
                {
                    workbook.CreateSheet(sheetName);
                }
            }
        }

        public void ApplyFormatting(int startRow, int startColumn, int endRow, int endColumn, Formatting formatting)
        {
            startRow--;
            startColumn--;
            endRow--;
            endColumn--;

            // Apply text formatting
            if (formatting.TextFormating is not null)
            {
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    var row = sheet.GetRow(rowIndex);
                    if (row == null)
                        continue;

                    for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
                    {
                        var cell = row.GetCell(columnIndex);
                        if (cell == null)
                            continue;

                        var cellStyle = cell.CellStyle ?? workbook.CreateCellStyle();
                        XSSFFont font = (XSSFFont)cellStyle.GetFont(workbook) ?? (XSSFFont)workbook.CreateFont();

                        font.IsBold = formatting.TextFormating.Bold;
                        font.IsItalic = formatting.TextFormating.Italic;
                        font.Underline = formatting.TextFormating.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
                        font.FontHeightInPoints = formatting.TextFormating.Size;
                        font.FontName = formatting.TextFormating.Font;
                        font.SetColor(GetXSSFColor(formatting.TextFormating.Color));
                        cellStyle.SetFont(font);

                        // Set horizontal alignment
                        cellStyle.Alignment = ConvertToNPOIHorizontalAlignment(formatting.TextFormating.HorizontalAlignment);

                        // Set vertical alignment
                        cellStyle.VerticalAlignment = ConvertToNPOIVerticalAlignment(formatting.TextFormating.VerticalAlignment);

                        cell.CellStyle = cellStyle;
                    }
                }

                // Apply cell formatting
                if (formatting.CellFormating is not null)
                {
                    for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                    {
                        var row = sheet.GetRow(rowIndex);
                        if (row == null)
                            continue;

                        for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
                        {
                            XSSFCell cell = (XSSFCell)row.GetCell(columnIndex);
                            if (cell == null)
                                continue;

                            XSSFCellStyle cellStyle = (XSSFCellStyle)(cell.CellStyle ?? workbook.CreateCellStyle());
                            cellStyle.FillPattern = FillPattern.SolidForeground;

                            var xssfColor = GetXSSFColor(formatting.CellFormating.BackgroundColor);
                            cellStyle.SetFillForegroundColor(xssfColor);

                            cell.CellStyle = cellStyle;
                        }
                    }
                }
            }
        }


        private XSSFColor GetXSSFColor(System.Drawing.Color color)
        {
            byte[] rgb = new byte[3];
            rgb[0] = color.R;
            rgb[1] = color.G;
            rgb[2] = color.B;
            return new XSSFColor(rgb);
        }

        private HorizontalAlignment ConvertToNPOIHorizontalAlignment(HorizontalAllignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAllignment.LEFT:
                    return HorizontalAlignment.Left;
                case HorizontalAllignment.CENTER:
                    return HorizontalAlignment.Center;
                case HorizontalAllignment.RIGHT:
                    return HorizontalAlignment.Right;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment));
            }
        }

        private VerticalAlignment ConvertToNPOIVerticalAlignment(VerticalAllignment alignment)
        {
            switch (alignment)
            {
                case VerticalAllignment.TOP:
                    return VerticalAlignment.Top;
                case VerticalAllignment.MIDDLE:
                    return VerticalAlignment.Center;
                case VerticalAllignment.BOTTOM:
                    return VerticalAlignment.Bottom;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment));
            }
        }


    }
}
