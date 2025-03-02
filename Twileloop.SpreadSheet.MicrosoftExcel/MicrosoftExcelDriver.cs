using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Twileloop.SpreadSheet.Factory.Base;
using Twileloop.SpreadSheet.Styling;

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

        public string ReadCell(Addr addr)
        {
            IRow excelRow = sheet.GetRow(addr.Row - 1);
            if (excelRow is not null)
            {
                ICell cell = excelRow.GetCell(addr.Column - 1);
                return cell?.ToString();
            }
            return null;
        }

        public string[] ReadColumn(Addr addr)
        {
            var columnData = new List<string>();
            for (int rowIndex = 0; ; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                    break;
                ICell cell = row.GetCell(addr.Column - 1); // Adjust column index
                if (cell != null)
                    columnData.Add(cell.ToString());
            }
            return columnData.ToArray();
        }

        public string[] ReadRow(Addr addr)
        {
            List<string> rowData = new List<string>();
            IRow row = sheet.GetRow(addr.Row - 1); // Adjust row index
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

        public DataTable ReadSelection(Addr start, Addr end)
        {
            if (start.Row <= 0 || start.Column <= 0 || end.Row <= 0 || end.Column <= 0) // Update the condition for index check
                throw new ArgumentException("Cell index must be > 0");

            DataTable dataTable = new DataTable();
            for (int columnIndex = start.Column; columnIndex <= end.Column; columnIndex++)
            {
                string columnName = ToColumnName(columnIndex);
                dataTable.Columns.Add(columnName);
            }

            for (int rowIndex = start.Row; rowIndex <= end.Row; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex - 1); // Adjust row index
                if (row != null)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int columnIndex = start.Column; columnIndex <= end.Column; columnIndex++)
                    {
                        ICell cell = row.GetCell(columnIndex - 1); // Adjust column index
                        if (cell != null)
                        {
                            int dataTableColumnIndex = columnIndex - start.Column; // Adjust column index
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

        public void WriteCell(Addr addr, string data, SpreadsheetStyling style = null)
        {
            IRow excelRow = sheet.GetRow(addr.Row);
            if (excelRow is null)
            {
                excelRow = sheet.CreateRow(addr.Row);
            }

            ICell cell = excelRow.GetCell(addr.Column);
            if (cell is null)
            {
                cell = excelRow.CreateCell(addr.Column);
            }
            cell.SetCellValue(data);

            if (style is not null)
            {
                ApplyStyling(addr, addr, style);
            }
        }

        public void WriteColumn(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            for (int rowIndex = 0; rowIndex < data.Length; rowIndex++)
            {
                string cellValue = data[rowIndex];
                WriteCell((addr.Row + 1 + rowIndex, addr.Column + 1), cellValue); // FIX: Removed `+1`
            }

            if (style is not null)
            {
                ApplyStyling(addr, addr, style);
            }
        }

        public void WriteRow(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            for (int columnIndex = 0; columnIndex < data.Length; columnIndex++)
            {
                string cellValue = data[columnIndex];
                WriteCell((addr.Row + 1, addr.Column + 1 + columnIndex), cellValue); // FIX: Removed `+1`
            }

            if (style is not null)
            {
                ApplyStyling(addr, addr, style);
            }
        }

        public void WriteTable(Addr startAddr, DataTable data, SpreadsheetStyling style = null)
        {
            int rowCount = data.Rows.Count;
            int columnCount = data.Columns.Count;

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                DataRow dataRow = data.Rows[rowIndex];
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    string cellValue = dataRow[columnIndex].ToString();
                    WriteCell((startAddr.Row + 1 + rowIndex, startAddr.Column + 1 + columnIndex), cellValue);
                }
            }

            if (style is not null)
            {
                Addr endAddr = (startAddr.Row + rowCount, startAddr.Column + columnCount);
                ApplyStyling(startAddr, endAddr, style);
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

        public void ApplyBorder(Addr start, Addr end, BorderStyling styling)
        {
            for (int rowIndex = start.Row; rowIndex <= end.Row; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

                for (int columnIndex = start.Column; columnIndex <= end.Column; columnIndex++)
                {
                    XSSFCell cell = (XSSFCell)(row.GetCell(columnIndex) ?? row.CreateCell(columnIndex));

                    // Clone the existing style (important to retain previous styles)
                    XSSFCellStyle cellStyle = (XSSFCellStyle)(cell.CellStyle ?? workbook.CreateCellStyle());
                    XSSFCellStyle newStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                    newStyle.CloneStyleFrom(cellStyle);

                    XSSFColor xssfColor = new XSSFColor(new byte[] {
                styling.BorderColor.R, styling.BorderColor.G, styling.BorderColor.B
            });

                    BorderStyle npoiBorderStyle = ConvertToNPOIBorderStyle(styling.BorderType, styling.Thickness);

                    // Apply borders only to the perimeter of the selected range
                    bool isTop = rowIndex == start.Row;
                    bool isBottom = rowIndex == end.Row;
                    bool isLeft = columnIndex == start.Column;
                    bool isRight = columnIndex == end.Column;

                    if (isTop && styling.TopBorder)
                    {
                        newStyle.BorderTop = npoiBorderStyle;
                        newStyle.SetTopBorderColor(xssfColor);
                    }

                    if (isBottom && styling.BottomBorder)
                    {
                        newStyle.BorderBottom = npoiBorderStyle;
                        newStyle.SetBottomBorderColor(xssfColor);
                    }

                    if (isLeft && styling.LeftBorder)
                    {
                        newStyle.BorderLeft = npoiBorderStyle;
                        newStyle.SetLeftBorderColor(xssfColor);
                    }

                    if (isRight && styling.RightBorder)
                    {
                        newStyle.BorderRight = npoiBorderStyle;
                        newStyle.SetRightBorderColor(xssfColor);
                    }

                    // Assign modified style to cell
                    cell.CellStyle = newStyle;
                }
            }
        }

        public void MergeCells(Addr start, Addr end)
        {
            CellRangeAddress mergeRange = new CellRangeAddress(start.Row, end.Row, start.Column, end.Column);
            sheet.AddMergedRegion(mergeRange);
        }

        private BorderStyle ConvertToNPOIBorderStyle(BorderType borderType, BorderThickness thickness)
        {
            return borderType
            switch
            {
                BorderType.SOLID => thickness
                switch
                {
                    BorderThickness.Thin => BorderStyle.Thin,
                    BorderThickness.Medium => BorderStyle.Medium,
                    BorderThickness.Thick => BorderStyle.Thick,
                    BorderThickness.DoubleLined => BorderStyle.Double, // Closest for thickest possible
                    _ => BorderStyle.Thin
                },
                BorderType.DOTTED => BorderStyle.Dotted,
                BorderType.DASHED => BorderStyle.Dashed,
                _ => BorderStyle.Thin
            };
        }

        public void ApplyStyling(Addr start, Addr end, SpreadsheetStyling styling)
        {
            XSSFCellStyle cellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            XSSFFont font = (XSSFFont)workbook.CreateFont();

            // Apply text formatting if present
            if (styling.TextFormating is not null)
            {
                font.IsBold = styling.TextFormating.Bold;
                font.IsItalic = styling.TextFormating.Italic;
                font.Underline = styling.TextFormating.Underline ? FontUnderlineType.Single : FontUnderlineType.None;
                font.FontHeightInPoints = styling.TextFormating.Size;
                font.FontName = styling.TextFormating.Font;
                font.SetColor(GetXSSFColor(styling.TextFormating.FontColor));

                cellStyle.SetFont(font);
                cellStyle.Alignment = ConvertToNPOIHorizontalAlignment(styling.TextFormating.HorizontalAlignment);
                cellStyle.VerticalAlignment = ConvertToNPOIVerticalAlignment(styling.TextFormating.VerticalAlignment);
            }

            // Apply cell formatting if present
            if (styling.CellFormating is not null)
            {
                cellStyle.FillPattern = FillPattern.SolidForeground;
                cellStyle.SetFillForegroundColor(GetXSSFColor(styling.CellFormating.BackgroundColor));
            }

            // Apply the final style to all selected cells
            for (int rowIndex = start.Row; rowIndex <= end.Row; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);

                for (int columnIndex = start.Column; columnIndex <= end.Column; columnIndex++)
                {
                    var cell = (XSSFCell)(row.GetCell(columnIndex) ?? row.CreateCell(columnIndex));
                    cell.CellStyle = cellStyle;
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

        private HorizontalAlignment ConvertToNPOIHorizontalAlignment(HorizontalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalTxtAlignment.LEFT:
                    return HorizontalAlignment.Left;
                case HorizontalTxtAlignment.CENTER:
                    return HorizontalAlignment.Center;
                case HorizontalTxtAlignment.RIGHT:
                    return HorizontalAlignment.Right;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment));
            }
        }

        private VerticalAlignment ConvertToNPOIVerticalAlignment(VerticalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalTxtAlignment.TOP:
                    return VerticalAlignment.Top;
                case VerticalTxtAlignment.MIDDLE:
                    return VerticalAlignment.Center;
                case VerticalTxtAlignment.BOTTOM:
                    return VerticalAlignment.Bottom;
                default:
                    throw new ArgumentOutOfRangeException(nameof(alignment));
            }
        }

        public void ResizeColumn(Addr addr, int width)
        {
            sheet.SetColumnWidth(addr.Column, width * 256);
        }

        public void ResizeRow(Addr addr, float height)
        {
            IRow row = sheet.GetRow(addr.Row) ?? sheet.CreateRow(addr.Row);
            row.HeightInPoints = height;
        }

        public void AutoFitAllColumns()
        {
            int columnCount = sheet.GetRow(0).LastCellNum;
            for (int col = 0; col < columnCount; col++)
            {
                sheet.AutoSizeColumn(col);
            }
        }

        public void SaveWorkbook()
        {
            AutoFitAllColumns();
            using (FileStream fileStream = new FileStream(config.FileLocation, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream); // Write the workbook to the file
            }
        }

    }
}