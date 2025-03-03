using Google.Apis.Sheets.v4.Data;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;

namespace Twileloop.SpreadSheet.MicrosoftExcel
{
    public partial class MicrosoftExcelDriver
    {
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
    }
}