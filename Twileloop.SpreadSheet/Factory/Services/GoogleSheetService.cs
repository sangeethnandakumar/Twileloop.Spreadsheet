using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Twileloop.SpreadSheet.Factory.Abstractions;
using Twileloop.SpreadSheet.Factory.Configs;

namespace Twileloop.SpreadSheet.Factory.Services
{
    public class GoogleSheetService : ISpreadSheet
    {
        private readonly GoogleSheetConfiguration config;
        private SheetsService googleSheets;
        public string SheetName { get; set; }
        public string SheetId { get; set; }

        public GoogleSheetService(GoogleSheetConfiguration config)
        {
            this.config = config;
        }

        private void ValidatePrerequisites()
        {
            if (googleSheets is null)
            {
                throw new IOException($"Failed to load the GoogleSheet at '{config.SheetsURI}'");
            }
            if (SheetName is null)
            {
                throw new IOException($"Failed to resolve SheetName");
            }
            if (SheetId is null)
            {
                throw new IOException($"Failed to resolve SheetId");
            }
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

        private static string GetSpreadsheetIdFromUrl(Uri url)
        {
            // Check that the URL is for Google Sheets
            if (url.Host != "docs.google.com" || !url.AbsolutePath.StartsWith("/spreadsheets/d/"))
            {
                throw new ArgumentException("Invalid Google Sheets URL");
            }

            // Extract the spreadsheet ID from the URL
            string spreadsheetId = url.AbsolutePath.Substring("/spreadsheets/d/".Length);
            int end = spreadsheetId.IndexOf("/");
            if (end != -1)
            {
                spreadsheetId = spreadsheetId.Substring(0, end);
            }

            return spreadsheetId;
        }

        public void LoadWorkbook(string sheetName)
        {
            GoogleCredential credential;
            using (var stream = new FileStream(config.Credential, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential
                    .FromStream(stream)
                    .CreateScoped(SheetsService.Scope.Spreadsheets);
            }
            googleSheets = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = config.ApplicationName,
            });
            SheetName = sheetName;
            SheetId = GetSpreadsheetIdFromUrl(config.SheetsURI);
        }

        public string ReadCell(int row, int column)
        {
            string range = $"{SheetName}!{ToColumnName(column)}{row}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(SheetId, range).Execute();
            string cellValue = response.Values?[0]?.FirstOrDefault()?.ToString();
            return cellValue;
        }

        public string ReadCell(string address)
        {
            CellAddress cellAddress = new CellAddress(address);
            int row = cellAddress.Row + 1;
            int column = cellAddress.Column + 1;
            return ReadCell(row, column);
        }

        public string[] ReadColumn(int columnIndex)
        {
            string range = $"{SheetName}!{ToColumnName(columnIndex)}:{ToColumnName(columnIndex)}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(SheetId, range).Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                string[] columnData = values.Select(row => row.FirstOrDefault()?.ToString()).ToArray();
                return columnData;
            }
            return new string[0];
        }

        public string[] ReadColumn(string address)
        {
            CellAddress cellAddress = new CellAddress(address);
            int columnIndex = cellAddress.Column + 1;
            return ReadColumn(columnIndex);
        }

        public string[] ReadRow(int rowIndex)
        {
            string range = $"{SheetName}!A{rowIndex}:{ToColumnName(GetLastColumnIndex())}{rowIndex}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(SheetId, range).Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                string[] rowData = values[0].Select(cell => cell.ToString()).ToArray();
                return rowData;
            }
            return new string[0];
        }

        public string[] ReadRow(string address)
        {
            CellAddress cellAddress = new CellAddress(address);
            int rowIndex = cellAddress.Row + 1;
            return ReadRow(rowIndex);
        }

        private int GetLastColumnIndex()
        {
            string range = $"{SheetName}!1:1";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(SheetId, range).Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                return values[0].Count - 1;
            }
            return -1;
        }

        public DataTable ReadSelection(int startRow, int startColumn, int endRow, int endColumn)
        {
            string range = $"{SheetName}!{ToColumnName(startColumn)}{startRow}:{ToColumnName(endColumn)}{endRow}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(SheetId, range).Execute();
            IList<IList<object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                DataTable dataTable = new DataTable();

                // Create columns
                for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
                {
                    string columnName = ToColumnName(columnIndex);
                    dataTable.Columns.Add(columnName);
                }

                // Add rows
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    IList<object> rowValues = values[rowIndex - startRow];
                    for (int columnIndex = startColumn; columnIndex <= endColumn; columnIndex++)
                    {
                        int valueIndex = columnIndex - startColumn;
                        dataRow[columnIndex - startColumn] = valueIndex < rowValues.Count ? rowValues[valueIndex]?.ToString() : string.Empty;
                    }
                    dataTable.Rows.Add(dataRow);
                }

                return dataTable;
            }

            return new DataTable();
        }

        public DataTable ReadSelection(string startAddress, string endAddress)
        {
            CellAddress startCellAddress = new CellAddress(startAddress);
            CellAddress endCellAddress = new CellAddress(endAddress);

            int startRow = startCellAddress.Row + 1;
            int startColumn = startCellAddress.Column + 1;
            int endRow = endCellAddress.Row + 1;
            int endColumn = endCellAddress.Column + 1;

            return ReadSelection(startRow, startColumn, endRow, endColumn);
        }

        public void WriteCell(int row, int column, string data)
        {
            ValidatePrerequisites();
            string range = $"{SheetName}!{ToColumnName(column)}{row}";
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object> { data } }
            };
            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest =
                googleSheets.Spreadsheets.Values.Update(valueRange, SheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        public void WriteCell(string address, string data)
        {
            ValidatePrerequisites();
            CellAddress cellAddress = new CellAddress(address);
            int row = cellAddress.Row + 1;
            int column = cellAddress.Column + 1;
            WriteCell(row, column, data);
        }

        public void WriteColumn(int column, string[] data)
        {
            ValidatePrerequisites();
            int columnIndex = column;
            string range = $"{SheetName}!{ToColumnName(columnIndex)}:{ToColumnName(columnIndex)}";
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>>()
            };

            foreach (string cellValue in data)
            {
                valueRange.Values.Add(new List<object> { cellValue });
            }

            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest =
                googleSheets.Spreadsheets.Values.Update(valueRange, SheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        private int ParseColumnIndex(string columnName)
        {
            const int lettersCount = 26;
            int columnIndex = 0;
            int power = 1;

            for (int i = columnName.Length - 1; i >= 0; i--)
            {
                char letter = columnName[i];
                int value = letter - 'A' + 1; // Convert letter to corresponding value (A=1, B=2, etc.)
                columnIndex += value * power;
                power *= lettersCount;
            }

            return columnIndex;
        }






        public void WriteColumn(string column, string[] data)
        {
            ValidatePrerequisites();
            CellAddress cellAddress = new CellAddress(column);
            WriteColumn(cellAddress.Column + 1, data);
        }

        public void WriteRow(int row, string[] data)
        {
            ValidatePrerequisites();
            string range = $"{SheetName}!A{row}:{ToColumnName(data.Length)}{row}";
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>(data) }
            };
            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest =
                googleSheets.Spreadsheets.Values.Update(valueRange, SheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        public void WriteRow(string address, string[] data)
        {
            ValidatePrerequisites();
            CellAddress cellAddress = new CellAddress(address);
            int row = cellAddress.Row + 1;
            WriteRow(row, data);
        }

        public void WriteSelection(int startRow, int startColumn, DataTable data)
        {
            ValidatePrerequisites();
            int numRows = startRow + data.Rows.Count; // Calculate the end row based on the number of rows in the DataTable
            int numCols = startColumn + data.Columns.Count; // Calculate the end column based on the number of columns in the DataTable

            string range = $"{SheetName}!{ToColumnName(startColumn)}{startRow}:{ToColumnName(numCols - 1)}{numRows - 1}";
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>>(data.Rows.Count)
            };

            foreach (DataRow row in data.Rows)
            {
                List<object> rowData = new List<object>(data.Columns.Count);
                for (int columnIndex = startColumn; columnIndex < numCols; columnIndex++)
                {
                    int dataTableColumnIndex = columnIndex - startColumn;
                    if (dataTableColumnIndex < row.ItemArray.Length)
                        rowData.Add(row[dataTableColumnIndex].ToString());
                    else
                        rowData.Add(string.Empty);
                }
                valueRange.Values.Add(rowData);
            }

            SpreadsheetsResource.ValuesResource.UpdateRequest updateRequest =
                googleSheets.Spreadsheets.Values.Update(valueRange, SheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();
        }

        public void WriteSelection(string startAddress, DataTable data)
        {
            ValidatePrerequisites();
            CellReference startReference = new CellReference(startAddress);
            WriteSelection(startReference.Row + 1, startReference.Col + 1, data);
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }
    }
}
