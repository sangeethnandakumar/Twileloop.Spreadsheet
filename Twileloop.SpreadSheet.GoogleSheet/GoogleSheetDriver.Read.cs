using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Data;

namespace Twileloop.SpreadSheet.GoogleSheet
{
    public partial class GoogleSheetDriver
    {
        public string ReadCell(Addr addr)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column)}:{addr.Row}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(spreadSheetId, range).Execute();
            if (response.Values != null && response.Values.Count > 0 && response.Values[0].Count > 0)
                return response.Values[0][0]?.ToString();
            return null;
        }

        public string[] ReadColumn(Addr addr)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column)}:{ToColumnName(addr.Column)}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(spreadSheetId, range).Execute();

            if (response.Values != null && response.Values.Count > 0)
            {
                var columnData = new List<string>();
                foreach (var row in response.Values)
                    columnData.Add(row.Count > 0 ? row[0]?.ToString() : null);
                return columnData.ToArray();
            }
            return new string[0];
        }

        public string[] ReadRow(Addr addr)
        {
            string range = $"{currentSheetName}!{addr.Row}:{addr.Row}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(spreadSheetId, range).Execute();

            if (response.Values != null && response.Values.Count > 0)
            {
                var rowData = new List<string>();
                foreach (var cell in response.Values[0])
                    rowData.Add(cell?.ToString());
                return rowData.ToArray();
            }
            return new string[0];
        }

        public DataTable ReadSelection(Addr start, Addr end)
        {
            if (start.Row <= 0 || start.Column <= 0 || end.Row <= 0 || end.Column <= 0)
                throw new ArgumentException("Cell index must be > 0");

            string range = $"{currentSheetName}!{ToColumnName(start.Column)}{start.Row}:{ToColumnName(end.Column)}{end.Row}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(spreadSheetId, range).Execute();
            DataTable dataTable = new DataTable();

            // Create columns
            for (int columnIndex = start.Column; columnIndex <= end.Column; columnIndex++)
                dataTable.Columns.Add(ToColumnName(columnIndex));

            // Add data if there are values
            if (response.Values != null && response.Values.Count > 0)
            {
                for (int rowIndex = 0; rowIndex < response.Values.Count; rowIndex++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    var rowValues = response.Values[rowIndex];

                    for (int columnIndex = 0; columnIndex < end.Column - start.Column + 1; columnIndex++)
                        dataRow[columnIndex] = columnIndex < rowValues.Count ? rowValues[columnIndex]?.ToString() : string.Empty;

                    dataTable.Rows.Add(dataRow);
                }
            }
            return dataTable;
        }
    }
}