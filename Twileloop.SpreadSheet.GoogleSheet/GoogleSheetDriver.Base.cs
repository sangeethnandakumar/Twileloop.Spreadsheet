using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Twileloop.SpreadSheet.Factory.Base;
using Twileloop.SpreadSheet.Styling;

namespace Twileloop.SpreadSheet.GoogleSheet
{
    public partial class GoogleSheetDriver : ISpreadSheetDriver
    {
        private readonly GoogleSheetOptions config;
        private SheetsService googleSheets;
        private string currentSheetName;
        private string spreadSheetId;
        private int? sheetId;

        // Store pending requests for batch updates
        private List<Request> pendingRequests = new List<Request>();
        private List<KeyValuePair<string, ValueRange>> pendingValueUpdates = new List<KeyValuePair<string, ValueRange>>();

        // Cache for merged cells
        private List<GridRange> cachedMerges = new List<GridRange>();

        public string DriverName => "GoogleSheet";

        public GoogleSheetDriver(GoogleSheetOptions config)
        {
            this.config = config;
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
            if (url.Host != "docs.google.com" || !url.AbsolutePath.StartsWith("/spreadsheets/d/"))
                throw new ArgumentException("Invalid Google Sheets URL");

            string spreadsheetId = url.AbsolutePath.Substring("/spreadsheets/d/".Length);
            int end = spreadsheetId.IndexOf("/");
            return end != -1 ? spreadsheetId.Substring(0, end) : spreadsheetId;
        }

        public void InitialiseWorkbook()
        {
            GoogleCredential credential;
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(config.JsonCredentialContent)))
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

            spreadSheetId = GetSpreadsheetIdFromUrl(config.SheetsURI);
            pendingRequests.Clear();
            pendingValueUpdates.Clear();
            cachedMerges.Clear();
        }

        private int? GetActiveSheetId()
        {
            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == currentSheetName);
            return sheet?.Properties.SheetId;
        }

        public void Dispose()
        {
            GC.SuppressFinalize(this);
        }

        public string[] GetSheets()
        {
            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            return spreadsheet.Sheets.Select(sheet => sheet.Properties.Title).ToArray();
        }

        public void OpenSheet(string sheetName)
        {
            if (googleSheets == null)
                throw new IOException("Workbook has not been initialized");

            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetName);

            if (sheet == null)
                throw new IOException($"Sheet '{sheetName}' does not exist");

            currentSheetName = sheetName;
            sheetId = sheet.Properties.SheetId;

            // Prefetch merged cells
            cachedMerges.Clear();
            if (sheet.Merges != null)
            {
                foreach (var merge in sheet.Merges)
                {
                    cachedMerges.Add(merge);
                }
            }
        }

        public string GetActiveSheet()
        {
            return currentSheetName;
        }

        public void CreateSheets(params string[] sheetNames)
        {
            // Execute immediately regardless of BulkUpdate setting
            if (googleSheets == null)
                throw new IOException("Workbook has not been initialized");

            var requests = new List<Request>();
            foreach (string sheetName in sheetNames)
            {
                var sheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute().Sheets
                    .FirstOrDefault(s => s.Properties.Title == sheetName);

                if (sheet == null)
                {
                    var addSheetRequest = new AddSheetRequest
                    {
                        Properties = new SheetProperties { Title = sheetName }
                    };
                    requests.Add(new Request { AddSheet = addSheetRequest });
                }
            }

            if (requests.Count > 0)
            {
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        private (Addr start, Addr end) GetMergedCellRange(Addr addr)
        {
            // Use cached merged cells instead of making API call
            var mergedRange = cachedMerges.FirstOrDefault(range =>
                range.StartRowIndex <= addr.Row &&
                range.EndRowIndex > addr.Row &&
                range.StartColumnIndex <= addr.Column &&
                range.EndColumnIndex > addr.Column);

            if (mergedRange != null)
            {
                return (
                    (mergedRange.StartRowIndex.Value + 1, mergedRange.StartColumnIndex.Value + 1),
                    (mergedRange.EndRowIndex.Value, mergedRange.EndColumnIndex.Value)
                );
            }

            return (addr, addr);
        }

        public void MergeCells(Addr start, Addr end)
        {
            if (!sheetId.HasValue) return;

            var gridRange = new GridRange
            {
                SheetId = sheetId.Value,
                StartRowIndex = start.Row,
                EndRowIndex = end.Row + 1,  // Make exclusive
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 1 
            };

            var mergeCellsRequest = new MergeCellsRequest
            {
                Range = gridRange,
                MergeType = "MERGE_ALL"
            };

            if (config.BulkUpdate)
            {
                pendingRequests.Add(new Request { MergeCells = mergeCellsRequest });

                // Update cache for future reference
                cachedMerges.Add(gridRange);
            }
            else
            {
                var requests = new List<Request> { new Request { MergeCells = mergeCellsRequest } };
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();

                // Update cache for future reference
                cachedMerges.Add(gridRange);
            }
        }

        public void SaveWorkbook()
        {
            // Execute all pending batch update requests
            if (pendingRequests.Count > 0)
            {
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = pendingRequests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
                pendingRequests.Clear();
            }

            // Execute all pending value updates
            if (pendingValueUpdates.Count > 0)
            {
                var batchUpdateValuesRequest = new BatchUpdateValuesRequest
                {
                    Data = pendingValueUpdates.Select(kvp => new ValueRange
                    {
                        Range = kvp.Key,
                        Values = kvp.Value.Values
                    }).ToList(),
                    ValueInputOption = "RAW"
                };

                googleSheets.Spreadsheets.Values.BatchUpdate(batchUpdateValuesRequest, spreadSheetId).Execute();
                pendingValueUpdates.Clear();
            }
        }

        public void AutoFitAllColumns()
        {
            // This is a no-op in the Google Sheets implementation
        }

        private string ConvertToGoogleHorizontalAlignment(HorizontalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalTxtAlignment.LEFT: return "LEFT";
                case HorizontalTxtAlignment.CENTER: return "CENTER";
                case HorizontalTxtAlignment.RIGHT: return "RIGHT";
                default: return "LEFT";
            }
        }

        private string ConvertToGoogleVerticalAlignment(VerticalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalTxtAlignment.TOP: return "TOP";
                case VerticalTxtAlignment.MIDDLE: return "MIDDLE";
                case VerticalTxtAlignment.BOTTOM: return "BOTTOM";
                default: return "MIDDLE";
            }
        }

        private string ConvertToGoogleBorderStyle(BorderType borderType, BorderThickness thickness)
        {
            switch (borderType)
            {
                case BorderType.SOLID:
                    switch (thickness)
                    {
                        case BorderThickness.Thin: return "SOLID";
                        case BorderThickness.Medium: return "SOLID_MEDIUM";
                        case BorderThickness.Thick: return "SOLID_THICK";
                        case BorderThickness.DoubleLined: return "DOUBLE";
                        default: return "SOLID";
                    }
                case BorderType.DOTTED: return "DOTTED";
                case BorderType.DASHED: return "DASHED";
                default: return "SOLID";
            }
        }
    }
}