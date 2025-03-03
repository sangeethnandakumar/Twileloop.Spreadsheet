using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using NPOI.POIFS.Crypt;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using Twileloop.SpreadSheet.Factory.Base;
using Twileloop.SpreadSheet.Styling;

namespace Twileloop.SpreadSheet.GoogleSheet
{
    public class GoogleSheetDriver : ISpreadSheetDriver
    {
        private readonly GoogleSheetOptions config;
        private SheetsService googleSheets;
        private string currentSheetName;
        private string spreadSheetId;
        private int? sheetId;

        // Store pending requests for batch updates
        private List<Request> pendingRequests = new List<Request>();
        private List<KeyValuePair<string, ValueRange>> pendingValueUpdates = new List<KeyValuePair<string, ValueRange>>();

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
        }

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

        private (Addr start, Addr end) GetMergedCellRange(Addr addr)
        {
            var request = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = request.Sheets.FirstOrDefault(s => s.Properties.SheetId == sheetId);

            if (sheet != null && sheet.Merges != null)
            {
                var mergedRange = sheet.Merges.FirstOrDefault(range =>
                    range.StartRowIndex <= addr.Row &&
                    range.EndRowIndex > addr.Row &&
                    range.StartColumnIndex <= addr.Column &&
                    range.EndColumnIndex > addr.Column);

                if (mergedRange != null)
                {
                    return ((mergedRange.StartRowIndex.Value + 1, mergedRange.StartColumnIndex.Value + 1),
                            (mergedRange.EndRowIndex.Value, mergedRange.EndColumnIndex.Value));
                }
            }
            return (addr, addr);
        }

        public void WriteCell(Addr addr, string data, SpreadsheetStyling style = null)
        {
            var (start, end) = GetMergedCellRange(addr);
            string range = $"{currentSheetName}!{ToColumnName(start.Column + 1)}{start.Row + 1}:{ToColumnName(end.Column + 1)}{end.Row + 1}";

            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object> { data } }
            };

            if (config.BulkUpdate)
            {
                var index = pendingValueUpdates.FindIndex(x => x.Key == range);
                if (index != -1)
                {
                    pendingValueUpdates[index] = new KeyValuePair<string, ValueRange>(range, valueRange);
                }
                else
                {
                    pendingValueUpdates.Add(new KeyValuePair<string, ValueRange>(range, valueRange));
                }

                if (style != null)
                    QueueStylingRequest(start, end, style);
            }
            else
            {
                var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateRequest.Execute();

                if (style != null)
                    ApplyStylingImmediate(start, end, style);
            }
        }

        public void WriteColumn(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column + 1)}{addr.Row + 1}:{ToColumnName(addr.Column + 1)}{addr.Row + data.Length}";

            var valueRange = new ValueRange { Values = new List<IList<object>>() };
            foreach (var value in data)
                valueRange.Values.Add(new List<object> { value });

            if (config.BulkUpdate)
            {
                var index = pendingValueUpdates.FindIndex(x => x.Key == range);
                if (index != -1)
                {
                    pendingValueUpdates[index] = new KeyValuePair<string, ValueRange>(range, valueRange);
                }
                else
                {
                    pendingValueUpdates.Add(new KeyValuePair<string, ValueRange>(range, valueRange));
                }
                if (style != null)
                {
                    Addr endAddr = (addr.Row + data.Length - 1, addr.Column);
                    QueueStylingRequest(addr, endAddr, style);
                }
            }
            else
            {
                var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateRequest.Execute();

                if (style != null)
                {
                    Addr endAddr = (addr.Row + data.Length - 1, addr.Column);
                    ApplyStylingImmediate(addr, endAddr, style);
                }
            }
        }

        public void WriteRow(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column + 1)}{addr.Row + 1}:{ToColumnName(addr.Column + data.Length)}{addr.Row + 1}";

            var valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>(data) }
            };

            if (config.BulkUpdate)
            {
                var index = pendingValueUpdates.FindIndex(x => x.Key == range);
                if (index != -1)
                {
                    pendingValueUpdates[index] = new KeyValuePair<string, ValueRange>(range, valueRange);
                }
                else
                {
                    pendingValueUpdates.Add(new KeyValuePair<string, ValueRange>(range, valueRange));
                }

                if (style != null)
                {
                    Addr endAddr = (addr.Row, addr.Column + data.Length - 1);
                    QueueStylingRequest(addr, endAddr, style);
                }
            }
            else
            {
                var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateRequest.Execute();

                if (style != null)
                {
                    Addr endAddr = (addr.Row, addr.Column + data.Length - 1);
                    ApplyStylingImmediate(addr, endAddr, style);
                }
            }
        }

        public void WriteTable(Addr startAddr, DataTable data, SpreadsheetStyling style = null)
        {
            // Execute immediately regardless of BulkUpdate setting
            int rowCount = data.Rows.Count;
            int columnCount = data.Columns.Count;

            string range = $"{currentSheetName}!{ToColumnName(startAddr.Column + 1)}{startAddr.Row + 1}:{ToColumnName(startAddr.Column + columnCount)}{startAddr.Row + rowCount}";

            var valueRange = new ValueRange { Values = new List<IList<object>>() };
            foreach (DataRow row in data.Rows)
            {
                var rowValues = new List<object>();
                foreach (var item in row.ItemArray)
                    rowValues.Add(item.ToString());
                valueRange.Values.Add(rowValues);
            }

            var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            if (style != null)
            {
                Addr endAddr = (startAddr.Row + rowCount, startAddr.Column + columnCount);
                ApplyStylingImmediate(startAddr, endAddr, style);
            }
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
            sheetId = GetActiveSheetId();
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

        private int? GetActiveSheetId()
        {
            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == currentSheetName);
            return sheet?.Properties.SheetId;
        }

        private void QueueStylingRequest(Addr start, Addr end, SpreadsheetStyling styling)
        {
            if (!sheetId.HasValue) return;

            var gridRange = new GridRange
            {
                SheetId = sheetId.Value,
                StartRowIndex = start.Row,
                EndRowIndex = end.Row + 1,
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 1
            };

            // Apply text formatting
            if (styling.TextFormating != null)
            {
                var textFormat = new TextFormat
                {
                    Bold = styling.TextFormating.Bold,
                    Italic = styling.TextFormating.Italic,
                    Underline = styling.TextFormating.Underline,
                    FontSize = styling.TextFormating.Size,
                    FontFamily = styling.TextFormating.Font
                };

                if (styling.TextFormating.FontColor != null)
                {
                    textFormat.ForegroundColor = new Color
                    {
                        Red = styling.TextFormating.FontColor.R / 255f,
                        Green = styling.TextFormating.FontColor.G / 255f,
                        Blue = styling.TextFormating.FontColor.B / 255f
                    };
                }

                var cellFormat = new CellFormat
                {
                    TextFormat = textFormat,
                    HorizontalAlignment = ConvertToGoogleHorizontalAlignment(styling.TextFormating.HorizontalAlignment),
                    VerticalAlignment = ConvertToGoogleVerticalAlignment(styling.TextFormating.VerticalAlignment)
                };

                pendingRequests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = gridRange,
                        Cell = new CellData { UserEnteredFormat = cellFormat },
                        Fields = "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
                    }
                });
            }

            // Apply cell formatting
            if (styling.CellFormating != null)
            {
                var cellFormat = new CellFormat
                {
                    BackgroundColor = new Color
                    {
                        Red = styling.CellFormating.BackgroundColor.R / 255f,
                        Green = styling.CellFormating.BackgroundColor.G / 255f,
                        Blue = styling.CellFormating.BackgroundColor.B / 255f
                    }
                };

                pendingRequests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = gridRange,
                        Cell = new CellData { UserEnteredFormat = cellFormat },
                        Fields = "userEnteredFormat.backgroundColor"
                    }
                });
            }
        }

        private void ApplyStylingImmediate(Addr start, Addr end, SpreadsheetStyling styling)
        {
            if (!sheetId.HasValue) return;

            var requests = new List<Request>();
            var gridRange = new GridRange
            {
                SheetId = sheetId.Value,
                StartRowIndex = start.Row,
                EndRowIndex = end.Row + 1,
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 1
            };

            // Apply text formatting
            if (styling.TextFormating != null)
            {
                var textFormat = new TextFormat
                {
                    Bold = styling.TextFormating.Bold,
                    Italic = styling.TextFormating.Italic,
                    Underline = styling.TextFormating.Underline,
                    FontSize = styling.TextFormating.Size,
                    FontFamily = styling.TextFormating.Font
                };

                if (styling.TextFormating.FontColor != null)
                {
                    textFormat.ForegroundColor = new Color
                    {
                        Red = styling.TextFormating.FontColor.R / 255f,
                        Green = styling.TextFormating.FontColor.G / 255f,
                        Blue = styling.TextFormating.FontColor.B / 255f
                    };
                }

                var cellFormat = new CellFormat
                {
                    TextFormat = textFormat,
                    HorizontalAlignment = ConvertToGoogleHorizontalAlignment(styling.TextFormating.HorizontalAlignment),
                    VerticalAlignment = ConvertToGoogleVerticalAlignment(styling.TextFormating.VerticalAlignment)
                };

                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = gridRange,
                        Cell = new CellData { UserEnteredFormat = cellFormat },
                        Fields = "userEnteredFormat.textFormat,userEnteredFormat.horizontalAlignment,userEnteredFormat.verticalAlignment"
                    }
                });
            }

            // Apply cell formatting
            if (styling.CellFormating != null)
            {
                var cellFormat = new CellFormat
                {
                    BackgroundColor = new Color
                    {
                        Red = styling.CellFormating.BackgroundColor.R / 255f,
                        Green = styling.CellFormating.BackgroundColor.G / 255f,
                        Blue = styling.CellFormating.BackgroundColor.B / 255f
                    }
                };

                requests.Add(new Request
                {
                    RepeatCell = new RepeatCellRequest
                    {
                        Range = gridRange,
                        Cell = new CellData { UserEnteredFormat = cellFormat },
                        Fields = "userEnteredFormat.backgroundColor"
                    }
                });
            }

            if (requests.Count > 0)
            {
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        public void ApplyStyling(Addr start, Addr end, SpreadsheetStyling styling)
        {
            // Execute immediately regardless of BulkUpdate setting
            ApplyStylingImmediate(start, end, styling);
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

        public void ApplyBorder(Addr start, Addr end, BorderStyling styling)
        {
            // Execute immediately regardless of BulkUpdate setting
            if (!sheetId.HasValue) return;

            // Convert border color to Google's color format
            var borderColor = new Color
            {
                Red = styling.BorderColor.R / 255f,
                Green = styling.BorderColor.G / 255f,
                Blue = styling.BorderColor.B / 255f
            };

            // Convert border style to Google's format
            string style = ConvertToGoogleBorderStyle(styling.BorderType, styling.Thickness);

            var updateBordersRequest = new UpdateBordersRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId.Value,
                    StartRowIndex = Math.Max(0, start.Row),
                    EndRowIndex = end.Row + 1,
                    StartColumnIndex = Math.Max(0, start.Column),
                    EndColumnIndex = end.Column + 1
                }
            };

            if (styling.TopBorder)
                updateBordersRequest.Top = new Border { Style = style, Color = borderColor };

            if (styling.BottomBorder)
                updateBordersRequest.Bottom = new Border { Style = style, Color = borderColor };

            if (styling.LeftBorder)
                updateBordersRequest.Left = new Border { Style = style, Color = borderColor };

            if (styling.RightBorder)
                updateBordersRequest.Right = new Border { Style = style, Color = borderColor };

            var requests = new List<Request> { new Request { UpdateBorders = updateBordersRequest } };
            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
            googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
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

        public void MergeCells(Addr start, Addr end)
        {
            if (!sheetId.HasValue) return;

            var mergeCellsRequest = new MergeCellsRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId.Value,
                    StartRowIndex = start.Row,
                    EndRowIndex = end.Row + 1,
                    StartColumnIndex = start.Column,
                    EndColumnIndex = end.Column + 1
                },
                MergeType = "MERGE_ALL"
            };

            if (config.BulkUpdate)
            {
                pendingRequests.Add(new Request { MergeCells = mergeCellsRequest });
            }
            else
            {
                var requests = new List<Request> { new Request { MergeCells = mergeCellsRequest } };
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        public void ResizeColumn(Addr addr, int width)
        {
            if (!sheetId.HasValue) return;

            var updateDimensionPropertiesRequest = new UpdateDimensionPropertiesRequest
            {
                Range = new DimensionRange
                {
                    SheetId = sheetId.Value,
                    Dimension = "COLUMNS",
                    StartIndex = addr.Column,
                    EndIndex = addr.Column + 1
                },
                Properties = new DimensionProperties { PixelSize = width * 4 },
                Fields = "pixelSize"
            };

            if (config.BulkUpdate)
            {
                pendingRequests.Add(new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest });
            }
            else
            {
                var requests = new List<Request> { new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest } };
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        public void ResizeRow(Addr addr, float height)
        {
            if (!sheetId.HasValue) return;

            var updateDimensionPropertiesRequest = new UpdateDimensionPropertiesRequest
            {
                Range = new DimensionRange
                {
                    SheetId = sheetId.Value,
                    Dimension = "ROWS",
                    StartIndex = addr.Row,
                    EndIndex = addr.Row + 1
                },
                Properties = new DimensionProperties { PixelSize = (int)(height * 2) },
                Fields = "pixelSize"
            };

            if (config.BulkUpdate)
            {
                pendingRequests.Add(new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest });
            }
            else
            {
                var requests = new List<Request> { new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest } };
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        public void AutoFitAllColumns()
        {
            // This is a no-op in the Google Sheets implementation
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
    }
}