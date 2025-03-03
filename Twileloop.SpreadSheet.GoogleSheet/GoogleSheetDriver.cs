using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
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
        }


        public string ReadCell(Addr addr)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column)}:{addr.Row}";
            ValueRange response = googleSheets.Spreadsheets.Values.Get(spreadSheetId, range).Execute();
            if (response.Values != null && response.Values.Count > 0 && response.Values[0].Count > 0)
            {
                return response.Values[0][0]?.ToString();
            }
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
                {
                    if (row.Count > 0)
                    {
                        columnData.Add(row[0]?.ToString());
                    }
                    else
                    {
                        columnData.Add(null);
                    }
                }
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
                {
                    rowData.Add(cell?.ToString());
                }
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
            {
                string columnName = ToColumnName(columnIndex);
                dataTable.Columns.Add(columnName);
            }

            // Add data if there are values
            if (response.Values != null && response.Values.Count > 0)
            {
                for (int rowIndex = 0; rowIndex < response.Values.Count; rowIndex++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    var rowValues = response.Values[rowIndex];

                    for (int columnIndex = 0; columnIndex < end.Column - start.Column + 1; columnIndex++)
                    {
                        if (columnIndex < rowValues.Count)
                        {
                            dataRow[columnIndex] = rowValues[columnIndex]?.ToString();
                        }
                        else
                        {
                            dataRow[columnIndex] = string.Empty;
                        }
                    }

                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        public void WriteCell(Addr addr, string data, SpreadsheetStyling style = null)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column + 1)}{addr.Row + 1}";

            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object> { data } }
            };

            var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            if (style != null)
            {
                ApplyStyling(addr, addr, style);
            }
        }

        public void WriteColumn(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column + 1)}{addr.Row + 1}:{ToColumnName(addr.Column + 1)}{addr.Row + data.Length}";

            var valueRange = new ValueRange
            {
                Values = new List<IList<object>>()
            };

            foreach (var value in data)
            {
                valueRange.Values.Add(new List<object> { value });
            }

            var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            if (style != null)
            {
                Addr endAddr = (addr.Row + data.Length - 1, addr.Column);
                ApplyStyling(addr, endAddr, style);
            }
        }

        public void WriteRow(Addr addr, string[] data, SpreadsheetStyling style = null)
        {
            string range = $"{currentSheetName}!{ToColumnName(addr.Column + 1)}{addr.Row + 1}:{ToColumnName(addr.Column + data.Length)}{addr.Row + 1}";

            var valueRange = new ValueRange
            {
                Values = new List<IList<object>> { new List<object>(data) }
            };

            var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            if (style != null)
            {
                Addr endAddr = (addr.Row, addr.Column + data.Length - 1);
                ApplyStyling(addr, endAddr, style);
            }
        }

        public void WriteTable(Addr startAddr, DataTable data, SpreadsheetStyling style = null)
        {
            
            int rowCount = data.Rows.Count;
            int columnCount = data.Columns.Count;

            string range = $"{currentSheetName}!{ToColumnName(startAddr.Column + 1)}{startAddr.Row + 1}:{ToColumnName(startAddr.Column + columnCount)}{startAddr.Row + rowCount}";


            var valueRange = new ValueRange
            {
                Values = new List<IList<object>>()
            };

            foreach (DataRow row in data.Rows)
            {
                var rowValues = new List<object>();
                foreach (var item in row.ItemArray)
                {
                    rowValues.Add(item.ToString());
                }
                valueRange.Values.Add(rowValues);
            }

            var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            updateRequest.Execute();

            if (style != null)
            {
                Addr endAddr = (startAddr.Row + rowCount - 1, startAddr.Column + columnCount - 1);
                ApplyStyling(startAddr, endAddr, style);
            }
        }

        public void Dispose()
        {
            // No need to explicitly dispose the SheetsService,
            // but we should suppress finalization
            GC.SuppressFinalize(this);
        }

        public string[] GetSheets()
        {
            
            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheetTitles = spreadsheet.Sheets.Select(sheet => sheet.Properties.Title).ToArray();
            return sheetTitles;
        }

        public void OpenSheet(string sheetName)
        {
            if (googleSheets == null)
            {
                throw new IOException("Workbook has not been initialized");
            }

            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == sheetName);

            if (sheet == null)
            {
                throw new IOException($"Sheet '{sheetName}' does not exist");
            }

            currentSheetName = sheetName;
        }

        public string GetActiveSheet()
        {
            
            return currentSheetName;
        }

        public void CreateSheets(params string[] sheetNames)
        {
            if (googleSheets == null)
            {
                throw new IOException("Workbook has not been initialized");
            }

            var requests = new List<Request>();

            foreach (string sheetName in sheetNames)
            {
                var sheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute().Sheets
                    .FirstOrDefault(s => s.Properties.Title == sheetName);

                if (sheet == null)
                {
                    var addSheetRequest = new AddSheetRequest
                    {
                        Properties = new SheetProperties
                        {
                            Title = sheetName
                        }
                    };

                    requests.Add(new Request { AddSheet = addSheetRequest });
                }
            }

            if (requests.Count > 0)
            {
                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = requests
                };

                googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
            }
        }

        private int? GetActiveSheetId()
        {
            var spreadsheet = googleSheets.Spreadsheets.Get(spreadSheetId).Execute();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == currentSheetName);
            return sheet?.Properties.SheetId;
        }

        public void ApplyStyling(Addr start, Addr end, SpreadsheetStyling styling)
        {
            int? sheetId = GetActiveSheetId();
            if (!sheetId.HasValue) return;

            var requests = new List<Request>();

            var gridRange = new GridRange
            {
                SheetId = sheetId.Value,
                StartRowIndex = start.Row,
                EndRowIndex = end.Row + 2,  // Exclusive, so add +1
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 2 // Exclusive, so add +1
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




        private string ConvertToGoogleHorizontalAlignment(HorizontalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalTxtAlignment.LEFT:
                    return "LEFT";
                case HorizontalTxtAlignment.CENTER:
                    return "CENTER";
                case HorizontalTxtAlignment.RIGHT:
                    return "RIGHT";
                default:
                    return "LEFT";
            }
        }

        private string ConvertToGoogleVerticalAlignment(VerticalTxtAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalTxtAlignment.TOP:
                    return "TOP";
                case VerticalTxtAlignment.MIDDLE:
                    return "MIDDLE";
                case VerticalTxtAlignment.BOTTOM:
                    return "BOTTOM";
                default:
                    return "MIDDLE";
            }
        }

        public void ApplyBorder(Addr start, Addr end, BorderStyling styling)
        {
            int? sheetId = GetActiveSheetId();
            if (!sheetId.HasValue) return;

            var requests = new List<Request>();

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
                    StartRowIndex = Math.Max(0, start.Row),   // Ensure non-negative
                    EndRowIndex = end.Row + 1,               // Include last row
                    StartColumnIndex = Math.Max(0, start.Column), // Ensure non-negative
                    EndColumnIndex = end.Column + 1          // Include last column
                }
            };

            if (styling.TopBorder)
            {
                updateBordersRequest.Top = new Border
                {
                    Style = style,
                    Color = borderColor
                };
            }

            if (styling.BottomBorder)
            {
                updateBordersRequest.Bottom = new Border
                {
                    Style = style,
                    Color = borderColor
                };
            }

            if (styling.LeftBorder)
            {
                updateBordersRequest.Left = new Border
                {
                    Style = style,
                    Color = borderColor
                };
            }

            if (styling.RightBorder)
            {
                updateBordersRequest.Right = new Border
                {
                    Style = style,
                    Color = borderColor
                };
            }

            requests.Add(new Request { UpdateBorders = updateBordersRequest });

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
        }


        private string ConvertToGoogleBorderStyle(BorderType borderType, BorderThickness thickness)
        {
            switch (borderType)
            {
                case BorderType.SOLID:
                    switch (thickness)
                    {
                        case BorderThickness.Thin:
                            return "SOLID";
                        case BorderThickness.Medium:
                            return "SOLID_MEDIUM";
                        case BorderThickness.Thick:
                            return "SOLID_THICK";
                        case BorderThickness.DoubleLined:
                            return "DOUBLE";
                        default:
                            return "SOLID";
                    }
                case BorderType.DOTTED:
                    return "DOTTED";
                case BorderType.DASHED:
                    return "DASHED";
                default:
                    return "SOLID";
            }
        }

        public void MergeCells(Addr start, Addr end)
        {
            int? sheetId = GetActiveSheetId();

            if (!sheetId.HasValue)
            {
                return;
            }

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

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { new Request { MergeCells = mergeCellsRequest } }
            };

            googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
        }


        public void ResizeColumn(Addr addr, int width)
        {
            int? sheetId = GetActiveSheetId();

            if (!sheetId.HasValue)
            {
                return;
            }

            var updateDimensionPropertiesRequest = new UpdateDimensionPropertiesRequest
            {
                Range = new DimensionRange
                {
                    SheetId = sheetId.Value,
                    Dimension = "COLUMNS",
                    StartIndex = addr.Column,
                    EndIndex = addr.Column + 1
                },
                Properties = new DimensionProperties
                {
                    PixelSize = width * 4  // Approximate conversion from Excel's width units to pixels
                },
                Fields = "pixelSize"
            };

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest } }
            };

            googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
        }


        public void ResizeRow(Addr addr, float height)
        {
            int? sheetId = GetActiveSheetId();

            if (!sheetId.HasValue)
            {
                return;
            }

            var updateDimensionPropertiesRequest = new UpdateDimensionPropertiesRequest
            {
                Range = new DimensionRange
                {
                    SheetId = sheetId.Value,
                    Dimension = "ROWS",
                    StartIndex = addr.Row,
                    EndIndex = addr.Row + 1
                },
                Properties = new DimensionProperties
                {
                    PixelSize = (int)(height * 4)  // Approximate conversion from Excel's points to pixels
                },
                Fields = "pixelSize"
            };

            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { new Request { UpdateDimensionProperties = updateDimensionPropertiesRequest } }
            };

            googleSheets.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadSheetId).Execute();
        }


        public void AutoFitAllColumns()
        {
        }


        public void SaveWorkbook()
        {
        }
    }
}