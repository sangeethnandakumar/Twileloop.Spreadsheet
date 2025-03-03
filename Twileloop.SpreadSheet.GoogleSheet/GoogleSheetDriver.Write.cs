using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Data;
using Twileloop.SpreadSheet.Styling;

namespace Twileloop.SpreadSheet.GoogleSheet
{
    public partial class GoogleSheetDriver
    {
        public void WriteCell(Addr addr, string data, SpreadsheetStyling style = null)
        {
            var (start, end) = GetMergedCellRange(addr);
            string range = $"{currentSheetName}!{start}:{end}";

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

        public void WriteColumn(Addr start, string[] data, SpreadsheetStyling style = null)
        {
            var end = start.MoveBelow(data.Length);
            string range = $"{currentSheetName}!{start}:{end}";

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
                    Addr endAddr = (start.Row + data.Length - 1, start.Column);
                    QueueStylingRequest(start, endAddr, style);
                }
            }
            else
            {
                var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateRequest.Execute();

                if (style != null)
                {
                    Addr endAddr = (start.Row + data.Length - 1, start.Column);
                    ApplyStylingImmediate(start, endAddr, style);
                }
            }
        }

        public void WriteRow(Addr start, string[] data, SpreadsheetStyling style = null)
        {
            var end = start.MoreRight(data.Length);
            string range = $"{currentSheetName}!{start}:{end}";

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
                    Addr endAddr = (start.Row, start.Column + data.Length - 1);
                    QueueStylingRequest(start, endAddr, style);
                }
            }
            else
            {
                var updateRequest = googleSheets.Spreadsheets.Values.Update(valueRange, spreadSheetId, range);
                updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                updateRequest.Execute();

                if (style != null)
                {
                    Addr endAddr = (start.Row, start.Column + data.Length - 1);
                    ApplyStylingImmediate(start, endAddr, style);
                }
            }
        }

        public void WriteTable(Addr startAddr, DataTable data, SpreadsheetStyling style = null)
        {
            // Execute immediately regardless of BulkUpdate setting
            int rowCount = data.Rows.Count;
            int columnCount = data.Columns.Count;
            var endAddrs = startAddr.MoveBelowAndRight(rowCount, columnCount);

            string range = $"{currentSheetName}!{startAddr}:{endAddrs}";

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
                ApplyStylingImmediate(startAddr, endAddrs, style);
            }
        }

        public void ApplyStyling(Addr start, Addr end, SpreadsheetStyling styling)
        {
            // Execute immediately regardless of BulkUpdate setting
            ApplyStylingImmediate(start, end, styling);
        }

        private void QueueStylingRequest(Addr start, Addr end, SpreadsheetStyling styling)
        {
            if (!sheetId.HasValue) return;

            var gridRange = new GridRange
            {
                SheetId = sheetId.Value,
                StartRowIndex = start.Row,
                EndRowIndex = end.Row + 1,  // Make exclusive
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 1  // Make exclusive
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
                EndRowIndex = end.Row + 1,  // Make exclusive
                StartColumnIndex = start.Column,
                EndColumnIndex = end.Column + 1  // Make exclusive
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
                    EndIndex = addr.Column + 1  // Make sure this is exclusive
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
                    EndIndex = addr.Row + 1  // Make sure this is exclusive
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
    }
}