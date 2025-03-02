using System;
using System.Data;
using Twileloop.SpreadSheet.Styling;

namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetWriter : IDisposable
    {
        void WriteCell(Addr start, string data, SpreadsheetStyling styling = null);
        void WriteRow(Addr start, string[] data, SpreadsheetStyling styling = null);
        void WriteColumn(Addr start, string[] data, SpreadsheetStyling styling = null);
        void WriteTable(Addr start, DataTable data, SpreadsheetStyling styling = null);
        void ResizeColumn(Addr addr, int width);
        void ResizeRow(Addr addr, float height);
        void AutoFitAllColumns();
        void MergeCells(Addr start, Addr end);
        void ApplyStyling(Addr start, Addr end, SpreadsheetStyling styling);
        void ApplyBorder(Addr start, Addr end, BorderStyling styling);
    }
}
