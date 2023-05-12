using System;
using System.Data;

namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetWriter : IDisposable
    {
        public void WriteCell(int row, int column, string data);
        public void WriteCell(string address , string data);
        public void WriteRow(int row, params string[] data);
        public void WriteRow(string address, params string[] data);
        public void WriteColumn(int column, params string[] data);
        public void WriteColumn(string column, params string[] data);
        public void WriteSelection(int startRow, int startColumn, DataTable data);
        public void WriteSelection(string startAddress, DataTable data);
    }
}

