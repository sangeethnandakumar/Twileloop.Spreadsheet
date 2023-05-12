using System.Data;

namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetReader
    {
        public string ReadCell(int row, int column);
        public string ReadCell(string address);
        public string[] ReadRow(int row);
        public string[] ReadRow(string address);
        public string[] ReadColumn(int column);
        public string[] ReadColumn(string column);
        public DataTable ReadSelection(int startRow, int startColumn, int endRow, int endColumn);
        public DataTable ReadSelection(string startAddress, string endAddress);
    }
}
