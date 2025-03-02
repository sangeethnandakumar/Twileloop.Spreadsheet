using System.Data;
using Twileloop.SpreadSheet.Extensions;

namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetReader
    {
        string ReadCell(Addr addr);
        string[] ReadRow(Addr start);
        string[] ReadColumn(Addr start);
        DataTable ReadSelection(Addr start, Addr end);
    }
}
