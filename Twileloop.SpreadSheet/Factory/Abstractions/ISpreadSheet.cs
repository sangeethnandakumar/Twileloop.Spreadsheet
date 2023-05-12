using System;

namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheet : ISpreadSheetReader, ISpreadSheetWriter, ISpreadSheetController, IDisposable
    {

    }
}
