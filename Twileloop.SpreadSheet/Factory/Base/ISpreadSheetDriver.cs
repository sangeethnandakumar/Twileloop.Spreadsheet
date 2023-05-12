using System;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Factory.Base
{
    public interface ISpreadSheetDriver : ISpreadSheetReader, ISpreadSheetWriter, ISpreadSheetController, IDisposable
    {

    }
}
