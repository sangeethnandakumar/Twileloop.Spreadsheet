using System;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Factory
{
    public interface ISpreadSheetAdapter : IDisposable
    {
        public string DriverName { get; set; }
        ISpreadSheetReader Reader { get; set; }
        ISpreadSheetWriter Writer { get; set; }
        ISpreadSheetController Controller { get; set; }
    }
}