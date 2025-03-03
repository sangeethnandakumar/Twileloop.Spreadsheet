using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Factory
{
    public class SpreadSheetAdapter : ISpreadSheetAdapter
    {
        public ISpreadSheetReader Reader { get; set; }
        public ISpreadSheetWriter Writer { get; set; }
        public ISpreadSheetController Controller { get; set; }

        public string DriverName { get; set; }

        public void Dispose()
        {
            Writer.Dispose();
        }
    }
}
