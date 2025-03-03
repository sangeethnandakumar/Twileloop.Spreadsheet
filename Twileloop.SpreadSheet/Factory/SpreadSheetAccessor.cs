using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Factory
{
    public class SpreadSheetAccessor : ISpreadSheetAdapter
    {
        public ISpreadSheetReader Reader { get; set; }
        public ISpreadSheetWriter Writer { get; set; }
        public ISpreadSheetController Controller { get; set; }

        public void Dispose()
        {
            Writer.Dispose();
        }
    }
}
