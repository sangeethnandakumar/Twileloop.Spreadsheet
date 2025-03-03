using Twileloop.SpreadSheet.Factory.Base;

namespace Twileloop.SpreadSheet.Factory
{
    public static class SpreadSheetFactory
    {
        public static SpreadSheetAdapter CreateAdapter(ISpreadSheetDriver driver)
        {
            var accessor = new SpreadSheetAdapter();
            accessor.Reader = driver;
            accessor.Writer = driver;
            accessor.Controller = driver;
            accessor.DriverName = driver.DriverName;
            return accessor;
        }
    }
}
