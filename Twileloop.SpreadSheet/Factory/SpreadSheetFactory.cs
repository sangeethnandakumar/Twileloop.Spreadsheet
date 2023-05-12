using Twileloop.SpreadSheet.Factory.Base;

namespace Twileloop.SpreadSheet.Factory
{
    public static class SpreadSheetFactory
    {
        public static SpreadSheetAccessor CreateAccessor(ISpreadSheetDriver driver)
        {
            var accessor = new SpreadSheetAccessor();
            accessor.Reader = driver;
            accessor.Writer = driver;
            accessor.Controller = driver;
            return accessor;
        }
    }
}
