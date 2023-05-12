namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetController
    {
        void LoadSheet(string sheetName);
        void CreateSheets(params string[] sheetNames);
        string[] GetSheets();
        string GetActiveSheet();
    }

}
