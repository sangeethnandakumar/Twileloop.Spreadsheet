namespace Twileloop.SpreadSheet.Factory.Abstractions
{
    public interface ISpreadSheetController
    {
        void InitialiseWorkbook();
        void OpenSheet(string sheetName);
        void CreateSheets(params string[] sheetNames);
        string[] GetSheets();
        string GetActiveSheet();
        void SaveWorkbook();
    }

}
