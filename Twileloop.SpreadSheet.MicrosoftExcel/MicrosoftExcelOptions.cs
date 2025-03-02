namespace Twileloop.SpreadSheet.MicrosoftExcel
{
    public class MicrosoftExcelOptions
    {
        public MicrosoftExcelOptions(string fileLocation)
        {
            FileLocation = fileLocation;
        }

        public string FileLocation { get; set; }
    }
}