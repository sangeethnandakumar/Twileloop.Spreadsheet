using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Formating
{
    public class Formatting : IFormatting
    {
        public TextFormating TextFormating { get; set; }
        public CellFormating CellFormating { get; set; }
        public BorderFormating BorderFormating { get; set; }
    }
}
