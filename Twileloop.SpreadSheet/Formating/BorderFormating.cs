using Twileloop.SpreadSheet.Constructs;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Formating
{
    public class BorderFormating : IBorderFormating
    {
        public bool LeftBorder { get; set; }
        public bool RightBorder { get; set; }
        public bool TopBorder { get; set; }
        public bool BottomBorder { get; set; }
        public int Thickness { get; set; }
        public BorderType BorderType { get; set; }
    }
}
