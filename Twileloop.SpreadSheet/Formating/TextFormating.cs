using System.Drawing;
using Twileloop.SpreadSheet.Constructs;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Formating
{
    public class TextFormating : ITextFormating
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public int Size { get; set; }
        public Color Color { get; set; }
        public string Font { get; set; }
        public VerticalAllignment VerticalAlignment { get; set; } = VerticalAllignment.TOP;
        public HorizontalAllignment HorizontalAlignment { get; set; } = HorizontalAllignment.LEFT;
    }
}
