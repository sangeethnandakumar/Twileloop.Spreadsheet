using Newtonsoft.Json;
using System.Drawing;
using System.IO;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Styling
{
    public class SpreadsheetStyling : IFormatting
    {
        public TextStyling TextFormating { get; set; } = new();
        public CellStyling CellFormating { get; set; } = new();
    }

    public class TextStyling : ITextFormating
    {
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public int Size { get; set; } = 11;
        public Color FontColor { get; set; } = Color.Black;
        public string Font { get; set; } = "Arial";
        public VerticalTxtAlignment VerticalAlignment { get; set; } = VerticalTxtAlignment.MIDDLE;
        public HorizontalTxtAlignment HorizontalAlignment { get; set; } = HorizontalTxtAlignment.LEFT;
    }

    public class CellStyling : ICellFormating
    {
        public Color BackgroundColor { get; set; } = Color.Transparent;
    }

    public class BorderStyling : IBorderFormating
    {
        public bool LeftBorder { get; set; }
        public bool RightBorder { get; set; }
        public bool TopBorder { get; set; }
        public bool BottomBorder { get; set; }
        public BorderThickness Thickness { get; set; } = BorderThickness.Thin;
        public Color BorderColor { get; set; } = Color.Black;
        public BorderType BorderType { get; set; }
    }

    public enum BorderThickness
    {
        Thin,
        Medium,
        Thick,
        DoubleLined
    }

    public enum VerticalTxtAlignment { TOP, MIDDLE, BOTTOM }
    public enum HorizontalTxtAlignment { LEFT, CENTER, RIGHT }
    public enum BorderType { SOLID, DOTTED, DASHED }

    public class StyleBuilder
    {
        private readonly SpreadsheetStyling _formatting;

        public StyleBuilder() => _formatting = new SpreadsheetStyling();

        public StyleBuilder Bold() { _formatting.TextFormating.Bold = true; return this; }
        public StyleBuilder Italic() { _formatting.TextFormating.Italic = true; return this; }
        public StyleBuilder Underline() { _formatting.TextFormating.Underline = true; return this; }
        public StyleBuilder WithFontSize(int size) { _formatting.TextFormating.Size = size; return this; }
        public StyleBuilder WithFont(string font) { _formatting.TextFormating.Font = font; return this; }
        public StyleBuilder WithTextColor(Color color) { _formatting.TextFormating.FontColor = color; return this; }
        public StyleBuilder WithTextAllignment(HorizontalTxtAlignment h, VerticalTxtAlignment v) { 
            _formatting.TextFormating.HorizontalAlignment = h; 
            _formatting.TextFormating.VerticalAlignment = v; 
            return this; }
        public StyleBuilder WithBackgroundColor(Color color) { _formatting.CellFormating.BackgroundColor = color; return this; }

        public SpreadsheetStyling Build() => _formatting;

        public string ToJson() => JsonConvert.SerializeObject(_formatting, Formatting.Indented);

        public void SaveToFile(string filePath) => File.WriteAllText(filePath, ToJson());

        public static SpreadsheetStyling LoadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Style file not found.");
            return JsonConvert.DeserializeObject<SpreadsheetStyling>(File.ReadAllText(filePath));
        }
    }
}