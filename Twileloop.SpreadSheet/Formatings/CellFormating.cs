﻿using System.Drawing;
using Twileloop.SpreadSheet.Factory.Abstractions;

namespace Twileloop.SpreadSheet.Formatings
{
    public class CellFormating : ICellFormating
    {
        public Color BackgroundColor { get; set; }
        public BorderFormating BorderFormat { get; set; }
    }
}