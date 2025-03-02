using System;
using System.Text.RegularExpressions;

namespace Twileloop.SpreadSheet.Extensions
{

    public static class SheetHelpers
    {
        public static (int row, int col) ToAddr(this string address)
        {
            var match = Regex.Match(address, @"^([A-Z]+)(\d+)$");
            if (!match.Success) throw new ArgumentException("Invalid address format");

            string colLetters = match.Groups[1].Value;
            int row = int.Parse(match.Groups[2].Value);
            int col = colLetters.ToColumnNumber();

            return (row, col);
        }

        public static string ToAddr(this (int row, int col) addr)
        {
            if (addr.row < 1 || addr.col < 1) throw new ArgumentException("Row and column must be greater than zero");
            return addr.col.ToColumnLetter() + addr.row;
        }

        public static int ToColumnNumber(this string column)
        {
            int colNum = 0;
            foreach (char c in column)
            {
                colNum = colNum * 26 + (c - 'A' + 1);
            }
            return colNum;
        }

        public static string ToColumnLetter(this int column)
        {
            if (column < 1) throw new ArgumentException("Column number must be greater than zero");

            string colLetter = "";
            while (column > 0)
            {
                column--;
                colLetter = (char)('A' + column % 26) + colLetter;
                column /= 26;
            }
            return colLetter;
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            var (row, col) = "B12".ToAddr();
            Console.WriteLine($"'B12'.ToAddr() -> Row: {row}, Col: {col}");

            string sheetAddress = (5, 28).ToAddr();
            Console.WriteLine($"(5, 28).ToAddr() -> {sheetAddress}");
        }
    }

}
