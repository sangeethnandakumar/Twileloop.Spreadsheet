using Twileloop.SpreadSheet.Extensions;

public readonly struct Addr
{
    public int Row { get; }
    public int Column { get; }

    public Addr(string address)
    {
        var resolvedAddr = address.ToAddr();
        Row = resolvedAddr.row - 1;
        Column = resolvedAddr.col - 1;
    }

    public Addr((int row, int col) address)
    {
        Row = address.row - 1;
        Column = address.col - 1;
    }

    public static implicit operator Addr(string address) => new Addr(address);

    public static implicit operator Addr((int, int) address) => new Addr(address);

    public Addr MoreRight(int byNColumns)
    {
        return new Addr((Row + 1, Column + byNColumns));
    }

    public Addr MoveBelow(int byNRows)
    {
        return new Addr((Row + byNRows, Column + 1));
    }

    public Addr MoveBelowAndRight(int byNRows, int byNCols)
    {
        return new Addr((Row + byNRows, Column + byNCols));
    }

    public override string ToString()
    {
        return $"{GetExcelColumn(Column + 1)}{Row + 1}";
    }

    private static string GetExcelColumn(int columnNumber)
    {
        string columnName = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            columnName = (char)('A' + (columnNumber % 26)) + columnName;
            columnNumber /= 26;
        }
        return columnName;
    }
}
