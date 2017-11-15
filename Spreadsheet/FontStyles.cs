using System;

namespace CloudyWing.Spreadsheet {

    [Flags]
    public enum FontStyles {
        None = 0,
        IsBold = 1,
        IsItalic = 2,
        HasUnderline = 4,
        IsStrikeout = 8,
    }
}
