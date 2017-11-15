using System.Drawing;
using CloudyWing.Spreadsheet.Properties;

namespace CloudyWing.Spreadsheet {

    public struct CellFont {

        public static CellFont Empty = new CellFont();

        public static bool operator ==(CellFont font1, CellFont font2) {
            return font1.Equals(font2);
        }

        public static bool operator !=(CellFont font1, CellFont font2) {
            return !font1.Equals(font2);
        }

        public CellFont(
            string name, short size = 0, Color? color = null, FontStyles style = FontStyles.None
        ) {
            color = color ?? Color.Empty;
            Name = name;
            Size = size;
            Color = (Color)color;
            Style = style;
        }

        /// <summary>
        /// 字體名稱
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 字體大小
        /// </summary>
        public short Size { get; private set; }

        public Color Color { get; private set; }

        /// <summary>
        /// 字體樣式，如粗體、斜體等
        /// </summary>
        public FontStyles Style { get; private set; }

        /// <summary>
        /// 建立一個不同Name的副本
        /// </summary>
        /// <param name="name">字體名稱</param>
        public CellFont CloneAndSetName(string name) {
            CellFont font = this;
            font.Name = name;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本Size
        /// </summary>
        /// <param name="size">字體大小</param>
        public CellFont CloneAndSetSize(short size) {
            CellFont font = this;
            font.Size = size;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本Color
        /// </summary>
        /// <param name="color">字體顏色</param>
        public CellFont CloneAndSetColor(Color color) {
            CellFont font = this;
            font.Color = color;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本字體樣式
        /// </summary>
        /// <param name="style">副本字體樣式</param>
        public CellFont CloneAndSetStyle(FontStyles style) {
            CellFont font = this;
            font.Style = style;
            return font;
        }

        public override bool Equals(object obj) {
            if (obj == null) {
                return false;
            }
            if (obj is CellFont) {
                return Equals((CellFont)obj);
            }
            return false;
        }

        public bool Equals(CellFont font) {
            return Name == font.Name &&
                Size == font.Size &&
                Color == font.Color &&
                Style == font.Style;
        }

        public override int GetHashCode() {
            return base.GetHashCode();
        }

        public static CellFont CreateConfigFont() {
            FontStyles style = FontStyles.None;
            if (Settings.Default.IsBold) {
                style = style | FontStyles.IsBold;
            }

            if (Settings.Default.IsItalic) {
                style = style | FontStyles.IsItalic;
            }

            if (Settings.Default.HasUnderline) {
                style = style | FontStyles.HasUnderline;
            }

            if (Settings.Default.IsStrikeout) {
                style = style | FontStyles.IsStrikeout;
            }

            return new CellFont(
                Settings.Default.FontName,
                Settings.Default.FontSize,
                Settings.Default.FontColor,
                style
            );
        }
    }
}