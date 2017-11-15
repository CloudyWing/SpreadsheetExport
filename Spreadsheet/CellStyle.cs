using System;
using System.Drawing;
using CloudyWing.Spreadsheet.Properties;

namespace CloudyWing.Spreadsheet {

    public struct CellStyle {

        public static CellStyle Empty = new CellStyle();

        public static bool operator ==(CellStyle style1, CellStyle style2) {
            return style1.Equals(style2);
        }

        public static bool operator !=(CellStyle style1, CellStyle style2) {
            return !style1.Equals(style2);
        }

        public CellStyle(
            HorizontalAlignment halign, VerticalAlignment valign = VerticalAlignment.Middle,
            bool hasBorder = false, bool wrapText = false, Color? backgroundColor = null,
            CellFont? font = null
        ) : this() {
            backgroundColor = backgroundColor ?? Color.Empty;
            font = font ?? CellFont.Empty;
            HorizontalAlignment = halign;
            VerticalAlignment = valign;
            HasBorder = hasBorder;
            WrapText = wrapText;
            BackgroundColor = (Color)backgroundColor;
            Font = (CellFont)font;
        }

        /// <summary>
        /// 水平對齊
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; private set; }

        /// <summary>
        /// 垂直對齊
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; private set; }

        /// <summary>
        /// 是否有格線
        /// </summary>
        public bool HasBorder { get; private set; }

        /// <summary>
        /// 是否自動換行
        /// </summary>
        public bool WrapText { get; private set; }

        /// <summary>
        /// 背景顏色
        /// </summary>
        public Color BackgroundColor { get; private set; }

        /// <summary>
        /// 字體格式
        /// </summary>
        public CellFont Font { get; private set; }

        /// <summary>
        /// 資料格式字串
        /// </summary>
        public string DataFormat { get; private set; }

        /// <summary>
        /// 建立一個副本，並設定副本的的水平對齊
        /// </summary>
        /// <param name="align">水平對齊</param>
        public CellStyle CloneAndSetHorizontalAlignment(HorizontalAlignment align) {
            CellStyle style = this;
            style.HorizontalAlignment = align;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本的的垂直對齊
        /// </summary>
        /// <param name="valign">垂直對齊</param>
        public CellStyle CloneAndSetVerticalAlignment(VerticalAlignment valign) {
            CellStyle style = this;
            style.VerticalAlignment = valign;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本是否顯示框線
        /// </summary>
        /// <param name="hasBolder">是否有框線</param>
        public CellStyle CloneAndSetBorder(bool hasBolder) {
            CellStyle style = this;
            style.HasBorder = hasBolder;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本是否自動換行
        /// </summary>
        /// <param name="wrapText">是否自動換行</param>
        public CellStyle CloneAndSetWrapText(bool wrapText) {
            CellStyle style = this;
            style.WrapText = wrapText;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本背景顏色
        /// </summary>
        /// <param name="backgroundColor">背景顏色</param>
        public CellStyle CloneAndSetBackgroundColor(Color backgroundColor) {
            CellStyle style = this;
            style.BackgroundColor = backgroundColor;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本字體
        /// </summary>
        /// <param name="font">字體</param>
        public CellStyle CloneAndSetFont(CellFont font) {
            CellStyle style = this;
            style.Font = font;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本的資料格式字串
        /// </summary>
        /// <param name="font">字體</param>
        public CellStyle CloneAndSetDataFormat(string dataForamt) {
            CellStyle style = this;
            style.DataFormat = dataForamt;
            return style;
        }

        public override bool Equals(object obj) {
            if (obj == null) {
                return false;
            }

            if (obj is CellStyle) {
                return Equals((CellStyle)obj);
            }
            return false;
        }

        public bool Equals(CellStyle style) {
            return HorizontalAlignment == style.HorizontalAlignment &&
                VerticalAlignment == style.VerticalAlignment &&
                HasBorder == style.HasBorder &&
                WrapText == style.WrapText &&
                BackgroundColor == style.BackgroundColor &&
                Font == style.Font;
        }

        public override int GetHashCode() {
            return base.GetHashCode();
        }

        public static CellStyle CreateConfigStyle() {
            return new CellStyle(
                Settings.Default.HorizontalAlignment,
                Settings.Default.VerticalAlignment,
                Settings.Default.HasBorder,
                Settings.Default.WrapText,
                Color.Empty,
                CellFont.CreateConfigFont()
            );
        }
    }
}