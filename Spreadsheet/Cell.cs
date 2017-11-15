using System.Drawing;

namespace CloudyWing.Spreadsheet {

    /// <summary>
    /// Excel匯出資料各標題欄位設定
    /// </summary>
    public class Cell {

        public Cell() { }

        /// <summary>
        /// 儲存格內容
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// 座標
        /// </summary>
        public Point Point { get; set; }

        /// <summary>
        /// 欄、列跨度
        /// </summary>
        public Size Size { get; set; } = new Size(1, 1);

        /// <summary>
        /// 儲存格樣式
        /// </summary>
        public CellStyle CellStyle { get; set; }
    }
}