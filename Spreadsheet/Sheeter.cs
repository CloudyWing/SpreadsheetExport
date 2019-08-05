using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace CloudyWing.Spreadsheet {
    /// <summary>
    /// 紀錄工作表的格式
    /// </summary>
    public class Sheeter {
        public const double AutoColumnWidth = -1d;
        public const double HiddenColumnWidth = 0d;
        public const double HiddenRowHeight = 0d;

        private readonly IDictionary<int, double> columnWidths = new Dictionary<int, double>();
        private readonly IList<ITemplate> templates = new List<ITemplate>();

        internal Sheeter(string sheetName) {
            SheetName = sheetName;
        }

        public string SheetName { get; set; }

        public IReadOnlyDictionary<int, double> ColumnWidths => new ReadOnlyDictionary<int, double>(columnWidths);

        public IReadOnlyList<ITemplate> Templates => new ReadOnlyCollection<ITemplate>(templates);

        /// <summary>
        /// 設定欄位寬度
        /// </summary>
        /// <param name="index">第幾欄</param>
        /// <param name="width">欄寬度，0表示default，-1為自動調整寬度</param>
        public void SetColumnWidth(int index, double width) {
            columnWidths[index] = width;
        }

        public void AddTemplate(ITemplate template) {
            templates.Add(template);
        }
    }
}