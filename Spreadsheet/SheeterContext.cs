using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;

namespace CloudyWing.Spreadsheet {
    public class SheeterContext {
        public SheeterContext(string sheetName, IEnumerable<TemplateContext> templateContexts, IReadOnlyDictionary<int, double> columnWidths) {
            SheetName = sheetName;
            InitializeCellsAndRowHeights(templateContexts);
            ColumnWidths = columnWidths is IDictionary<int, short> ?
                new ReadOnlyDictionary<int, double>((IDictionary<int, double>)columnWidths) :
                columnWidths;
        }

        private void InitializeCellsAndRowHeights(IEnumerable<TemplateContext> templateContexts) {
            List<Cell> cells = new List<Cell>();
            Dictionary<int, double> rowHeights = new Dictionary<int, double>();
            int rowIndex = 0;
            foreach (TemplateContext context in templateContexts) {

                foreach (KeyValuePair<int, double> pair in context.RowHeights) {
                    rowHeights.Add(rowIndex + pair.Key, pair.Value);
                }

                foreach (Cell cell in context.Cells) {
                    Cell fixedCell = new Cell() {
                        Value = cell.Value,
                        Point = new System.Drawing.Point() {
                            X = cell.Point.X,
                            Y = cell.Point.Y + rowIndex
                        },
                        Size = cell.Size,
                        CellStyle = cell.CellStyle
                    };

                    Debug.Assert(fixedCell.Point.Y == cell.Point.Y + rowIndex);

                    cells.Add(fixedCell);
                }

                rowIndex += context.RowSpan;
            }

            Cells = new ReadOnlyCollection<Cell>(cells);
            RowHeights = new ReadOnlyDictionary<int, double>(rowHeights);
        }

        public string SheetName { get; private set; }

        public IReadOnlyList<Cell> Cells { get; private set; }

        public IReadOnlyDictionary<int, double> ColumnWidths { get; private set; }

        public IReadOnlyDictionary<int, double> RowHeights { get; private set; }
    }
}