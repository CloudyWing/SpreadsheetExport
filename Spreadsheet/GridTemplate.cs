using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;

namespace CloudyWing.Spreadsheet {

    /// <summary>
    /// 使用CreateRow()與CreateCell()建立Spreadsheet樣板，概念上等同於html的table
    /// </summary>
    public class GridTemplate : ITemplate {

        private IList<CellCollection> rows = new List<CellCollection>();
        private IList<Point> points = new List<Point>();
        private IDictionary<int, double> rowHeights = new Dictionary<int, double>();

        public GridTemplate() { }

        public int ColumnSpan => points.Count == 0 ? 0 : points.Max(x => x.X) + 1;

        public int RowSpan => points.Count == 0 ? 0 : points.Max(x => x.Y) + 1;

        public IEnumerable<Cell> Cells {
            get {
                foreach (CellCollection cells in rows) {
                    foreach (Cell cell in cells) {
                        yield return cell;
                    }
                }
            }
        }

        public IReadOnlyDictionary<int, double> RowHeights => new ReadOnlyDictionary<int, double>(rowHeights);

        public void CreateRow(double height = 0) {
            // 避免建立最後一筆Row，卻沒加入Cell導致用作標算列數會不正確，所以加一個-x座標
            points.Add(new Point(-1, rows.Count));
            rowHeights.Add(rows.Count, height);
            rows.Add(new CellCollection(this));
        }

        /// <exception cref="ArgumentOutOfRangeException">ColumnSpan不可小於1</exception>
        /// <exception cref="ArgumentOutOfRangeException">RowSpan不可小於1</exception>
        public Cell CreateCell(
            object value, int columnSpan = 1, int rowSpan = 1,
            CellStyle? cellStyle = null
        ) {
            if (columnSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnSpan), "ColumnSpan不可小於1");
            }
            if (rowSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowSpan), "RowSpan不可小於1");
            }
            cellStyle = cellStyle ?? CellStyle.CreateConfigStyle();

            if (rows.Count == 0) {
                CreateRow();
            }

            CellCollection lastRow = rows.Last();
            Cell cell = new Cell() {
                Value = value,
                Size = new Size(columnSpan, rowSpan),
                CellStyle = (CellStyle)cellStyle
            };
            lastRow.Add(cell);

            return cell;
        }

        public TemplateContext GetContext() {
            return new TemplateContext(
                Cells, ColumnSpan, RowSpan, RowHeights.ToDictionary(pair => pair.Key, pair => pair.Value)
            );
        }

        private class CellCollection : IEnumerable<Cell>, IEnumerable {
            private GridTemplate grid;
            private IList<Cell> items = new List<Cell>();

            internal CellCollection(GridTemplate grid) {
                this.grid = grid;
            }

            public void Add(Cell item) {
                items.Add(item);

                item.Point = new Point(0, grid.rows.Count - 1);

                while (IsPointExists(item.Point)) {
                    item.Point = item.Point + new Size(1, 0);
                }
                grid.points.Add(item.Point);

                for (int colOffset = 0; colOffset < item.Size.Width; colOffset++) {
                    for (int rowOffset = 0; rowOffset < item.Size.Height; rowOffset++) {
                        // 因為是自己，加入過，所以不用理會
                        if (colOffset == 0 && rowOffset == 0) {
                            continue;
                        }
                        Point point = item.Point + new Size(colOffset, rowOffset);
                        if (IsPointExists(point)) {
                            throw new ArgumentException($"座標({colOffset}, {rowOffset})重複");
                        }
                        grid.points.Add(point);
                    }
                }
            }

            private bool IsPointExists(Point point) {
                return grid.points.Contains(point);
            }

            IEnumerator<Cell> IEnumerable<Cell>.GetEnumerator() {
                return items.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator() {
                return ((IEnumerable)items).GetEnumerator();
            }
        }
    }
}