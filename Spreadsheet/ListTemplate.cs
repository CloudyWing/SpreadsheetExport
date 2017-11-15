using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

using static CloudyWing.Spreadsheet.ListTemplateUtils;

namespace CloudyWing.Spreadsheet {

    /// <summary>
    /// 利用設定DataSource和DataColumn來產出資料清單式樣板
    /// </summary>
    /// <typeparam name="T">每筆資料的型別</typeparam>
    public class ListTemplate<T> : ITemplate {

        public IEnumerable<T> DataSource { get; set; }

        public DataColumnCollection<T> Columns { get; } = new DataColumnCollection<T>(null);

        public double HeaderHeight { get; set; }

        public double ItemHeight { get; set; }

        public int ColumnSpan => Columns.ColumnSpan;

        public int RowSpan => DataSource.Count() + Columns.RowSpan;

        public IEnumerable<Cell> Cells => GetHearderCells(Columns).Union(GetItemCells());

        private IEnumerable<Cell> GetHearderCells(DataColumnCollection<T> cols) {
            List<Cell> cells = new List<Cell>();
            foreach (DataColumn<T> col in cols) {
                cells.Add(new Cell() {
                    Value = col.HeaderText,
                    CellStyle = col.HeaderStyle,
                    Point = col.Point,
                    Size = new Size(col.ColumnSpan, col.RowSpan)
                });

                cells.AddRange(GetHearderCells(col.ChildColumns));
            }
            return cells;
        }

        private IEnumerable<Cell> GetItemCells() {
            Point p = new Point(0, Columns.RowSpan);

            foreach (T valueObj in DataSource) {
                foreach (DataColumn<T> col in Columns.DataSourceColumns) {
                    yield return new Cell() {
                        Value = col.GetContentValue(valueObj),
                        CellStyle = col.GetDataCellStyle(valueObj),
                        Point = p,
                        Size = new Size(col.ColumnSpan, 1)
                    };
                    p += new Size(col.ColumnSpan, 0);
                }
                p = new Point(0, p.Y + 1);
            }
        }

        public IReadOnlyDictionary<int, double> RowHeights {
            get {
                int i = 0;
                Dictionary<int, double> dic = new Dictionary<int, double>();
                for (i = 0; i < Columns.RowSpan; i++) {
                    dic.Add(i, HeaderHeight);
                }

                foreach (var item in DataSource) {
                    dic.Add(i++, ItemHeight);
                }

                return dic;
            }
        }

        /// <exception cref="ArgumentNullException">未指定資料來源</exception>
        /// <exception cref="ArgumentNullException">資料來源未包含Property</exception>
        public TemplateContext GetContext() {
            Validate();
            return new TemplateContext(Cells, ColumnSpan, RowSpan, RowHeights);
        }

        /// <exception cref="ArgumentNullException">未指定資料來源</exception>
        /// <exception cref="ArgumentNullException">資料來源未包含Property</exception>
        private void Validate() {
            if (Columns.Count > 0) {
                if (DataSource == null) {
                    throw new ArgumentNullException("未指定資料來源!");
                }

                if (DataSource.Count() == 0) {
                    return;
                }

                IDictionary<string, object> firstData = ConvertToDictionary(DataSource.First());
                foreach (DataColumn<T> col in Columns.DataSourceColumns) {
                    if (!string.IsNullOrWhiteSpace(col.DataKey) && !firstData.ContainsKey(col.DataKey)) {
                        throw new ArgumentNullException($"資料來源未包含Property「{col.DataKey}」!");
                    }
                }
            }
        }
    }
}