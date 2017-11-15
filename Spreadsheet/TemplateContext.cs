using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CloudyWing.Spreadsheet {

    public class TemplateContext {

        public TemplateContext(
            IEnumerable<Cell> cells, int columnSpan, int rowSpan, IReadOnlyDictionary<int, double> rowHeights
        ) {
            Cells = new ReadOnlyCollection<Cell>(cells.ToList());
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            if (rowHeights is IDictionary<int, short>) {
                RowHeights = new ReadOnlyDictionary<int, double>((IDictionary<int, double>)rowHeights);
            } else {
                RowHeights = rowHeights;
            }
        }

        public IReadOnlyList<Cell> Cells { get; private set; }

        public int ColumnSpan { get; private set; }

        public int RowSpan { get; private set; }

        public IReadOnlyDictionary<int, double> RowHeights { get; private set; }
    }
}
