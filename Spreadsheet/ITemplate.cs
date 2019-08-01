using System.Collections.Generic;

namespace CloudyWing.Spreadsheet {
    public interface ITemplate {
        IReadOnlyDictionary<int, double> RowHeights { get; }

        TemplateContext GetContext();
    }
}