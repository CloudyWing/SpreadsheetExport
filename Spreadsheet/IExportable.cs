using System.Collections.Generic;
using System.IO;

namespace CloudyWing.Spreadsheet {

    public interface IExportable {

        string ContentType { get; }

        string FileNameExtension { get; }

        byte[] Export(IEnumerable<SheeterContext> contexts);
    }
}
