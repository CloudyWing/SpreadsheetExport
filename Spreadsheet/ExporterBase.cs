using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CloudyWing.Spreadsheet {
    /// <summary>
    /// 電子表格，用來設定電子表格內容將之輸出至Exporter
    /// </summary>
    public abstract class ExporterBase {
        private readonly IList<Sheeter> sheeters = new List<Sheeter>();

        /// <summary>
        /// 取得最新的工作表，若無則建立一個工作表
        /// </summary>
        public Sheeter LastSheeter => sheeters.LastOrDefault() ?? CreateSheeter(null);

        public abstract string ContentType { get; }

        public abstract string FileNameExtension { get; }

        public Sheeter CreateSheeter(string sheetName = "") {
            if (string.IsNullOrWhiteSpace(sheetName)) {
                sheetName = GetDefaultSheetName();
            } else if (IsSheetNameExists(sheetName)) {
                sheetName = FixSheetName(sheetName);
            }

            Sheeter sheeter = new Sheeter(sheetName);
            sheeters.Add(sheeter);

            return sheeter;
        }

        private bool IsSheetNameExists(string sheetName) =>
            sheeters.Select(x => x.SheetName).Contains(sheetName);

        private string GetDefaultSheetName() {
            string baseSheetName = "工作表";
            string defaultSheetName;
            int i = 1;
            do {
                defaultSheetName = baseSheetName + i++;
            } while (IsSheetNameExists(defaultSheetName));

            return defaultSheetName;
        }

        private string FixSheetName(string sheetName) {
            string fixedSheetName;
            int i = 1;
            do {
                fixedSheetName = $"{sheetName}({i++})";
            } while (IsSheetNameExists(fixedSheetName));

            return fixedSheetName;
        }

        /// <summary>
        /// 電子表格匯出
        /// </summary>
        /// <exception cref="ArgumentNullException">未建立任何工作表。</exception>
        public byte[] Export() {
            Validate();
            SheeterContext[] contexts = new SheeterContext[sheeters.Count];
            for (int i = 0; i < sheeters.Count; i++) {
                Sheeter sheeter = sheeters[i];

                contexts[i] = new SheeterContext(
                    sheeter.SheetName,
                    sheeter.Templates.Select(x => x.GetContext()),
                    sheeter.ColumnWidths.ToDictionary(x => x.Key, x => x.Value)
                );
            }
            return ExecuteExport(contexts);
        }

        protected abstract byte[] ExecuteExport(SheeterContext[] contexts);

        /// <exception cref="NullReferenceException">未建立任何工作表。</exception>
        private void Validate() {
            if (sheeters.Count == 0) {
                throw new NullReferenceException("未建立任何工作表。");
            }
        }

        /// <summary>
        /// 匯出至檔案，若檔案已存在則覆寫
        /// </summary>
        /// <param name="path">欲儲存檔案路徑</param>
        /// <exception cref="NullReferenceException">未建立任何工作表。</exception>
        public void ExportFile(string path) {
            using (FileStream fileStream = new FileStream(path, FileMode.Create, FileAccess.Write)) {
                byte[] bytes = Export();
                fileStream.Write(bytes, 0, bytes.Length);
            }
        }
    }
}