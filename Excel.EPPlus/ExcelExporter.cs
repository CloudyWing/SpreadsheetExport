using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace CloudyWing.Spreadsheet.Excel.EPPlus {
    public class ExcelExporter : ExporterBase {
        private readonly Dictionary<HorizontalAlignment, ExcelHorizontalAlignment> horizontalAlignmentMap = new Dictionary<HorizontalAlignment, ExcelHorizontalAlignment>() {
            [HorizontalAlignment.None] = ExcelHorizontalAlignment.General,
            [HorizontalAlignment.Left] = ExcelHorizontalAlignment.Left,
            [HorizontalAlignment.Center] = ExcelHorizontalAlignment.Center,
            [HorizontalAlignment.Right] = ExcelHorizontalAlignment.Right,
            [HorizontalAlignment.Justify] = ExcelHorizontalAlignment.Justify
        };

        private readonly Dictionary<VerticalAlignment, ExcelVerticalAlignment> verticalAlignmentMap = new Dictionary<VerticalAlignment, ExcelVerticalAlignment>() {
            [VerticalAlignment.Top] = ExcelVerticalAlignment.Top,
            [VerticalAlignment.Middle] = ExcelVerticalAlignment.Center,
            [VerticalAlignment.Bottom] = ExcelVerticalAlignment.Bottom
        };

        public override string ContentType => "application/ms-excel";

        public override string FileNameExtension => ".xlsx";

        protected override byte[] ExecuteExport(SheeterContext[] contexts) {
            using (ExcelPackage ep = new ExcelPackage()) {
                foreach (SheeterContext context in contexts) {
                    CreateSheetToWorkbook(ep.Workbook, context);
                }

                return ep.GetAsByteArray();
            }
        }

        private void CreateSheetToWorkbook(ExcelWorkbook workbook, SheeterContext context) {
            ExcelWorksheet sheet = workbook.Worksheets.Add(context.SheetName);
            sheet.DefaultRowHeight = 16.5;
            SetSheetCell(sheet, context.Cells);
            SetSheetColumnWidths(sheet, context.ColumnWidths);
            SetSheetRowHeights(sheet, context.RowHeights);
        }

        private void SetSheetCell(ExcelWorksheet sheet, IReadOnlyList<Cell> cells) {
            foreach (Cell cell in cells) {
                int startRow = cell.Point.Y + 1;
                int startColumn = cell.Point.X + 1;
                int endRow = cell.Point.Y + cell.Size.Height;
                int endColumn = cell.Point.X + cell.Size.Width;
                ExcelRange range = sheet.Cells[startRow, startColumn, endRow, endColumn];
                range.Merge = true;
                range.Value = cell.Value;
                SetCellStyleToExcel(range.Style, cell.CellStyle);
            }
        }

        private void SetCellStyleToExcel(ExcelStyle excelCellStyle, CellStyle cellStyle) {

            excelCellStyle.HorizontalAlignment = horizontalAlignmentMap[cellStyle.HorizontalAlignment];
            excelCellStyle.VerticalAlignment = verticalAlignmentMap[cellStyle.VerticalAlignment];

            if (cellStyle.HasBorder) {
                excelCellStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
                excelCellStyle.Border.Left.Style = ExcelBorderStyle.Thin;
                excelCellStyle.Border.Right.Style = ExcelBorderStyle.Thin;
                excelCellStyle.Border.Top.Style = ExcelBorderStyle.Thin;
            }

            excelCellStyle.WrapText = cellStyle.WrapText;
            if (cellStyle.BackgroundColor != Color.Empty) {
                excelCellStyle.Fill.PatternType = ExcelFillStyle.Solid;
                excelCellStyle.Fill.BackgroundColor.SetColor(cellStyle.BackgroundColor);
            }

            if (cellStyle.Font != CellFont.Empty) {
                SetFontToExcel(excelCellStyle.Font, cellStyle.Font);
            }

            if (!string.IsNullOrWhiteSpace(cellStyle.DataFormat)) {
                excelCellStyle.Numberformat.Format = cellStyle.DataFormat;
            }
        }

        private void SetFontToExcel(ExcelFont excelFont, CellFont font) {

            if (!string.IsNullOrWhiteSpace(font.Name)) {
                excelFont.Name = font.Name;
            }

            if (font.Size != 0) {
                excelFont.Size = font.Size;
            }

            if (font.Color != Color.Empty) {
                excelFont.Color.SetColor(font.Color);
            }
            excelFont.Bold = (font.Style & FontStyles.IsBold) == FontStyles.IsBold;
            excelFont.Italic = (font.Style & FontStyles.IsItalic) == FontStyles.IsItalic;
            excelFont.UnderLine = (font.Style & FontStyles.HasUnderline) == FontStyles.HasUnderline;
            excelFont.Strike = (font.Style & FontStyles.IsStrikeout) == FontStyles.IsStrikeout;
        }

        private void SetSheetColumnWidths(ExcelWorksheet sheet, IReadOnlyDictionary<int, double> columnWidths) {
            foreach (KeyValuePair<int, double> pair in columnWidths) {
                // EPPlus從1開始算
                ExcelColumn column = sheet.Column(pair.Key + 1);
                if (pair.Value < 0) {
                    column.AutoFit();
                } else if (pair.Value == 0) {
                    column.Hidden = true;
                } else {
                    column.Width = pair.Value;
                }
            }
        }

        private void SetSheetRowHeights(ExcelWorksheet sheet, IReadOnlyDictionary<int, double> rowHeights) {
            foreach (KeyValuePair<int, double> pair in rowHeights) {
                // EPPlus從1開始算
                ExcelRow row = sheet.Row(pair.Key + 1);
                if (pair.Value <= 0) {
                    row.Hidden = true;
                } else {
                    row.Height = pair.Value;
                }
            }
        }
    }
}