using System;
using System.Collections.Generic;
using System.Windows;
using CloudyWing.Spreadsheet;

namespace ExportSample {
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window {
        public MainWindow() {
            InitializeComponent();
        }

        private void ExportNpoiXls(object sender, RoutedEventArgs e) {
            ExporterBase exporter = new CloudyWing.Spreadsheet.Excel.NPOI.ExcelExporter();
            CreateExportkContent(exporter);
            exporter.ExportFile($@"{txtPath.Text}\{DateTime.Now.ToString("yyyyMMddhhmmss")}{exporter.FileNameExtension}");
        }

        private void ExportNpoiXlsx(object sender, RoutedEventArgs e) {
            ExporterBase exporter = new CloudyWing.Spreadsheet.Excel.NPOI.ExcelExporter(
                CloudyWing.Spreadsheet.Excel.NPOI.ExcelFormat.XLSX
            );
            CreateExportkContent(exporter);
            exporter.ExportFile($@"{txtPath.Text}\{DateTime.Now.ToString("yyyyMMddhhmmss")}{exporter.FileNameExtension}");
        }

        private void ExportEPPlus(object sender, RoutedEventArgs e) {
            ExporterBase exporter = new CloudyWing.Spreadsheet.Excel.EPPlus.ExcelExporter();
            CreateExportkContent(exporter);
            exporter.ExportFile($@"{txtPath.Text}\{DateTime.Now.ToString("yyyyMMddhhmmss")}{exporter.FileNameExtension}");
        }

        private void CreateExportkContent(ExporterBase exporter) {
            Sheeter sheeter = exporter.CreateSheeter();

            // 建立Excel 1~4列(第二列跨三列)
            GridTemplate gt = new GridTemplate();
            gt.CreateRow(25d);
            gt.CreateCell("0,0");
            gt.CreateCell("1,0");

            gt.CreateRow(25d);
            gt.CreateCell("0,1", 2, 3);
            gt.CreateCell("1,1", 1, 1);

            sheeter.AddTemplate(gt);

            // 建立Excel 5~8列
            gt = new GridTemplate();
            gt.CreateRow(25d);
            gt.CreateCell("0,2");
            gt.CreateCell("1,2");
            gt.CreateRow(25d);
            gt.CreateCell("0,3");
            gt.CreateCell("1,3");
            gt.CreateCell(666);
            gt.CreateRow(40d);
            gt.CreateRow(Sheeter.HiddenRowHeight);

            sheeter.AddTemplate(gt);

            // Excel 9~15列(一個父標題、一個主標題和五筆資料)
            ListTemplate<Person> lt = new ListTemplate<Person>();
            lt.HeaderHeight = 40d;
            lt.ItemHeight = 50d;
            lt.DataSource = People;
            lt.Columns.Add("合併標題");
            // 使用DataColumn<T>
            lt.Columns.AddChildToLast(
                "姓名", "name",
                (x, y) => y.Age > 18 ? x + " 大人" : x,
                ListTemplateUtils.HeaderStyle.CloneAndSetBackgroundColor(System.Drawing.Color.Blue)
                    .CloneAndSetFont(CellFont.CreateConfigFont().CloneAndSetColor(System.Drawing.Color.Red).CloneAndSetSize(18)),
                ListTemplateUtils.TextStyle.CloneAndSetWrapText(true)
                    .CloneAndSetBorder(false)
                    .CloneAndSetFont(CellFont.CreateConfigFont().CloneAndSetSize(15).CloneAndSetName("標楷體"))
            );
            // 使用DataColumn<TProperty, TObject>，與前者差別在於ContentRender和ItemStyleFunctor能否指定x參數的型別
            lt.Columns.AddChildToLast<int>(
                "年齡",
                x => x.Age,
                (x, y) => x == 18 ? x - 0.1 : x,
                ListTemplateUtils.HeaderStyle.CloneAndSetBackgroundColor(System.Drawing.Color.Blue)
                    .CloneAndSetFont(
                        CellFont.CreateConfigFont()
                            .CloneAndSetStyle(
                                CloudyWing.Spreadsheet.FontStyles.IsBold |
                                CloudyWing.Spreadsheet.FontStyles.IsItalic |
                                CloudyWing.Spreadsheet.FontStyles.HasUnderline |
                                CloudyWing.Spreadsheet.FontStyles.IsStrikeout
                        )
                ),
                (x, y) => x > 18 ?
                    ListTemplateUtils.NumberStyle.CloneAndSetFont(CellFont.CreateConfigFont().CloneAndSetColor(System.Drawing.Color.Red)) :
                    ListTemplateUtils.NumberStyle
                    .CloneAndSetDataFormat("0.00%")
                    .CloneAndSetFont(CellFont.CreateConfigFont().CloneAndSetColor(System.Drawing.Color.Blue))

            );

            sheeter.AddTemplate(lt);

            sheeter.SetColumnWidth(0, 30);

            // 第二欄自動調整欄寬
            sheeter.SetColumnWidth(1, Sheeter.AutoColumnWidth);

            // 測試健力第二個工作表
            exporter.CreateSheeter("工作表1");
            exporter.LastSheeter.SetColumnWidth(1, Sheeter.HiddenColumnWidth);
        }

        private class Person {
            public string Name { get; set; }
            public int Age { get; set; }
        }

        private IEnumerable<Person> People {
            get {
                yield return new Person { Name = "Wing分身1號", Age = 18 };
                yield return new Person { Name = "Wing分身2號", Age = 19 };
                yield return new Person { Name = "Wing分身3號", Age = 20 };
                yield return new Person { Name = "Wing分身4號", Age = 21 };
                yield return new Person { Name = "Wing分身5號", Age = 22 };
            }
        }
    }
}
