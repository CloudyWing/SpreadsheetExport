# SpreadsheetExport

用來封裝NPOI或EPPlus來進行Excel匯出，使用上並不會比使用原套件簡略多少，目的更多是想要統一Excel匯出的操作，並看是否為了能再加上其他類型的匯出。

## 專案說明
SpreadsheetExport：`主專案，用於定義介面，與抽象類別、通用類別。`

Excel.NPOI：`封裝NPOI進行Excel匯出。`

Excel.EPPlus：`封裝EPPlus進行Excel匯出。`

ExportSample：`Sample專案。`

目前使用套件：
```
NPOI.2.4.1
SharpZipLib.1.1.0
EPPlus.4.5.3.2
```