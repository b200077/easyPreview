# easyPreview

一款以 VB.NET / WinForms 開發的 Windows 桌面小工具，整合本機檔案瀏覽、媒體播放與網址收藏管理，方便在同一個視窗裡快速預覽檔案、播放影音、整理常用連結。

## 主要功能

- **本機檔案管理**：瀏覽、預覽資料夾內容
- **媒體播放**：透過 LibVLC 播放影音檔案
- **網頁 / 內嵌瀏覽**：透過 WebView2 顯示網頁內容
- **圖片處理**：以 MagickNET (ImageMagick) 進行影像轉換
- **壓縮檔支援**：透過 SharpCompress 讀取壓縮檔內容
- **圖片轉 PDF**：以 GDI+ 流程將圖片轉為 PDF（透過 PdfSharp / Spire.Pdf 產出）
- **檔案快速搜尋**：整合 Everything64.dll 提供快速搜尋
- **網址收藏管理**：可收集、去除重複、標準化常見網站（YouTube、Twitter/X、Pixiv、DLsite、nhentai、Facebook 等）的網址格式

## 技術棧

| 項目 | 說明 |
| --- | --- |
| 語言 / 框架 | VB.NET, .NET Framework (WinForms) |
| 媒體播放 | LibVLC |
| 網頁元件 | WebView2 |
| 影像處理 | MagickNET / ImageMagick |
| 壓縮檔 | SharpCompress |
| PDF 產出 | PdfSharp 6.2.4、Spire.Pdf |
| 搜尋 | Everything64.dll |
| 其他 | Newtonsoft.Json |

## 資料儲存

- `.trf`：純文字格式的路徑清單
- `.apr`：播放紀錄
- （規劃中）以 SQLite（`Microsoft.Data.Sqlite`）作為未來的補充 / 替代儲存方案

## 開發環境需求

- Windows 11
- Visual Studio（含 VB.NET / WinForms 支援）
- .NET Framework（依專案設定版本）

## 建置與執行

1. Clone 本專案
2. 以 Visual Studio 開啟 `easyPreview.sln`
3. 還原 NuGet 套件後建置（Build）
4. 執行 `EasyPreview.exe`

## 授權

尚未指定授權方式。
