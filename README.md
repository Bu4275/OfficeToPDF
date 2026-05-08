# OfficeToPDF

將 Microsoft Office 文件 (doc / docx / xls / xlsx) 批次轉換為 PDF。

## 功能

- 支援 Word (.doc, .docx) 和 Excel (.xls, .xlsx) 轉換為 PDF
- 資料夾拖曳或瀏覽選擇
- 可調整並行執行緒數 (1-4)
- 自動產生輸出資料夾，不覆蓋已有結果
- 深色主題 UI
- 即時進度與錯誤紀錄

## 需求

- Windows 10 / 11
- Microsoft Office (Word / Excel) 已安裝
- **不需要安裝 .NET Runtime**（exe 已自帶）

## 使用方式

1. 從 [Releases](https://github.com/Bu4275/OfficeToPDF/releases) 下載 `OfficeToPDF.exe`
2. 直接執行，選擇或拖入包含 Office 文件的資料夾
3. 點擊「開始轉換」

輸出的 PDF 會自動存放在同層目錄下的 `<資料夾名稱>-pdf` 資料夾中。

## 從原始碼建置

```bash
dotnet publish -c Release
```

輸出位置：`bin\Release\net9.0-windows\win-x64\publish\OfficeToPDF.exe`

## 授權

MIT License
