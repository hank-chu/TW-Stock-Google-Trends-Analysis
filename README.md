# TW-Stock-Google-Trends-Analysis
台灣上市公司 Google 搜尋趨勢分析工具

## 專案簡介
此專案通過 Google Trends API 收集並分析台灣上市公司的搜尋數據，包含公司名稱與股票代碼等等的搜尋趨勢分析。通過視覺化呈現，協助使用者了解公司關注度的時間序列變化。

## 功能特點
- 自動批量獲取多家公司的 Google Trends 數據
- 支援自定義時間範圍的數據收集,使用者可自行設定目標時間與關鍵字
- 生成多種視覺化圖表：
  - 趨勢折線圖
  - 搜尋熱度熱力圖
  - 關鍵字相關性分析
  - 季節性分析圖
- 自動處理 API 限制與錯誤處理

## 安裝與使用
1. 克隆專案
```bash
git clone https://github.com/hank-chu/TW-Stock-Google-Trends-Analysis.git
cd TW-Stock-Google-Trends-Analysis
```
2. 準備資料
- 在 `data/input/` 中放入包含公司資訊的 Excel 檔案
- Excel 檔案格式：每個欄位包含不同類型的關鍵字（如公司名稱、股票代碼）
- 將包含公司名稱和/或股票代碼的 Excel 檔案放入項目資料夾，命名為 example.xlsx（或更改程式中的檔名參數）。
- Excel 檔案的格式範例：

  
![Excel檔案範例](https://github.com/user-attachments/assets/faf39d62-e2b3-4eb1-ac6a-b2e1f4b5f1af)















