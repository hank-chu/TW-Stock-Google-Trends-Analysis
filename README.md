# TW-Stock-Google-Trends-Analysis
台灣上市公司搜尋趨勢 (Google trends) 分析工具

## 專案簡介
此專案透過 Google Trends API 自動化擷取台灣上市公司的搜尋趨勢資料。目標是收集特定時間內台灣上市公司（包含公司名稱和股票代碼等等變數）的網路搜尋熱度，並將資料視覺化，為市場情緒分析、投資人關注度和股市趨勢研究提供參考。
此專案適合對市場分析、趨勢研究或股市動向有興趣的開發者、金融分析師或數據科學家，亦可作為網路搜尋趨勢與市場波動關係的分析基礎。

## 關鍵字 / Keywords
- Google Trends API
- pytrends

## 使用情境
金融趨勢分析：觀察台灣股市特定公司的搜尋熱度波動，分析市場情緒,投資人關注度變化。
行銷研究：透過網路搜尋熱度變化，檢視特定公司在消費者間的關注度。
數據視覺化：提供豐富的圖表呈現方式，直觀展示公司搜尋趨勢。

## 功能特點
- 透過Google Trends Api, pytrends自動獲取多家公司的 Google Trends 數據
- 支援自定義時間範圍的數據收集,使用者可自行設定目標時間與關鍵字
- 生成多種視覺化圖表：
  - 趨勢折線圖
  - 搜尋熱度熱力圖
  - 關鍵字相關性分析
  - 季節性分析圖
- 自動處理 API 限制與錯誤處理

## 項目結構
```
TW-Stock-Trends-Analysis/
├── src/
│   ├── data_collection.py        # Google Trends 數據收集
│   ├── data_visualization.py     # 數據視覺化
├── Google_Trends_Output          # google trends的結果輸出位置
├── trend_plots                   # 視覺化圖表輸出位置
├── requirements.txt              # 依賴套件
└── README.md                     # 專案文檔
```

## 安裝與使用
1. 克隆專案
```bash
git clone https://github.com/hank-chu/TW-Stock-Google-Trends-Analysis.git
cd TW-Stock-Google-Trends-Analysis
```

2. 安裝相關套件
```bash
pip install -r requirements.txt
```

3. 準備資料
- 將包含公司名稱或股票代碼的 Excel 檔案放入TW-Stock-Trends-Analysis資料夾，命名為 example.xlsx（或更改程式中的檔名參數）。
- Excel 檔案格式：每個欄位包含不同類型的關鍵字（如公司名稱、股票代碼）,範例檔案已提供欄位名稱與格式，可根據需求自行編輯。
- Excel 檔案的格式範例：

  
![Excel檔案範例](https://github.com/user-attachments/assets/faf39d62-e2b3-4eb1-ac6a-b2e1f4b5f1af)


4. 執行分析
透過 data_collection.py 擷取公司搜尋趨勢數據，結果將儲存至 Google_Trends_Output 資料夾：
```bash
python src/data_collection.py  # 收集google trends數據
```

5. 生成視覺化圖表
執行 data_visualization.py，自動生成並儲存四種圖表至 trend_plots 資料夾：
```bash
python src/data_visualization.py  # 生成視覺化圖片
```

## 視覺化圖表說明
趨勢折線圖：展示特定公司在各時間點的搜尋趨勢變化。

熱度熱力圖：以月為單位，顯示每家公司搜尋熱度的分布情況。

關鍵字相關性圖：公司間搜尋熱度的相關係數圖，適合用於分析不同公司間的趨勢相似性。

季節性分析圖：呈現搜尋趨勢的季節性特徵，適合觀察是否存在週期性趨勢。


## 範例圖表
- **趨勢折線圖**
  ![公司中文全稱_trend_line](https://github.com/user-attachments/assets/87cd2c6a-60c3-4db6-b7d1-a17e42f9a782)

- **熱度熱力圖**
  ![公司中文全稱_heatmap](https://github.com/user-attachments/assets/7fdf32c7-b8e7-433e-b000-fc773561569a)

- **關鍵字相關性圖**
  ![公司中文全稱_correlation](https://github.com/user-attachments/assets/01805f3c-0526-4fc9-a397-a5a201920d23)

- **季節性分析圖**
  ![公司中文全稱_seasonal](https://github.com/user-attachments/assets/3de083a4-7674-49e4-ab57-9fa75ea29ade)



## 注意事項
API 限制：Google Trends API 對於每次請求有數量限制，因此程式包含了等待時間以避免被限制,如遇到 "The request failed: Google returned a response with code 429" 問題，代表太頻繁送出要求,請等待一段時間過後重新嘗試。












