import pandas as pd
from pytrends.request import TrendReq
import time
import os

# 初始化 Google Trends API
# hl='zh-TW': 設定語言為繁體中文
# tz=480: 設定時區為 UTC+8（台灣時區）
pytrends = TrendReq(hl='zh-TW', tz=480)

# 讀取包含關鍵字的 Excel 檔案
# 檔案中應包含要查詢的公司名稱或股票代碼
file_path = "example.xlsx"  # 請替換為實際的 Excel 檔案路徑
data_frame = pd.read_excel(file_path)

# 建立輸出資料夾，用於存放結果
# exist_ok=True: 若資料夾已存在則不會報錯
output_folder = "Google_Trends_Output"
os.makedirs(output_folder, exist_ok=True)

# 取得 Excel 檔案中所有欄位的名稱
# 每個欄位應該包含不同類型的關鍵字（如公司名稱、股票代碼等）
column_names = data_frame.columns

# 針對每個欄位進行處理
for column in column_names:
    # 取得當前欄位的所有資料，並移除空值
    # 將資料轉換為列表形式
    column_data = data_frame[column].dropna().values.tolist()
    
    # 將所有數值轉換為字串格式
    # 因為Google Trends API 需要字串形式的關鍵字
    column_data = [str(item) for item in column_data]
    
    # 顯示當前處理進度
    print(f"正在處理欄位: {column}")
    print("當前欄位資料:", column_data)
    print("-" * 20)
    
    # 將關鍵字分組，每組最多 5 個
    # Google Trends 每次最多只能查詢 5 個關鍵字
    chunks = [column_data[i:i+5] for i in range(0, len(column_data), 5)]
    
    # 創建空的 DataFrame 用於存儲所有查詢結果
    all_data = pd.DataFrame()
    
    # 處理每一組關鍵字
    for chunk in chunks:
        try:
            # 設定 Google Trends 查詢參數
            # timeframe: 設定查詢時間範圍
            # geo='TW': 限定查詢台灣地區的資料
            # gprop='': 查詢一般網頁搜尋資料（不包含新聞、圖片等）
            pytrends.build_payload(
                kw_list=chunk, 
                timeframe='2020-01-01 2022-12-31',
                geo='TW',
                gprop=''
            )
            
            # 獲取搜尋趨勢資料
            interest_over_time_df = pytrends.interest_over_time()
            
            # 將新獲取的資料與現有資料合併
            # axis=1 表示水平方向合併（新增欄位）
            all_data = pd.concat([all_data, interest_over_time_df], axis=1)
            
            # 暫停 60 秒，避免觸發 Google Trends 的請求限制
            time.sleep(60)
        
        except Exception as e:
            print(f"發生錯誤: {e}")
    
    # 將處理完的資料儲存為 Excel 檔案
    # 檔名為欄位名稱
    output_file_path = os.path.join(output_folder, f"{column}.xlsx")
    # index=True 確保保留時間戳記資訊
    all_data.to_excel(output_file_path, index=True)
    
    print(f"已將 {column} 的趨勢資料儲存至 {output_file_path}")