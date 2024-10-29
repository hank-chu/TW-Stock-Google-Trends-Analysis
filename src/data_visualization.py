import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os

# 設定中文字型
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei']
plt.rcParams['axes.unicode_minus'] = False

# 建立輸出資料夾
output_folder = "trend_plots"
os.makedirs(output_folder, exist_ok=True)

def plot_trend_line(excel_path):
    """繪製搜尋趨勢折線圖"""
    # 讀取資料
    df = pd.read_excel(excel_path, index_col=0)
    
    # 設定圖表大小
    plt.figure(figsize=(15, 8))
    
    # 繪製每個關鍵字的趨勢線
    for column in df.columns:
        if column != 'isPartial':
            plt.plot(df.index, df[column], label=column, linewidth=2)
    
    # 設定圖表標題和標籤
    plt.title('Google Trends 搜尋趨勢', fontsize=16)
    plt.xlabel('日期', fontsize=12)
    plt.ylabel('搜尋熱度', fontsize=12)
    plt.legend()
    plt.grid(True)
    
    # 儲存圖表
    file_name = os.path.basename(excel_path).replace('.xlsx', '')
    plt.savefig(f'{output_folder}/{file_name}_trend_line.png')
    plt.close()

def plot_heatmap(excel_path):
    """繪製搜尋熱度熱力圖"""
    # 讀取資料
    df = pd.read_excel(excel_path, index_col=0)
    
    # 計算月均值
    monthly_data = df.groupby(pd.Grouper(freq='ME')).mean()
    monthly_data.index = monthly_data.index.strftime('%Y-%m')
    
    # 繪製熱力圖
    plt.figure(figsize=(15, 10))
    sns.heatmap(monthly_data.T, cmap='YlOrRd', 
                xticklabels=12, 
                cbar_kws={'label': '搜尋熱度'})
    
    # 設定圖表標題和標籤
    plt.title('搜尋熱度月度分布', fontsize=16)
    plt.xlabel('時間', fontsize=12)
    plt.ylabel('關鍵字', fontsize=12)
    plt.xticks(rotation=45)
    
    # 儲存圖表
    file_name = os.path.basename(excel_path).replace('.xlsx', '')
    plt.savefig(f'{output_folder}/{file_name}_heatmap.png')
    plt.close()

def plot_correlation(excel_path):
    """繪製關鍵字相關性圖"""
    # 讀取資料
    df = pd.read_excel(excel_path, index_col=0)
    
    # 計算相關係數
    correlation = df.corr()
    
    # 繪製相關性熱力圖
    plt.figure(figsize=(10, 8))
    sns.heatmap(correlation, 
                annot=True, 
                cmap='coolwarm', 
                center=0,
                fmt='.2f')
    
    # 設定圖表標題
    plt.title('關鍵字相關性分析', fontsize=16)
    
    # 儲存圖表
    file_name = os.path.basename(excel_path).replace('.xlsx', '')
    plt.savefig(f'{output_folder}/{file_name}_correlation.png')
    plt.close()

def plot_seasonal(excel_path):
    """繪製季節性分析圖"""
    # 讀取資料
    df = pd.read_excel(excel_path, index_col=0)
    
    # 添加月份資訊
    df['month'] = df.index.month
    
    # 設定圖表
    plt.figure(figsize=(15, 8))
    
    # 繪製每個關鍵字的月度平均
    for column in df.columns:
        if column not in ['isPartial', 'month']:
            monthly_avg = df.groupby('month')[column].mean()
            plt.plot(monthly_avg.index, monthly_avg.values, 
                    label=column, marker='o')
    
    # 設定圖表標題和標籤
    plt.title('搜尋趨勢季節性分析', fontsize=16)
    plt.xlabel('月份', fontsize=12)
    plt.ylabel('平均搜尋熱度', fontsize=12)
    plt.legend()
    plt.grid(True)
    plt.xticks(range(1, 13))
    
    # 儲存圖表
    file_name = os.path.basename(excel_path).replace('.xlsx', '')
    plt.savefig(f'{output_folder}/{file_name}_seasonal.png')
    plt.close()

def process_all_files(data_folder="Google_Trends_Output"):
    """處理資料夾中的所有檔案"""
    # 確保資料夾存在
    if not os.path.exists(data_folder):
        print(f"找不到資料夾: {data_folder}")
        return
    
    # 處理每個 Excel 檔案
    for filename in os.listdir(data_folder):
        if filename.endswith('.xlsx'):
            file_path = os.path.join(data_folder, filename)
            print(f"正在處理: {filename}")
            
            try:
                # 繪製四種圖表
                plot_trend_line(file_path)
                plot_heatmap(file_path)
                plot_correlation(file_path)
                plot_seasonal(file_path)
                print(f"完成 {filename} 的圖表生成")
                
            except Exception as e:
                print(f"處理 {filename} 時發生錯誤: {e}")

# 主程式
if __name__ == "__main__":
    process_all_files()
    print("所有圖表已生成完成！")