import pandas as pd

def main():  # 自訂主程式 (Define the main program)
    """
    This function processes the '上市公司資料.csv' file, cleans the data, 
    renames columns, and saves the processed data to both CSV and Excel formats.
    """
    df = pd.read_csv('上市公司資料.csv')  # 讀取上市公司資料.csv 檔案 (Read the '上市公司資料.csv' file)
    df1 = df.dropna()  # 移除包含 NaN 值的列 (Remove rows containing NaN values)
    df2 = df1.reindex(columns=['公司代號','出表日期','公司名稱','產業別','營業收入-當月營收','營業收入-上月營收'])  # 重新索引，僅保留指定欄位 (Reindex the DataFrame to keep only specified columns)
    df3 = df2.rename(columns={  # 重新命名欄位名稱 (Rename the column names)
            '營業收入-上月營收':'上月營收',  # 將 '營業收入-上月營收' 改名為 '上月營收' (Rename '營業收入-上月營收' to '上月營收')
            '營業收入-當月營收':'當月營收'  # 將 '營業收入-當月營收' 改名為 '當月營收' (Rename '營業收入-當月營收' to '當月營收')
            })
    df3.to_csv('上市公司資料整理.csv',encoding='utf-8-sig')  # 將處理後的資料存成 CSV 檔案，並使用 utf-8-sig 編碼 (Save the processed data to a CSV file with utf-8-sig encoding)
    df3.to_excel('上市公司資料整理.xlsx')  # 將處理後的資料存成 Excel 檔案 (Save the processed data to an Excel file)
    print("存檔完成")  # 印出「存檔完成」訊息 (Print "存檔完成" - "Saving completed" message)

if __name__ == '__main__':  # 判斷是否為程式的入口 (Check if the script is the main entry point)
    main()  # 呼叫主程式 (Call the main function)
