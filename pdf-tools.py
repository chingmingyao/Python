



def pdf_to_excel(path,pdf_name,excel_output_name):
    import pdfplumber
    import openpyxl
    """
    將pdf檔轉換為文字檔並存到excel
    """


    # PDF 報表路徑
    # pdf_name = "B1Report08_2024-06-05-17-25-39.pdf"
    # pdf_path = "D:\\00 - 線損資料\\00-線損工作事項\\002_3_分析再生能源購電量\\000_再生能源購電系統抓資料(每月約5日才會出來)\\113\\5月\\" + pdf_name
    
    pdf_name = pdf_name
    pdf_path = path + pdf_name


    

    # 開啟 Excel 檔案
    wb = openpyxl.Workbook()
    sheet = wb.active

    # 開啟 PDF 報表  
    with pdfplumber.open(pdf_path) as pdf:
        
        # 迭代每頁
        for page in pdf.pages: 
            
            # 提取文字 
            text = page.extract_text()  
            
            # 提取表格 (如果有的話)
            table = page.extract_table()  
            
            # 寫入頁碼
            sheet.append(['Page ' + str(page.page_number)])  
            sheet.append([])

            # 寫入文字 
            sheet.append(['Text'])
            sheet.append([text])  
            sheet.append([])  
            
            # 寫入表格
            if table:
                headers = [str(cell) for cell in table[0]]
                sheet.append(headers)
                
                for row in table[1:]:
                    data = [str(cell) for cell in row]
                    sheet.append(data)  
                    
                sheet.append([])
                
    # 儲存 Excel 檔案        
     
    excel_path = r"D:\00 - 線損資料\00-線損工作事項\002_3_分析再生能源購電量\000_再生能源購電系統抓資料(每月約5日才會出來)\113\5月\113_05_data.xlsx"    
    wb.save(excel_path) 

    # print('PDF 轉換 Excel 完成!')


#產出報表處理

def excel_process(file_path:str,file_name:str)-> None:
    
    import xlwings as xw
    wb = xw.Book(file_path+file_name)
    sheet = wb.sheets['Sheet']
    wb.app.screen_updating = False
    for row in range(7983, 18, -1): 
        # 從7127開始,向上檢查到18停止將多除文字刪除
        cell_value = sheet.range('A' + str(row)).value
        if not cell_value or "Page" in cell_value or "Text" in cell_value or "113年" in cell_value or "再生能源" in cell_value or "None" in cell_value:
            sheet.api.Rows(row).Delete()

    wb.save()
    wb.app.screen_updating = True



   

def show_message(mytitle,mymessage):
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()  # 隱藏空白視窗
    root.attributes("-topmost", True)  # 設置視窗在最上層
    # 顯示訊息框
    messagebox.showinfo(mytitle, mymessage,master=root)
    root.destroy()  # 關閉視窗
# 調用 show_message 函數


if __name__ == '__main__':
    import time
    start_time = time.time()
    # PDF 報表路徑
    #以下資料需修改
    pdf_name = "B1Report08_2024-06-05-17-25-39.pdf" 
    
    file_path = "D:\\00 - 線損資料\\00-線損工作事項\\002_3_分析再生能源購電量\\000_再生能源購電系統抓資料(每月約5日才會出來)\\113\\5月\\" 

    excel_output_name = "113_05_data.xlsx"

    #執行PDF 轉 excel
    pdf_to_excel(file_path, pdf_name, excel_output_name)

    #執行Excel 後續處理 去除不要資料及空白行
    excel_process(file_path, excel_output_name)

    #執行Excel 後續處理 去除不要資料及空白行    
   
        
    end_time = time.time()

    total_time = end_time - start_time
    show_message("處理狀態",f"Pdf to Excel 轉檔完成 總花費時間 : {total_time}")

    print("處理完成")
