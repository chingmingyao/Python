def user_update(mybook="D:/每日用戶數.xlsm"):
    """
    mybook = 檔案名稱及路徑 "c:/123/123.xlsx"
    
    存放位置
    工作表名稱   
        
    """


    #用戶數更新    
    import pandas as pd
    import requests 
    import xlwings as xw
    import datetime
    import logging
    from io import StringIO
     
    # 設置日誌檔案名稱和級別
   

    app = xw.App(visible=False,add_book=False)
    wb = app.books.open(mybook)
    today =datetime.datetime.today()
    mytoday =today.strftime("%Y%m%d%H%M")
    if mytoday in wb.sheet_names:
        print("已存在")
        wb.close()
        app.quit()
    else:


        
        # 您的程式碼
    

        #用戶更新網址
        url = "http://10.210.35.202/OmsApp201/Users/AllFeederUser.aspx"

        headers = {
        "Accept":"text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding":"gzip, deflate",
        "Accept-Language":"zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7",
        "Cache-Control":"max-age=0",
        "Connection":"keep-alive",
        "Content-Length":"287",
        "Content-Type":"application/x-www-form-urlencoded",
        "Cookie":"ASPSESSIONIDSSRDQRDC=AIFKCNCBNLGLKDACJFKJJNOE; ASPSESSIONIDSSQASRDC=LNAEBJPBLJJKDLOFFFDKBLAL; ASP.NET_SessionId=wvo53nrxpobvjtib1pioyobf",
        "Host":"10.210.35.202",
        "Origin":"http://10.210.35.202",
        "Referer":"http://10.210.35.202/OmsApp201/Users/AllFeederUser.aspx",
        "Upgrade-Insecure-Requests":"1",
        "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36"

        }
        payload = ("__VIEWSTATE=%2FwEPDwUJNTY1MTgzNjU5D2QWAgIBD2QWBAIHDzwrAAsAZAIJDw9kFgIeB29uY2xpY2sFD19TaG93TG9hZGluZygpO2Rkg62rXE9%2Br8sYmI7whrKeX3rknxs%3D&__VIEWSTATEGENERATOR=21A69B0D&__EVENTVALIDATION=%2FwEWAwKU%2F4a1AwL8o76uBAKyuOLUC3vnzAZb8Q%2Bf3Vn5rc%2BGcuplCmFq&SubmitButton=%C1%F4%C2%C3%AB%F6%B6s")

        res = requests.post(url,headers=headers,data=payload) 
        df = pd.read_html(StringIO(res.text),header=0)


        
        formatted_today = today.strftime("%Y-%m-%d")

        
        sht = wb.sheets["用戶數"]
        sht.range("A1").expand().clear_contents()
        sht["A1"].expand().options(index=None,header=True).value = df[1]
        sht["F1"].value = "饋線更新日期:"
        sht["G1"].value = formatted_today
        #複製用戶數工作表
        sht.copy()
        sht2 = wb.sheets["用戶數 (2)"]
        sht2.name = mytoday
        wb.save()
        wb.close()
        app.quit()
            
      
           
           
    
          
if __name__ == "__main__":
    user_update()
   
    # import sys
    # print(sys.argv[0:2])

    # # functions = {'user_update': user_update, 'user_update2': user_update2} 
    # if len(sys.argv) < 2:
    #     print("請提供要執行的函數名稱")
    #     sys.exit(1)    
    # function_name = sys.argv[1]   
    # if function_name in functions:
    #         # 呼叫相應的函數
    #     functions[function_name](*sys.argv[2:])
    # else:
    #     print("未知的函數:", function_name)
# 

   
