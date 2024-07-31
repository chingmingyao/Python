import pdfplumber

# 读取pdf文件，保存为pdf实例
pdf =  pdfplumber.open("c:\\1111.pdf") 

# 访问第二页
#first_page = pdf.pages[0]

# 自动读取表格信息，返回列表
table = pdf.extract_table()

table