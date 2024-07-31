# PDF 文件拆除

import pikepdf
import os

# 定义输入文件路径和输出目录
input_file = 'D:/02-導線失竊/008導線失竊會議資料/112年//112年10月電纜線失竊會議(配技會)/112年10月彰化區處防制電纜線失竊專案會議資料.pdf'
output_directory = 'D:/02-導線失竊/008導線失竊會議資料/112年/112年10月電纜線失竊會議(配技會)/output'

# 创建输出目录
os.makedirs(output_directory, exist_ok=True)

# 打开输入文件
pdf = pikepdf.open(input_file)

# 拆分每一页并保存到输出目录
for i, page in enumerate(pdf.pages, start=1):
    output_path = os.path.join(output_directory, f'page_{i}.pdf')
    new_pdf = pikepdf.new()
    new_pdf.pages.append(page)
    new_pdf.save(output_path)

# 关闭PDF文件
pdf.close()