class Ebooktools():
    """
    電子書截圖及轉檔工具
    """
    def __init__(self):
        pass


    def sceenshottojpg(self,outputpath,pagecount):
        self.outputpath = outputpath
        self.pagecount = pagecount
        import time
        import pyautogui

        # 设置截图参数
        screenshot_directory = self.outputpath #"C:\Python 邁向數據科學之路 ChatGPT助攻/"  # 存储截图的目录
        screenshot_count = self.pagecount #985  # 截图数量
        next_page_button_position = (1877, 512)  # 下一页按钮的坐标位置

        """
        第二頁 左上 x=959  y=28
        右上 x=1700 y=29
        左下x =959 y=1028
        右下 x=1700 y=1028

        第一頁 左上 x=217 y=28
        右上 x=959  y=28
        左下 x=217 y=1028
        右下 X=959 Y=1028
        """


        #pyautogui.position(): 获取当前鼠标指针的位置。

        # 获取指定画面的位置
        #第一頁
        left1= 217
        top1= 28
        width1= 742 #1645
        height1= 1000 #1019
        #第二頁
        left2= 959
        top2= 28
        width2= 742 #1645
        height2= 1000 #1019

        # 设置截图区域
        region1 = (left1, top1, width1, height1)
        region2 = (left2, top2, width2, height2)

        # 等待一些时间，以便你能够切换到需要截图的窗口
        time.sleep(1)

        # 循环进行截图
        for i in range(screenshot_count):
            page=1 if i%2==0 else 2
            # 根据当前的计数生成文件名
            filename = f"{str(i+1).zfill(3)}.jpg"  # 如001.jpg
            filepath = screenshot_directory + filename

            # 进行屏幕截图
            region = region1 if i%2==0 else region2
            screenshot = pyautogui.screenshot(region=region)

            # 保存截图
            screenshot.save(filepath)

            # 点击下一页按钮
            if i%2==1:
                pyautogui.click(next_page_button_position)

            # 等待五秒钟
            time.sleep(3)

        def jpgtopdf(self,inputpath,outputpath):
            self.inputpath = inputpath
            self.outputpath = outputpath
            from PIL import Image
            from fpdf import FPDF
            import glob

            # 图像文件夹路径和输出PDF文件路径
            image_folder = "C:/1234/"
            pdf_filename = "output.pdf"

            # 创建PDF对象
            pdf = FPDF(orientation="L", unit="pt")

            # 获取图像文件列表
            image_files = glob.glob(image_folder + "*.jpg")

            # 遍历图像文件列表并将图像添加到PDF中
            for image_file in image_files:
                # 打开图像文件
                image = Image.open(image_file)
                
                # 调整图像大小以适应PDF页面
                max_width = 842  # PDF页面的宽度
                max_height = 595  # PDF页面的高度
                image.thumbnail((max_width, max_height))
                
                # 获取调整后的图像尺寸
                width, height = image.size
                
                # 添加PDF页面并将图像添加到页面上
                pdf.add_page()
                pdf.image(image_file, x=0, y=0, w=width, h=height)

            # 保存PDF文件
            pdf.output(pdf_filename, "F")