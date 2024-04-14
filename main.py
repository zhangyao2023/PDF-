import tkinter as tk
from tkinter import filedialog
from typing import Union, Tuple
from reportlab.lib import units
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF4 import PdfFileWriter, PdfFileReader
from typing import List
from pikepdf import Pdf, Page, Rectangle
import PyPDF4
from pdf2docx import Converter
import datetime
pdfmetrics.registerFont(TTFont('msyh', r'./msyh.ttc'))

'''

用于生成包含content文字内容的水印pdf文件

content: 水印文本内容
filename: 导出的水印文件名
width: 画布宽度，单位：mm
height: 画布高度，单位：mm
font: 对应注册的字体代号
fontsize: 字号大小
angle: 旋转角度
text_stroke_color_rgb: 文字轮廓rgb色
text_fill_color_rgb: 文字填充rgb色
text_fill_alpha: 文字透明度


'''



# 创建主窗口
root = tk.Tk()
root.title("PDF使用小工具byZY")
root.geometry('425x150')
# 创建文本输入框
label = tk.Label(root, text="水印：文本框输入水印文字，点击开始添加水印按钮\n删除或保留页码：文本框输入需要删除或保留的页码，点击删除或保留页\n如1,3,5,7,9使用英文逗号分隔")
label.pack()
entry = tk.Entry(root)
entry.pack()



def get_input_text():
    entered_text = entry.get()
    # print("Entered text:", entered_text)
    return entered_text



def create_wartmark(content: str,
                    filename: str,
                    width: Union[int, float],
                    height: Union[int, float],
                    font: str,
                    fontsize: int,
                    angle: Union[int, float] = 45,
                    text_stroke_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                    text_fill_color_rgb: Tuple[int, int, int] = (0, 0, 0),
                    text_fill_alpha: Union[int, float] = 1) -> None:
    # 创建PDF文件，指定文件名及尺寸，以像素为单位
    c = canvas.Canvas(f'{filename}.pdf', pagesize=(width * units.mm, height * units.mm))

    # 画布平移保证文字完整性
    c.translate(0.1 * width * units.mm, 0.1 * height * units.mm)

    # 设置旋转角度
    c.rotate(angle)

    # 设置字体大小
    c.setFont(font, fontsize)

    # 设置字体轮廓彩色
    c.setStrokeColorRGB(*text_stroke_color_rgb)

    # 设置填充色
    c.setFillColorRGB(*text_fill_color_rgb)

    # 设置字体透明度
    c.setFillAlpha(text_fill_alpha)

    # 绘制字体内容
    c.drawString(0, 0, content)

    # 保存文件

    c.save()

'''
向目标pdf文件批量添加水印
target_pdf_path:目标pdf文件路径+文件名
watermark_pad_path:水印pdf文件路径+文件名
nrow:水印平铺的行数
ncol:水印平铺的列数
skip_pages:需要跳过不添加水印的页数

'''


def select_multiple_files():
    # 打开多文件选择对话框
    files = filedialog.askopenfilenames( initialdir = '/',filetypes=[( "PDF文件", ".pdf"), ('All Files', ' *')])
    # 如果用户选择了文件，则打印文件路径
    return files

# 创建一个按钮，点击后打开多选文件框
# open_button = tk.Button(root, text = '多选PDF文件', command = select_multiple_files)
# open_button.pack(expand = True, fill = 'both')



def add_watemark(target_pdf_path: str,
                 watermark_pdf_path: str,
                 nrow: int,
                 ncol: int,
                 skip_pages: List[int] = []) -> None:
    # 选择需要添加水印的pdf文件
    target_pdf = Pdf.open(target_pdf_path)

    # 读取水印pdf文件并提取水印
    watermark_pdf = Pdf.open(watermark_pdf_path)
    watermark_page = watermark_pdf.pages[0]

    # 遍历目标pdf文件中的所有页，批量添加水印
    for idx, target_page in enumerate(target_pdf.pages):
        for x in range(ncol):
            for y in range(nrow):
                # 向目标页指定范围添加水印
                target_page.add_overlay(watermark_page,
                                        Rectangle(target_page.trimbox[2] * x / ncol,
                                                  target_page.trimbox[3] * y / nrow,
                                                  target_page.trimbox[2] * (x + 1) / ncol,
                                                  target_page.trimbox[3] * (y + 1) / nrow
                                                  ))
    # 保存PDF文件，同时对pdf文件进行重命名，从文件名第7位置写入后缀名
    target_pdf.save(target_pdf_path[:-4] + '水印.pdf')





def dashuiyin():
    create_wartmark(content = get_input_text(),
                    filename = '水印',
                    width = 200,
                    height = 200,
                    font = 'msyh',
                    fontsize = 40,
                    text_fill_alpha = 0.3)# 透明度0.3
    paths=select_multiple_files()
    for path in paths:
        add_watemark(target_pdf_path = path,
                     # 把生成的水印示例，添加到目标水印文件中
                     watermark_pdf_path = '水印.pdf',
                     nrow = 5,
                     ncol = 3,
                     skip_pages = [0])

    close_window()
def close_window():
    root.destroy()


def remove_pages(input_file, output_file, pages_to_remove):
    # 使用PdfReader读取原始PDF
    reader = PdfFileReader(input_file)
    # zongyeshu=range(1,len(reader.pages)+1)
    # baoliu=list(set(zongyeshu)-set(shanchuyema))
    # print("删除页码："+str(shanchuyema)+"\n保留页码："+str(baoliu))
    # 使用PdfWriter创建一个新的PDF
    writer = PdfFileWriter()

    # 遍历PDF的每一页

    for page_number in range(len(reader.pages)):
        # 如果这个页面不在要删除的页面列表中
        if str(page_number+1) not in pages_to_remove:

            writer.addPage(reader.pages[page_number])


    # 写入新的PDF文件，覆盖原有文件
    with open(output_file, 'wb') as out:
        writer.write(out)


# 使用函数，删除第1页和第3页，保留其他页面
def shanchu():
    shanchuyema1 = get_input_text()
    shanchuyema=shanchuyema1.split(',')
    path1 = select_multiple_files()
    path=path1[0]

    wenjianname=path[:-4] + '已删除.pdf'
    remove_pages(path, wenjianname, shanchuyema)

    close_window()



def baoliu_pages(input_file, output_file, pages_to_baoliu):
    # 使用PdfReader读取原始PDF
    reader = PdfFileReader(input_file)
    zongyeshu=range(1,len(reader.pages)+1)
    shanchu1=list(set(zongyeshu)-set(pages_to_baoliu))
    # print("总页码"+str(zongyeshu)+"保留页码："+str(pages_to_baoliu)+"\n删除页码："+str(shanchu))
    # print(type(zongyeshu),type(pages_to_baoliu),type(shanchu))
    # 使用PdfWriter创建一个新的PDF
    writer = PdfFileWriter()
    shanchu= ['{}'.format(num) for num in shanchu1]
    # 遍历PDF的每一页
    for page_number in range(len(reader.pages)):
        # 如果这个页面不在要删除的页面列表中
        if str(page_number + 1) not in shanchu:
            # 将这个页面添加到新的PDF中
            writer.addPage(reader.pages[page_number])

    # 写入新的PDF文件，覆盖原有文件
    with open(output_file, 'wb') as out:
        writer.write(out)


# 使用函数，删除第1页和第3页，保留其他页面

def baoliu():
    baoliuyema1 = get_input_text()
    baoliuyema2=baoliuyema1.split(',')
    baoliuyema = [int(num) for num in baoliuyema2]
    paths = select_multiple_files()

    for path in paths:
        baoliuname = path[:-4] + '已保留.pdf'
        baoliu_pages(path, baoliuname, baoliuyema)
    close_window()


def merge_pdfs(input_paths, output_path):
    pdf_writer = PyPDF4.PdfFileWriter()
    with open(output_path, 'wb') as output_file:
        for path in input_paths:
            with open(path, 'rb') as input_file:
                pdf_reader = PyPDF4.PdfFileReader(input_file)
                for page_num in range(pdf_reader.getNumPages()):
                    pdf_writer.addPage(pdf_reader.getPage(page_num))

                pdf_writer.write(output_file)




def hebing():
    string = str(datetime.datetime.now()).split('.')[0].replace(':', '').replace("-", "").replace(" ", "")
    output_path = "合并"+string + '.pdf'  # 输出文件的地址
    input_paths = select_multiple_files()  # 输入你需要合并的所有pdf,按先有后顺序合并
    merge_pdfs(input_paths, output_path)  #调用合并
    close_window()

def convert_pdf_to_word(pdf_path, docx_path):
    # 创建一个Converter对象
    cv = Converter(pdf_path)
    # 转换PDF为Word文档
    cv.convert(docx_path, start = 0, end = None)
    # 关闭转换器
    cv.close()

def zhuanword():
    paths = select_multiple_files()

    for path in paths:
        wordname=path[:-4] + '已转换.docx'
        convert_pdf_to_word(path, wordname)

    close_window()



# 创建按钮，点击按钮时获取文本框中的内容
button1 = tk.Button(root, text="开始添加水印", command=dashuiyin)
button1.pack(side=tk.LEFT, padx=5, pady=5)

button2 = tk.Button(root, text="删除指定页", command=shanchu)
button2.pack(side=tk.LEFT, padx=5, pady=5)

button3 = tk.Button(root, text="保留指定页", command=baoliu)
button3.pack(side=tk.LEFT, padx=5, pady=5)

button4 = tk.Button(root, text="合并多个文件", command=hebing)
button4.pack(side=tk.LEFT, padx=5, pady=5)

button5 = tk.Button(root, text="转为WORD", command=zhuanword)
button5.pack(side=tk.LEFT, padx=5, pady=5)


# 运行主循环
root.mainloop()