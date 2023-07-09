"""
安装 ttkbootstrap 模块
pip install ttkbootstrap
安装 windnd 模块
pip install windnd
安装 shutil 模块
pip install shutilwhich
"""

import ttkbootstrap as ttk
from PIL import ImageFilter, Image
from ttkbootstrap.constants import *
import windnd
from threading import Thread
import os
from os import remove, listdir, mkdir
from os.path import join, isdir, split, splitext, basename
import shutil
import pythoncom
from reportlab.lib.pagesizes import A4, portrait
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfMerger
from pdf2image import convert_from_path
from win32com.client import constants, gencache

doc_files = set()
error_files = []

# 实例化创建应用程序窗口
Converter = ttk.Window(
    title="Docx2PDF",  # 设置窗口的标题
    themename="solar",  # 设置主题
    size=(612, 288),  # 窗口的大小
    position=(612, 288),  # 窗口所在的位置
    minsize=(612, 288),  # 窗口的最小宽高
    maxsize=(612, 288),  # 窗口的最大宽高
    resizable=None,  # 设置窗口是否可以更改大小
    alpha=0.9,  # 设置窗口的透明度(0.0完全透明）
)
Converter.iconbitmap(f"convert.ico")

# 标签
ttk.Label(Converter, text="作者：LZY\t邮箱：1343161318@qq.com", bootstyle=SECONDARY).place(x=14, y=260)

# 文件夹数量展示
var_Display_volume = ttk.StringVar()
ttk.Label(Converter, textvariable=var_Display_volume).place(x=14, y=180)
var_Display_volume.set('文件数量：0\t转换成功：0\t')
# 线程检测
var_Line = ttk.StringVar()
ttk.Label(Converter, textvariable=var_Line).place(x=534, y=260)
var_Line.set('等待执行...')

# 窗口展示内容
text = ttk.Text(Converter, width=84, heigh=6)
text.pack(padx=0, pady=50)
text.insert('insert', '请将待转换的文档拖进此窗口内\n')  # 插入内容
text['state'] = 'disabled'  # 【禁用】 disabled   可用 normal


def Clean():
    if var_Line.get() == '执行中...':
        text.insert(END, '当前正在转换，请勿进行其他操作！\n')
    else:
        global doc_files
        doc_files.clear()
        text['state'] = 'normal'
        text.delete("0.0", 'end')
        text.insert('insert', '请将待转换的文档拖进此窗口内\n')
        text['state'] = 'disabled'
        var_Display_volume.set('文件数量：0\t转换成功：0\t')


ttk.Button(text="清除", width=10, command=Clean, bootstyle="primary").place(x=510, y=170)


# 拖拽触发DEF
def Dragoon(paths):
    global doc_files

    if doc_files:
        text['state'] = 'normal'  # 【可用】 禁用 disabled   可用 normal
    else:
        text['state'] = 'normal'
        text.delete('0.0', 'end')

    # 将拖拽的文件循环便利出来并解码 存储到列表内
    for path in paths:
        path = path.decode('gbk')  # 解码
        if os.path.isfile(path):  # 检测是否是文件 是文件True 否则False
            if path.endswith(('.doc', '.docx')):
                doc_files.add(os.path.abspath(path))
        elif os.path.isdir(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith(('.doc', '.docx')):
                        doc_files.add(os.path.abspath(os.path.join(root, file)))

    for doc in doc_files:
        if doc not in text.get('1.0', END):
            text.insert('insert', doc + '\n')

    var_Display_volume.set(f'文件数量：{len(doc_files)}\t转换成功：0\t')
    text['state'] = 'disabled'


def handle_and_export(file):
    global error_files
    if os.path.isfile(file):
        if get_doc(file):
            text['state'] = 'normal'
            text.insert('insert', '·处理完成 >>' + file + '\n')
            text.see(ttk.END)
            text['state'] = 'disabled'
        else:
            text.insert('insert', '未安装word或wps\n')
    else:
        error_files = error_files + ['【无法打开该文档】 >>' + file]


# def执行  -  主
def main():
    def T():
        if doc_files:
            text['state'] = 'normal'  # 【可用】 禁用 disabled   可用 normal
            text.delete("0.0", 'end')  # 删除内容
            text.insert('1.0', '正在转换\n')
            # text['state'] = 'disabled'  # 【禁用】 disabled   可用 normal
            var_Line.set('执行中...')

            for doc in doc_files:
                handle_and_export(doc)
            for file in error_files:
                text['state'] = 'normal'  # 【可用】 禁用 disabled   可用 normal
                text.insert('insert', f'{file}\n')  # 插入内容
                text.see(ttk.END)  # 光标跟随着插入的内容移动
                text['state'] = 'disabled'  # 【禁用】 disabled   可用 normal
            var_Display_volume.set(
                f'文件数量：{len(doc_files)}\t执行成功：{len(doc_files) - len(error_files)}\t异常文件：{len(error_files)}')
            var_Line.set('等待执行...')
            doc_files.clear()
            error_files.clear()

    if var_Line.get() == '执行中...':
        text.insert(END, '当前正在转换，请勿进行其他操作！\n')
    else:
        Thread(target=T).start()


# 按钮
ttk.Button(text="开始转换", width=20, command=main, bootstyle='primary').pack()

# 拖拽模块
windnd.hook_dropfiles(Converter, func=Dragoon)


# 把word文档转为pdf，适用于doc和docx
def doc2pdf(word_file, pdf_file):
    pythoncom.CoInitialize()
    try:
        word = gencache.EnsureDispatch('Kwps.Application')
    except:
        word = gencache.EnsureDispatch('word.Application')
    print(word)
    word.Visible = False
    document = word.Documents.Open(word_file)
    document.ExportAsFixedFormat(pdf_file,
                                 constants.wdExportFormatPDF,
                                 Item=constants.wdExportDocumentWithMarkup,
                                 CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
    document.Close()
    word.Quit(constants.wdDoNotSaveChanges)


# 把pdf文件拆分成jpg图片，每页一张
def pdf2jpgs(pdf_path):
    dst_dir, pdf_fn = split(pdf_path)
    if not dst_dir:
        dst_dir = pdf_fn[:-4] + '-tmp'
    else:
        dst_dir = join(dst_dir, pdf_fn[:-4]) + '-tmp'
    if not isdir(dst_dir):
        mkdir(dst_dir)
    base_path = os.getcwd()
    # base_path = os.path.dirname(os.path.abspath(__file__))
    # poppler_path = os.path.join(base_path, '../poppler-0.68.0/bin')
    poppler_path = os.path.join(base_path, 'poppler-0.68.0/bin')
    images = convert_from_path(pdf_path,
                               dpi=350,
                               fmt='JPEG',
                               thread_count=4,
                               poppler_path=poppler_path)
    for index, image in enumerate(images):
        image_path = '{}\{}.jpg'.format(dst_dir, index)
        image.save(image_path)
        smoothed_image = Image.open(image_path)
        smoothed_image = smoothed_image.filter(ImageFilter.SMOOTH_MORE)
        smoothed_image.save(image_path, "JPEG")


def merge_jpg2pdf(path):
    # 以防同名文件夹被误删
    jpg_path = path + '-tmp'
    jpg_files = [join(jpg_path, fn) for fn in listdir(jpg_path) if fn.endswith('.jpg')]
    jpg_files.sort(key=lambda fn: int(splitext(basename(fn))[0]))
    result_pdf = PdfMerger()
    temp_pdf = 'temp.pdf'

    for fn in jpg_files:
        c = canvas.Canvas(temp_pdf, pagesize=portrait(A4))
        c.drawImage(fn, 0, 0, *portrait(A4))
        c.save()

        with open(temp_pdf, 'rb') as fp:
            pdf_reader = PdfReader(fp)
            result_pdf.append(pdf_reader)
    result_pdf.write(path + '.pdf')
    result_pdf.close()
    remove(temp_pdf)
    # 以防同名文件夹被误删
    for pic in jpg_files:
        remove(pic)
    if not len(os.listdir(jpg_path)):
        shutil.rmtree(jpg_path)


def check_and_del(pdf_file):
    if os.path.exists(pdf_file) and os.path.isfile(pdf_file):
        os.remove(pdf_file)


def get_doc(path):
    pdf_file = os.path.splitext(path)[0] + ".pdf"
    check_and_del(pdf_file)
    try:
        doc2pdf(path, pdf_file)
    except:
        return False
    pdf2jpgs(pdf_file)
    merge_jpg2pdf(splitext(pdf_file)[0])
    return True


if __name__ == '__main__':
    Converter.mainloop()
