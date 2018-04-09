#!/usr/bin/python
# -*- coding: utf-8 -*-

from PyPDF2 import PdfFileReader, PdfFileWriter
from pdfminer.pdfparser import PDFParser,PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import *
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
import warnings
import sys
import importlib
importlib.reload(sys)
__author__ = 'mei'
warnings.filterwarnings('ignore')


# 解析pdf文件函数
def parse(pdf_path):
    fp = open(pdf_path, 'rb')  # 以二进制读模式打开
    # 用文件对象来创建一个pdf文档分析器
    parser = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument()
    # 连接分析器 与文档对象
    parser.set_document(doc)
    doc.set_parser(parser)
    doc.initialize()
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        rsrcmgr = PDFResourceManager()
        laparams = LAParams()
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        # 用来计数页面，图片，曲线，figure，水平文本框等对象的数量
        num_page = 0
        text0_now, text0_last, text1_now, text1_last = "", "", "", ""
        page_list = []
        for page in doc.get_pages():  # doc.get_pages() 获取page列表
            num_page += 1  # 页面增一
            interpreter.process_page(page)
            layout = device.get_result()
            for x in layout:
                if isinstance(x, LTTextBoxHorizontal):  # 获取文本内容
                    if num_page == 1:
                        if x.index == 0:
                            text0_now = x.get_text()
                        if x.index == 1:
                            text1_now = x.get_text()
                    else:
                        if x.index == 0:
                            text0_last = text0_now
                            text0_now = x.get_text()
                        if x.index == 1:
                            text1_last = text1_now
                            text1_now = x.get_text()
            if num_page != 1:
                if text1_now != text1_last:
                    page_list.append(num_page-1)  # last page
                if text1_now == text0_last:
                    page_list.append(num_page)  # now page
        return page_list


def info_page(readFile):
    pdfFileReader = PdfFileReader(readFile)  # 或者这个方式：pdfFileReader = PdfFileReader(open(readFile, 'rb'))
    # 获取 PDF 文件的文档信息
    documentInfo = pdfFileReader.getDocumentInfo()
    print('documentInfo = %s' % documentInfo)
    # 获取页面布局
    pageLayout = pdfFileReader.getPageLayout()
    print('pageLayout = %s ' % pageLayout)

    # 获取页模式
    pageMode = pdfFileReader.getPageMode()
    print('pageMode = %s' % pageMode)

    xmpMetadata = pdfFileReader.getXmpMetadata()
    print('xmpMetadata  = %s ' % xmpMetadata)

    # 获取 pdf 文件页数
    pageCount = pdfFileReader.getNumPages()

    print('pageCount = %s' % pageCount)
    for index in range(0, pageCount):
        # 返回指定页编号的 pageObject
        pageObj = pdfFileReader.getPage(index)
        print('index = %d , pageObj = %s' % (index, pageObj))  # <class 'PyPDF2.pdf.PageObject'>
        # 获取 pageObject 在 PDF 文档中处于的页码
        pageNumber = pdfFileReader.getPageNumber(pageObj)
        print('pageNumber = %s ' % pageNumber)


def add_page_pdf(infn_, list_):
    output_ = PdfFileWriter()
    input_ = PdfFileReader(open(infn_, 'rb'))
    for i in list_:
        output_.addPage(input_.getPage(i-1))
    return output_


if __name__ == '__main__':
    infn = r'C:\Users\lenovo\Desktop\田老师现代控制ppt\mcsChapt132013.pdf'
    outfn = r'C:\Users\lenovo\Desktop\田老师现代控制ppt\mcsChapt132013_01.pdf'

    total = parse(infn)
    print(total)

    pdf_output = add_page_pdf(infn, total)
    try:
        pdf_output.write(open(outfn, 'wb'))
    except Exception as e:
        print(e)
    print("output success!")
