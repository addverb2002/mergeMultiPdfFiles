#coding=utf-8  
"""========================
指定python的安装源进行模块安装:
pip install 模块 -i 安装源
国内的安装源有:
清华:https://pypi.tuna.tsinghua.edu.cn/simple
阿里云:https://mirrors.aliyun.com/pypi/simple/
中国科技大学 https://pypi.mirrors.ustc.edu.cn/simple/
华中理工大学:https://pypi.hustunique.com/
山东理工大学:https://pypi.sdutlinux.org/
豆瓣:https://pypi.douban.com/simple/
默认的安装源为:https://pypi.Python.org/simple/

需要安装的模块：pip install xlrd -i https://pypi.tuna.tsinghua.edu.cn/simple
pip install xlwt -i https://pypi.tuna.tsinghua.edu.cn/simple
created by:z語默 已于 2022-06-15 11:21:45 修改
 
"""
'''
利用Python将多个PDF文件的指定页进行合并，输出为一个文件
author: lc
date:20220710
status: ok
#按照下列步骤进行操作： 
#1. 把需要合并的pdf文档放置在一个目录下,注意文件顺序
#2.记录下文件的绝对路径
#3. 输入新的pdf文件的输出路径
#4. 运行程序得到新文件

'''

from datetime import datetime
 
from PyPDF2 import PdfFileMerger
import PyPDF2
import os

def split_pdf(source:str, pages:tuple=(), targetfile:str=""):
    """
    切割指定的PDF文件
    :param source: 被切割的文件路径和名称   
    :param pages: 若指定则只将指定的页切割出来, 否则将文件默认切割为单独的pdf文件.eg:(1,3,'5-7',9)
    :param target: 文件保存路径.不输入则返回需要合并的文件页类
    """
    filename = source.split('\\')[-1].rstrip('.pdf')
    pdf_input = PyPDF2.PdfFileReader(open(source, 'rb'))
    pdf_output = PyPDF2.PdfFileWriter()
    pstring=''
    # 获取 pdf 共有多少页
    page_count = pdf_input.getNumPages()
    print(f'分割前的PDF file is {source},页数: ', page_count)
    if page_count == 0:
        print('分割失败，原文档页数为0页!')
        return None,False
    for p in pages:
        pns=str(p).split('-')
        for pn in pns:
            pnumber=int(pn)-1 #pdf文件从第一页开始，内存中从第0页开始
            if(pnumber<page_count and pnumber>=0):
                pdf_output.addPage(pdf_input.getPage(pnumber))   
                pstring+='-'+str(pnumber)
    #for i in range(0, page_count):
    if(targetfile!=''):
        if(not targetfile.lower().endswith('.pdf')):
            targetfile+='.pdf'
        pdf_output.write(open(targetfile , 'wb'))
            
    return pdf_output,True
def mergePartialPdfs(pdfspages:tuple,outputpdfPath:str=''):
    '''合并指定路径下的所有pdf文件的指定页为一个新的pdf文件,新文件路径默认为pdfs文件路径，名称默认为new+日期+.pdf
    para pdfspages: (pdf文件夹路径和文件名,合并的指定页(1,3,5-9)),eg:(('d:\\123.pdf', (1,2, '4-5')),('d:\\b1.pdf',('1')))
    para outputpdfPath: 输出的pdf文件路径 
    '''
    if(pdfspages == None):
       return
    merger = PdfFileMerger()
    if(outputpdfPath ==''):         
        outputpdfPath='d:\\' 
    else:
        if(not outputpdfPath.endswith('\\')):
            outputpdfPath+='\\'
    dt='%Y%m%d%H%M%S%f'
    newpdf=outputpdfPath+f'new{datetime.now().strftime(dt)}.pdf'
    #处理源文件信息
    files=[]
    pages=[]
    for ele in pdfspages:#解析为文件和页码的列表
        files.append(ele[0])
        pages.append((ele[1]))
    fps=PyPDF2.parse_filename_page_ranges(list(pdfspages)) 
    print('开始处理pdf文档')
    pdf_output = PyPDF2.PdfFileWriter()
    for i in range(0,len(files),1):
        pdf_out, result=split_pdf(files[i], pages[i])
        if(result):
            for pg in range(pdf_out.getNumPages()):
                pdf_output.addPage(pdf_out.getPage(pg)) #添加指定页码
        else:
            pass
    #merger.pages.append()
    if(pdf_output.getNumPages()>0):
        pdf_output.write(open(newpdf, 'wb'))
        print(f'输出pdfw稳定{newpdf}成功')
    return True
def mergeFullPdfs(pdffolder:str,outputpdfPath:str='',isorderbydescending=False):
    '''合并指定路径下的所有pdf为一个新的pdf文件,新文件路径默认为pdfs文件路径，名称默认为new+日期+.pdf
    para pdffolder: pdf文件夹路径
    para outputpdfPath: 输出的pdf文件路径
    para isorderbyascending: false=升序排序 true=降序排序
    '''
    if(not pdffolder.endswith('\\')):
        pdffolder+='\\'
    merger = PdfFileMerger()
    if(outputpdfPath ==''):         
        outputpdfPath=pdffolder 
    else:
        if(not outputpdfPath.endswith('\\')):
            outputpdfPath+='\\'
    dt='%Y%m%d%H%M%S%f'
    newpdf=outputpdfPath+f'new{datetime.now().strftime(dt)}.pdf'
    files = os.listdir(pdffolder)#列出目录中的所有文件
    files.sort(reverse= isorderbydescending)
    print('开始处理pdf文档')
    for file in files: #从所有文件中选出pdf文件合并
        if file[-4:].lower() != ".pdf":
            continue
        merger.append(pdffolder+file)#open(file, 'rb')
    #merger.pages.append()
    with open(newpdf, 'wb') as fout:  #输出文件为newfile.pdf
        merger.write(fout)
        print(f'输出pdfw稳定{newpdf}成功')
if __name__ == "__main__":
    #test:
    #mergeFullPdfs('d:\\',isorderbydescending=True) #ok
    pps=(('d:\\123.pdf', (2,1, '4-5')),
         ('d:\\b1.pdf',('1'))
        )
    mergePartialPdfs(pps)