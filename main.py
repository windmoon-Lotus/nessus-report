

# -*-coding: utf-8 -*-

import tkinter as tk
import tkinter.filedialog as tkf
import os
import time
import googletranslater
import vuln_nessus
import win32com.client as win32
from docx import Document
from docxcompose.composer import Composer 
from mailmerge import MailMerge

def first():
    #参考了以下大佬的程序或博客,表示感谢。
    #vuln_nessus HTML部分：https://github.com/Bypass007/Nessus_to_report
    #vuln_venusHost.py  https://blog.csdn.net/u010984277/article/details/53695356
    #googletranslater.py 这个也找不到了，参考网上不知名大佬的

    pass


def selectFile1():
    #清除text1和result_text的文本内容
    text1.delete('1.0', 'end')
    result_text.delete('1.0', 'end')
    #显示文件名
    text1.insert('1.0', getFileName())


def selectFile2():
    #清除text2和result_text的文本内容
    text2.delete('1.0', 'end')
    result_text.delete('1.0', 'end')
    #显示文件名
    text2.insert('1.0', getFileName())


def selectFile3():
    #清除text3和result_text的文本内容
    text3.delete('1.0', 'end')
    result_text.delete('1.0', 'end')
    #显示文件名
    text3.insert('1.0', getFileName())


def selectFile4():
    #清除text4和result_text的文本内容
    text4.delete('1.0', 'end')
    result_text.delete('1.0', 'end')
    #显示文件名
    text4.insert('1.0', getFileName())

def getFileName():
    try:
        #因为按钮链接的函数不能有参数，所以需要全局变量file来接收打开的文件名
        global file
        file = str(tkf.askopenfilename())
        if file != '':
            return file
        else:
            return '无此文件。'
    except Exception as e:
        print("读取文件失败")
        print(e)

def getDestFileName(file1):
    #得到文件名和后缀
    (filename,extensionName) = os.path.splitext(file1)

    #生成目标文件名
    destFileName = filename + '漏洞整理.docx'
    return destFileName

#Nessus 漏洞整理函数
def zhengli_nessus():
    start_time = time.time()
    #提示信息
    result_text.insert('end',"开始整理中…… \n")
    result_text.update()

    (destF, extenN) = os.path.splitext(file)
    destFileName = destF + '漏洞整理.csv'

    if extenN == '.csv':

        #nessus 整理开始
        nessus_list = vuln_nessus.zhengli_csv(file, info_flag)
    elif extenN == '.html':
        #nessus 整理开始
         nessus_list = vuln_nessus.zhengli_html(file, info_flag)
    else:
        result_text.insert('end',"格式错误!\n")
        result_text.update()
        

    #提示信息
    result_text.insert('end',"写入报告中…… \n")
    result_text.update()
    #nessus开始写入
    vuln_nessus.write2csv(nessus_list, destFileName)
    #提示信息
    result_text.insert('end',"写入csv成功，请到原漏洞文件目录下查看。 \n")
    result_text.update()
    consume_time = time.time() - start_time
    #提示信息
    result_text.insert('end',"耗时时间为："+str(consume_time) + '\n')
    result_text.update()





def write2docx(name, vulns_list, destFile):
    #得到当前main.py文件的路径，以便放 生成的模板文件
    pwd = os.getcwd()

    #漏洞模板文件，最上面的换行不要删除
    template_file = pwd + '\\vuln_template.docx'
    name_list = []
    count = 0
    #循环生成每个漏洞文件
    for vuln in vulns_list:
        count += 1
        #读取模板文件
        template = MailMerge(template_file)
        #写入指定的字段，这里的字段和模板文件中设置的字段对应
        template.merge(num = str(count), vuln_name = vuln[0], vuln_risk = vuln[1], vuln_category = name, vuln_detail = vuln[2], vuln_solution = vuln[3])
        template.write(".//tmpDocx//test-{}.docx".format(count))


        name_list.append(pwd + ".//tmpDocx//test-{}.docx".format(count))

                    
    #如果destfile存在则删除
    if os.path.exists(destFile):
        os.remove(destFile)

        #这个速度慢，而且需要经常清理缓存，且打开的word会被关闭，没有安装office不能用
        # mergeDocx_win32(destFile, name_list)
        #这个速度保守估计快4倍左右，一切正常
        mergeDocx_pyDocx(destFile,name_list)
    else:
        # mergeDocx_win32(destFile, name_list)
        mergeDocx_pyDocx(destFile,name_list)



def mergeDocx_win32(destFileName,file_list):
    #系统的word程序，没有安装则会报错不能使用
    word = win32.gencache.EnsureDispatch('Word.Application')
    #word窗口隐形
    word.Visible = False
    output_file = word.Documents.Add()
    #win32 这个合并是倒序的，在这里要反转
    file_list.reverse()
    for name in file_list:
        #循环合并
        output_file.Application.Selection.Range.InsertFile(name)
        #删除test文件
        os.remove(name)
    output_file.SaveAs(destFileName)


def mergeDocx_pyDocx(destFileName,file_list):
    number = len(file_list)
    master = Document(file_list[0])

    docx_composer = Composer(master)

    for x in range(1,number):
        docx_tmp = Document(file_list[x])
        docx_composer.append(docx_tmp)
    docx_composer.save(destFileName)



def checkBT_click():
    global info_flag
    info_flag = not info_flag
    

    


if __name__ == "__main__":
    root = tk.Tk()
    root.title('漏洞整理工具V1.1  Powered By WY. @2019.11')

    file = ''

    #************
    #nessus 
    lb1 = tk.Label(root, text='Nessus CSV、HTML文件：')
    lb1.grid(row=0, column=0)

    text1 = tk.Text(root, width=80, height=1)
    
    text1.grid(row=0, column=1)
    text1.mark_set('here','1.0')
    b1 = tk.Button(root, text='请选择文件', command=selectFile1)
    b2 = tk.Button(root, text='整理', command=zhengli_nessus)
    b1.grid(row=0, column=2)
    b2.grid(row=0, column=3)
    #************************************************

    lb_text = tk.Label(root, text='结果输出：')
    lb_text.grid(row=4, column=0)

    result_text = tk.Text(root, width=80, height=7)
    result_text.grid(row=4,column=1)
    #全局 消息 级别漏洞
    info_flag = False
    lb_checkbt = tk.Checkbutton(root, text='是否输出\n消息级别\n漏洞?默认为否。', command=checkBT_click)
    lb_checkbt.grid(row=4, column=2)


    #使用说明
    lb_attention = tk.Label(root, text='使用说明：')
    lb_attention.grid(row=5, column=0)

    attention_text = tk.Text(root, width=80, height=7)
    attention_text.grid(row=5,column=1)
    #
    attention = '''Nessus 文件为CSV或HTML(Custom Group Vulnerabilities By HOST)格式 '''

    attention_text.insert('end', attention)
    attention_text.update()


    root.mainloop()