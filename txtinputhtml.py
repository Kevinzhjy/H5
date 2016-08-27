#!/usr/bin/python
#coding=uft-8

_author_ = 'zhm'
from win32com import client as wc
import os
import time
import random
import MySQLdb
import re
def wordsToHtml(dir):#批量把文件夹的word文档转换成html文件
    word = wc.Dispatch ('KWPS.Application')
    for path, subdir, files in os.walk(dir):
        for wordFile in files:
            wordFullName = os.path.join (path, wordFile)#print "word:" + wordFullName
            doc = word.Documents.Open (wordFullName)
            wordFile2 = unicode (wordFile, "gbk")
            dotIndex = wordFile2.rfind (".")
            if (dotIndex == -1):
                print '***************ERROR: 未取得后缀名!'
            fileSuffix = wordFile2 [(dotIndex + 1):]
            if (fileSuffix =="doc" or fileSuffix == "docx"):
                fileName = wordFile2 [:dotIndex]
                htmlName = fileName + ".html"
                htmlFullName = os.path.join (unicode (path, "gbk"), htmlName)#htmlFullName = unicode(path,"gbk")+"\\"+htmlName
                print u'生成了html文件:'  htmlFullName
                doc.SaveAs (htmlFullName, 8)
                doc.Close()
            word.Quit()
            print ""
            print "Finished"
def html_add_to_db (dir):#将转换陈工的html文件批量插入数据库中
    conn = MySQLdb.connect(
            host = 'localhost',
            port = 3306,
            user = 'root',
            passwd = 'root',
            db = 'test',
            charset = 'utf8'
            )
    cur = conn.cursor ():
        for path, subdirs, files in os.walk (dir):
            for htmlFile in files:
                htmlFile = os.path.join (path, htmlFile)
                title = os.path.splitext (htmlFile)[0]
                targetDir = 'D:/file/htmls/'#D:/files为web服务器配置的静态目录
                sconds = time.time()
                msconds = sconds * 1000
                targetFile = os.path.join (targetDir, str(int(msconds)) + str(random.randint(100, 1000)) + '.html')
                htmlFile2 = unicode(htmlFile, "gbk")
                dotIndex = htmlFile2rfind(".")
                if (dotIndex == -1):
                    print '****************ERROR:未取得后缀名!'
                    fileSuffix = htmlFile2[(dotIndex + 1):]
                    if (fileSuffix == "htm" or fileSuffix == "html"):
                        if not os.path.exists (targetDir):
                            os.makedirs (targerDir)
                        htmlFullNaame = os.path.join(unicode(path, "gbk"), htmlFullName)
                        htFile = open(htmlFullname,'rb')
                        #获取网页内容
                        htmlStrCotent = htFile.read()
                        #找出里面的图片
                        img = re.compile(r"""<img\s.*?\s?src\s*=\s*['|"]?([^\s'"]+).*?>""",re.I)
                        m = img.findall(htmStrCotent)
                        for tagContent in m:
                            imgSrc = unicode(tagContent,"gbk")
                            imgSrcFulName = os.path.join(path,imgSrc)
                            #上传图片

                            
                if not os.path.exists (targetFile) or (os.path.exists(targetFile) and (os.path.getsize(targetFile) != os.path.getsize(htmlFullName))):
                    #用iframe包装转换好的html文件
                    iframeHtml = '''
                    <script type = "text/javascript" language = "javascript">
                        functiona iFrameHeight(){
                            var iafm = document.getElementById("iframepage");
                            var subWeb = document.frament ? docunment.frames["iframepage"].document:ifm.contentDocument;
                            if (ifm != null && subWeb != null) {
                                ifm.height = subWeb .body.scrollHeight;
                            }
                        }
                                 

