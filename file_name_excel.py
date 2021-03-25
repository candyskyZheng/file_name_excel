import os
import sys
import os.path

import re
import os,sys

import xlrd
import xlwt


def del_xml(path):
    if os.path.isdir(path):
        files = os.listdir(path)
        workbook = xlwt.Workbook(encoding="utf-8")  # 新建一个工作簿
        xls_path = path.split('\\References')[0].split('气田\\')[1]
       
        sheet = workbook.add_sheet(xls_path)  # 在工作簿中新建一个表格
        sheet.write(0, 0, "作者")
        sheet.write(0, 1, "年份")
        sheet.write(0, 2, "文件名称")
        i = 1
        for file in files:
            n = 0
            for year in re.findall(r"\d+\.?\d*", file):
                while  n == 0:
                    n+=1
                    sheet.write(i, 1, year)
                    print(year)
                    zz = file.split(year)[0].strip(", ")
                    sheet.write(i, 0, zz)
                    print(zz)
                    wjmc = file.split(year)[1].strip(", ")
                    sheet.write(i, 2, wjmc)
                    print(wjmc)
            i += 1

        workbook.save(path +'.xls')  # 保存工作簿            
        print(xls_path, "xls格式表格写入数据成功！")
                    

def file_path(path):
    if os.path.isdir(path):
        file = os.listdir(path)
        for xml in file:
            if xml =='References':
                del_xml(os.path.join(path,xml))
            else:
                file_path(os.path.join(path,xml))


if __name__ == '__main__':

    path = r'D:\ziliao\中国2020热点油气田'
    file_path(path=path)
    
    
