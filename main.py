import pdfplumber
import PyPDF2
import os
import shutil
import xlrd
import time

#修复CropBox问题
def fix_cropbox(pdf_path, output_path):
    #"""修复PDF文件的CropBox问题，并生成新的PDF文件."""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        for page in reader.pages:
            if "/CropBox" not in page:
                page.cropbox = page.mediabox
            writer.add_page(page)

    with open(output_path, "wb") as output_file:
        writer.write(output_file)

#获取原始信息
def get_raw_info(filesName):
    i = 0
    rawInfo = []
    fixedFileFullName = "fixed\\" + filesName
    with pdfplumber.open(fixedFileFullName) as pdf:
        oneRawInfo = []
        for page in pdf.pages:
            tables = page.extract_text()
            for table in tables:
                for row in table:
                    infoInLine = [row]
                    oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]
    return(rawInfo)

#将原始数据转换为可读格式
def make_it_readable(raw_data):
    readable_data = []
    world = ""
    i = 0
    while i < len(raw_data):
        world += raw_data[i][0]
        if raw_data[i] == '\n':
            readable_data += [world]
            world = ""
        i = i + 1
    #print(readable_data)
    return readable_data

#获取pdf中信息
def get_pdf_info(pdf_data):
    pdf_info = []
    i = 0
    finalOrNot = ""
    poNo = ""
    while i < len(pdf_data):
        if pdf_data[i] == "FINAL CERTIFICATE OF ACCEPTANCE\n":
            finalOrNot = "FCOA"
        if pdf_data[i] == "CERTIFICATE OF ACCEPTANCE\n":
            finalOrNot = "COA"
        if pdf_data[i][:12] == "PO NUMBER : ":
            a = 0
            while a < len(pdf_data[i]) - 12:
                if pdf_data[i][12 + a] == ' ':
                    break
                poNo += pdf_data[i][12 + a]
                a = a + 1
            print("PO Number: ", poNo)
            print("Final or Not: ", finalOrNot)
            pdf_info = [poNo, finalOrNot]
        i = i + 1
    return pdf_info #格式：[PO Number, Final or Not]
    
#读取xlsx文件
def read_xlsx(xlsxFileName):
    xlsxData = xlrd.open_workbook(xlsxFileName).sheets()[0]
    print("读取xlsx文件: ", xlsxFileName)
    print("工作表名称: ", xlsxData.name)    
    return xlsxData  # 返回整个工作表数据
    
#根据PO Number从xlsx文件中获取数据
def get_data_from_xlsx(poNo, xlsxData):
    i = 0
    
    poNo = float(poNo)  # 将PO Number转换为浮点数
    
    
    try:
        PO_list_raw = xlsxData.col_values(0)  # 获取第一列的所有PO Number
        loPO = xlsxData.col_values(0).index(poNo)  # 获取PO Number所在行的索引
        print("PO Number所在行的索引: ", loPO)
        siteID = xlsxData.cell_value(loPO, 23)
        projectName = xlsxData.cell_value(loPO, 1)
        print("Site ID: ", siteID, "PO Number: ", poNo, "Project Name: ", projectName)
        
    except:
        print("PO Number不存在于xlsx文件中")
        return None
    
    xlsx_info = [siteID, projectName] #格式：[Site ID, Project Name]
    
    return xlsx_info
    




#读取input文件夹内文件的文件名
folder_path = "input"
filesName = os.listdir(folder_path)             #以列表方式记录所有文件名


#如果存在fixed文件夹，则删除
if os.path.exists('fixed'):
    shutil.rmtree('fixed')
    print("已删除旧的fixed文件夹")
    
#创建fixed文件夹
os.mkdir('fixed')


#循环修复所有文件
i = 0
while i < len(filesName):
    fileFullName = "input\\" + filesName[i]
    fixedFileName = "fixed\\fixed_" + filesName[i]
    fix_cropbox(fileFullName, fixedFileName)
    i = i + 1


filesName = os.listdir("fixed")

#读取xlsx文件
xlsxFileName = "TM PO TRACKER.xlsx"
print("读取xlsx文件")
xlsxData = read_xlsx(xlsxFileName)
print("读取xlsx文件完成")

#输出修复后的文件名
all_info = []
i = 0
for i in range(len(filesName)):
    print("fixed文件夹内的文件名：", filesName[i])

    raw_data = get_raw_info(filesName[i])

    pdf_data = make_it_readable(raw_data[0])
    
    pdf_info = get_pdf_info(pdf_data)
    
    xlsx_info = get_data_from_xlsx(pdf_info[0], xlsxData) 
    
    one_info = pdf_info + xlsx_info
    
    print("单个文件所有信息：", one_info)
    
    timeNow = "" + time.strftime('%Y-%m-%d', time.localtime())
    
    output_file_name = "input\\" + xlsx_info[0] + "_TM_" + xlsx_info[1] + "_PO" + pdf_info[0] + "_" + pdf_info[1] + "_" + timeNow + ".pdf"
    print("输出文件名：", output_file_name)
    print("原文件名", filesName[i][6:])
    os.rename("input\\"+filesName[i][6:],output_file_name)
    
    i = i + 1





shutil.rmtree('fixed')