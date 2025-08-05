import pdfplumber
import PyPDF2
import os
import shutil
import pandas

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
def read_xlsx(PoNo):
    rootDir = os.listdir()
    print("当前目录下的文件有：", rootDir)
    i = 0
    xlsxFileName = "TM PO TRACKER 20250804.xlsx"
    while i < len(rootDir):
        if rootDir[i][4:] == "xlsx":
            print("找到xlsx文件：", rootDir[i])
            xlsxFileName = rootDir[i]
            i = len(rootDir)  # 结束循环
        i = i + 1
    allData = pandas.read_excel(xlsxFileName)
    siteID = allData.loc[allData['PO NO'] == PoNo, 'Site ID']
    print(siteID)


os.mkdir('fixed')

#读取input文件夹内文件的文件名
folder_path = "input"
filesName = os.listdir(folder_path)             #以列表方式记录所有文件名



#循环修复所有文件
i = 0
while i < len(filesName):
    fileFullName = "input\\" + filesName[i]
    fixedFileName = "fixed\\fixed_" + filesName[i]
    fix_cropbox(fileFullName, fixedFileName)
    i = i + 1


filesName = os.listdir("fixed")

#输出修复后的文件名
i = 0
for i in range(len(filesName)):
    print("fixed文件夹内的文件名：", filesName[i])

    raw_data = get_raw_info(filesName[i])

    pdf_data = make_it_readable(raw_data[0])
    
    get_pdf_info(pdf_data)
    
    read_xlsx(get_pdf_info(pdf_data)[0])
    
    i = i + 1





shutil.rmtree('fixed')