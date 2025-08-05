import pdfplumber
import PyPDF2
import os
import shutil

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


def get_raw_info(filesName):
    i = 0
    #创建总输出数组

    rawInfo = []

    fixedFileFullName = "fixed\\" + filesName
    with pdfplumber.open(fixedFileFullName) as pdf:
        oneRawInfo = []
        for page in pdf.pages:
            print(page)
            tables = page.extract_text()
            for table in tables:
                for row in table:
                    #print(row)
                    infoInLine = [row]
                    oneRawInfo += infoInLine
        rawInfo += [oneRawInfo]

    return(rawInfo)

def make_it_readable(raw_data):
    #"""将原始数据转换为可读格式."""
    readable_data = []
    world = ""
    i = 0
    while i < len(raw_data):
        world += raw_data[i][0]
        if raw_data[i] == '\n':
            readable_data += [world]
            world = ""
        i = i + 1
        
    print(readable_data)
    return readable_data

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

    make_it_readable(raw_data[0])
    
    i = i + 1



#输出原始数据
'''
print(len(raw_data[0]))

i = 0
while i < len(raw_data):
    print(i, " ", raw_data[i])
    i = i + 1
'''


shutil.rmtree('fixed')