import pdfplumber
import PyPDF2
import os
import shutil
import xlrd
import time

def fix_cropbox(pdf_path, output_path):
    """修复PDF的CropBox并写入output_path"""
    with open(pdf_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()
        for page in reader.pages:
            # 如果没有 CropBox，使用 mediabox 作为 cropbox
            try:
                if "/CropBox" not in page:
                    page.cropbox = page.mediabox
            except Exception:
                # 某些 PyPDF2 版本 page 对象不同，尝试直接设置属性
                try:
                    page.cropbox = page.mediabox
                except Exception:
                    pass
            writer.add_page(page)
    with open(output_path, "wb") as output_file:
        writer.write(output_file)

def get_raw_info(filesName, fixed_folder):
    """从修复后的PDF中提取每行文本，返回行列表（每行以换行符结尾）"""
    fixedFileFullName = os.path.join(fixed_folder, filesName)
    lines = []
    with pdfplumber.open(fixedFileFullName) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                for line in text.splitlines():
                    # 保持换行符以兼容旧逻辑
                    lines.append(line + "\\n")
    return lines

def make_it_readable(raw_data):
    """
    将原始行列表转换为“可读”列表（这里保持每行为一个元素）。
    旧版本里有更复杂的合并逻辑，但基于pdfplumber的输出，
    直接使用行列表通常更稳定。
    """
    if not raw_data:
        return []
    # raw_data 已经是行的列表，直接返回
    return raw_data

def get_pdf_info(pdf_data):
    """从 pdf_data 的行列表中提取 PO Number 和是否为 FINAL 信息"""
    pdf_info = []
    finalOrNot = ""
    poNo = ""
    for line in pdf_data:
        if line.strip() == "FINAL CERTIFICATE OF ACCEPTANCE":
            finalOrNot = "FCOA"
        if line.strip() == "CERTIFICATE OF ACCEPTANCE":
            finalOrNot = "COA"
        # 支持以 'PO NUMBER : ' 开头（注意空格）
        if line.startswith("PO NUMBER : ") or line.startswith("PO NUMBER: "):
            # 去掉前缀，再取第一个由数字组成的连续片段作为 PO
            remainder = line.split(":",1)[1].strip()
            # 取首个由数字/字母混合的片段（直到遇到空格）
            poNo_candidate = remainder.split()[0]
            poNo = poNo_candidate
            pdf_info = [poNo, finalOrNot]
            break
    return pdf_info

def read_xlsx(xlsxFileName):
    """
    读取 xlsx/xls 文件。优先使用 xlrd（历史兼容），若失败尝试 openpyxl。
    返回一个包装对象，提供 col_values(index) 和 cell_value(row, col) 接口。
    """
    try:
        # 尝试使用 xlrd
        xlsxData = xlrd.open_workbook(xlsxFileName).sheets()[0]
        return xlsxData
    except Exception:
        # 如果 xlrd 无法读取（比如新版本不支持 xlsx），尝试 openpyxl
        try:
            from openpyxl import load_workbook
            wb = load_workbook(xlsxFileName, read_only=True, data_only=True)
            sheet = wb[wb.sheetnames[0]]
            # 包装成类似 xlrd 的接口
            class SheetWrapper:
                def __init__(self, sheet):
                    self.sheet = sheet
                    self.name = sheet.title
                    self._rows = list(sheet.values)
                def col_values(self, idx):
                    vals = []
                    for row in self._rows:
                        if idx < len(row):
                            vals.append(row[idx])
                        else:
                            vals.append(None)
                    return vals
                def cell_value(self, r, c):
                    if r < 0 or r >= len(self._rows):
                        raise IndexError("row out of range")
                    row = self._rows[r]
                    if c < len(row):
                        return row[c]
                    return None
            return SheetWrapper(sheet)
        except Exception as e:
            raise RuntimeError(f"无法读取 xlsx 文件: {e}")

def get_data_from_xlsx(poNo, xlsxData):
    """根据 PO Number 从 xlsx 表格里查找对应行并返回 [siteID, projectName]"""
    try:
        # 尝试将 poNo 转换为 float，再在第一列中查找（兼容数字格式）
        search_key = float(poNo)
        col0 = xlsxData.col_values(0)
        loPO = col0.index(search_key)
        siteID = xlsxData.cell_value(loPO, 23)
        projectName = xlsxData.cell_value(loPO, 1)
        return [siteID, projectName]
    except Exception:
        # 若不能转换为 float，尝试作为字符串直接匹配（某些 PO 可能包含字母）
        try:
            col0 = xlsxData.col_values(0)
            # 尝试字符串匹配
            for idx, val in enumerate(col0):
                if val is None:
                    continue
                if str(val).strip() == str(poNo).strip():
                    siteID = xlsxData.cell_value(idx, 23)
                    projectName = xlsxData.cell_value(idx, 1)
                    return [siteID, projectName]
        except Exception:
            pass
    return None

def _sanitize_filename(name):
    """去掉文件名中的非法字符，简单处理空格"""
    import re
    name = str(name)
    name = name.strip()
    # 替换斜杠等文件系统禁止字符
    name = re.sub(r'[\\\\/*?:"<>|]', '_', name)
    name = name.replace(' ', '_')
    return name

def process_files(input_folder, xlsx_path, output_folder):
    """
    主处理函数：
    - 修复 PDF 的 CropBox，产生 fixed/fixed_<name>.pdf
    - 依次从 fixed 中读取文本，提取 PO Number，并从 xlsx 中查找 siteID 与 projectName
    - 如果匹配成功，按规则命名输出文件；如果匹配失败，仍复制文件到 output，并以 unmatched_ 前缀标注
    - 返回一个 error_files 列表（哪些原始 PDF 名称在处理时遇到问题）
    """
    fixed_folder = os.path.join(input_folder, "fixed")
    if os.path.exists(fixed_folder):
        shutil.rmtree(fixed_folder)
    os.mkdir(fixed_folder)

    error_files = []

    # 1) 修复上传的 PDF 到 fixed 文件夹
    for file_name in os.listdir(input_folder):
        if not file_name.lower().endswith(".pdf"):
            continue
        src_path = os.path.join(input_folder, file_name)
        dst_path = os.path.join(fixed_folder, "fixed_" + file_name)
        try:
            fix_cropbox(src_path, dst_path)
        except Exception:
            # 如果修复失败，尝试简单复制以便后续仍可读取
            try:
                shutil.copyfile(src_path, dst_path)
            except Exception:
                # 无法复制或修复则记录为错误文件
                error_files.append(file_name)
                continue

    # 2) 读取 xlsx（外部异常向上抛出让调用者处理）
    xlsxData = read_xlsx(xlsx_path)

    # 3) 处理 fixed 中的每个文件
    for fixed_name in os.listdir(fixed_folder):
        fixed_path = os.path.join(fixed_folder, fixed_name)
        original_name = fixed_name[6:] if fixed_name.startswith("fixed_") else fixed_name
        try:
            raw_lines = get_raw_info(fixed_name, fixed_folder)
            if not raw_lines:
                # 没有抽取到文本，视为未匹配但仍输出原文件
                out_name = "unmatched_" + _sanitize_filename(original_name)
                out_path = os.path.join(output_folder, out_name)
                shutil.copyfile(fixed_path, out_path)
                error_files.append(original_name)
                continue

            pdf_data = make_it_readable(raw_lines)
            pdf_info = get_pdf_info(pdf_data)
            if not pdf_info or not pdf_info[0]:
                out_name = "unmatched_" + _sanitize_filename(original_name)
                out_path = os.path.join(output_folder, out_name)
                shutil.copyfile(fixed_path, out_path)
                error_files.append(original_name)
                continue

            poNo = pdf_info[0]
            finalOrNot = pdf_info[1] if len(pdf_info) > 1 else ""

            xlsx_info = get_data_from_xlsx(poNo, xlsxData)
            if not xlsx_info:
                # 无法从 xlsx 中匹配到 PO，仍然输出原文件并标注 unmatched
                out_name = "unmatched_PO" + _sanitize_filename(original_name)
                out_path = os.path.join(output_folder, out_name)
                shutil.copyfile(fixed_path, out_path)
                error_files.append(original_name)
                continue

            siteID, projectName = xlsx_info
            timeNow = time.strftime('%Y-%m-%d', time.localtime())
            out_filename = f"{_sanitize_filename(siteID)}_TM_{_sanitize_filename(projectName)}_PO{_sanitize_filename(poNo)}_{_sanitize_filename(finalOrNot)}_{timeNow}.pdf"
            out_path = os.path.join(output_folder, out_filename)
            shutil.copyfile(fixed_path, out_path)
        except Exception:
            # 捕获单文件处理异常，记录并把文件以 unmatched_ 前缀写入 output
            try:
                out_name = "error_" + _sanitize_filename(original_name)
                out_path = os.path.join(output_folder, out_name)
                shutil.copyfile(fixed_path, out_path)
            except Exception:
                pass
            error_files.append(original_name)

    return error_files
