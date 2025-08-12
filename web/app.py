import os
import shutil
import zipfile
import tempfile
import traceback
from flask import Flask, request, render_template, send_file, flash, jsonify
from main import process_files
from threading import Thread

app = Flask(__name__)
app.secret_key = "super_secret_key"

UPLOAD_FOLDER = tempfile.mkdtemp()
OUTPUT_FOLDER = tempfile.mkdtemp()
LOG_FILE = "error.log"

# 全局进度状态
progress_status = {
    "percent": 0,
    "message": "等待开始",
    "files": [],
    "finished": False
}

def clean_temp_folders():
    """清空上传和输出文件夹"""
    for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
        if os.path.exists(folder):
            shutil.rmtree(folder)
        os.mkdir(folder)

def run_process(pdf_paths, xlsx_path):
    """后台线程执行文件处理"""
    global progress_status
    try:
        progress_status.update({"percent": 5, "message": "开始处理文件...", "files": [], "finished": False})
        clean_temp_folders()

        # 复制文件到 UPLOAD_FOLDER
        for pdf_path in pdf_paths:
            shutil.copy(pdf_path, UPLOAD_FOLDER)
        shutil.copy(xlsx_path, UPLOAD_FOLDER)

        progress_status.update({"percent": 20, "message": "文件已上传，开始运行脚本..."})

        # 调用用户的处理逻辑
        error_files = process_files(UPLOAD_FOLDER, xlsx_path, OUTPUT_FOLDER)

        progress_status.update({"percent": 90, "message": "打包结果文件..."})
        # 打包 ZIP
        zip_path = os.path.join(OUTPUT_FOLDER, "result.zip")
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in os.listdir(OUTPUT_FOLDER):
                if file != "result.zip":
                    zipf.write(os.path.join(OUTPUT_FOLDER, file), file)

        # 列出生成的文件名
        files_generated = [f for f in os.listdir(OUTPUT_FOLDER) if f != "result.zip"]
        progress_status.update({
            "percent": 100,
            "message": "处理完成",
            "files": files_generated,
            "finished": True
        })

        if error_files:
            flash("以下 PDF 处理失败：" + "，".join(error_files))

    except Exception as e:
        with open(LOG_FILE, "a", encoding="utf-8") as log:
            log.write("=== 发生错误 ===\n")
            log.write(traceback.format_exc() + "\n")
        progress_status.update({"percent": 100, "message": f"处理出错：{str(e)}", "finished": True})

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/start_process", methods=["POST"])
def start_process():
    pdf_files = request.files.getlist("pdfs")
    xlsx_file = request.files.get("xlsx")
    if not pdf_files or not xlsx_file:
        return jsonify({"error": "请上传 PDF 文件和一个 XLSX 文件"}), 400

    # 先保存文件到临时路径（避免 read of closed file 错误）
    saved_pdf_paths = []
    for pdf in pdf_files:
        pdf_path = os.path.join(tempfile.gettempdir(), pdf.filename)
        pdf.save(pdf_path)
        saved_pdf_paths.append(pdf_path)

    xlsx_path = os.path.join(tempfile.gettempdir(), xlsx_file.filename)
    xlsx_file.save(xlsx_path)

    # 重置进度
    progress_status.update({"percent": 0, "message": "准备开始", "files": [], "finished": False})

    # 启动后台线程处理
    t = Thread(target=run_process, args=(saved_pdf_paths, xlsx_path))
    t.start()

    return jsonify({"message": "任务已启动"})

@app.route("/progress", methods=["GET"])
def progress():
    return jsonify(progress_status)

@app.route("/download")
def download_file():
    zip_path = os.path.join(OUTPUT_FOLDER, "result.zip")
    if not os.path.exists(zip_path):
        return "结果文件不存在，请重新上传处理。", 404
    return send_file(zip_path, as_attachment=True)

@app.route("/download_file/<path:filename>")
def download_single_file(filename):
    file_path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(file_path):
        return "文件不存在", 404
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
