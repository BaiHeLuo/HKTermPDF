"""
Flask Web应用 - 前端界面 + 后端生成PDF
启动后自动打开浏览器
打包方式: pyinstaller app.spec
"""

import os
import sys
import io
import zipfile
import tempfile
import webbrowser
import threading
from flask import Flask, render_template_string, request, send_file, jsonify

from pdf_generator import read_workers_from_excel, fill_pdf, validate_worker

app = Flask(__name__)


def get_base_dir():
    """获取程序所在目录（兼容 PyInstaller 打包后的路径）"""
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后，数据文件在 _MEIPASS 临时目录
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))


BASE_DIR = get_base_dir()
TEMPLATE_PDF = os.path.join(BASE_DIR, "勞工終止合約通知書-.pdf")

HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="zh-Hant">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>輸入勞工終止僱傭合約通知書 - 批量生成</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, "Microsoft JhengHei", "微软雅黑", sans-serif;
            background: #f0f2f5;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: #fff;
            border-radius: 16px;
            box-shadow: 0 4px 24px rgba(0,0,0,0.08);
            padding: 40px;
            max-width: 580px;
            width: 100%;
        }
        h1 { font-size: 22px; color: #1a1a2e; margin-bottom: 8px; }
        .subtitle { font-size: 14px; color: #888; margin-bottom: 32px; }
        .form-group { margin-bottom: 24px; }
        label {
            display: block; font-size: 14px; font-weight: 600;
            color: #333; margin-bottom: 8px;
        }
        label .hint { font-weight: 400; color: #999; font-size: 12px; }
        input[type="file"] { display: none; }
        .file-upload {
            border: 2px dashed #d0d5dd; border-radius: 10px; padding: 24px;
            text-align: center; cursor: pointer; transition: all 0.2s; background: #fafbfc;
        }
        .file-upload:hover { border-color: #4a6cf7; background: #f5f7ff; }
        .file-upload.has-file { border-color: #22c55e; background: #f0fdf4; }
        .file-upload .icon { font-size: 32px; margin-bottom: 8px; }
        .file-upload .text { font-size: 14px; color: #666; }
        .file-upload .filename { font-size: 14px; color: #22c55e; font-weight: 600; margin-top: 4px; }
        input[type="date"] {
            width: 100%; padding: 12px 16px; border: 2px solid #e0e0e0;
            border-radius: 10px; font-size: 16px; color: #333;
            outline: none; transition: border-color 0.2s;
        }
        input[type="date"]:focus { border-color: #4a6cf7; }
        .btn {
            width: 100%; padding: 14px; background: #4a6cf7; color: #fff; border: none;
            border-radius: 10px; font-size: 16px; font-weight: 600;
            cursor: pointer; transition: background 0.2s; margin-top: 8px;
        }
        .btn:hover { background: #3b5de7; }
        .btn:disabled { background: #ccc; cursor: not-allowed; }

        /* 状态消息 */
        .status { margin-top: 16px; padding: 12px 16px; border-radius: 8px; font-size: 14px; display: none; }
        .status.error { display: block; background: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
        .status.success { display: block; background: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0; }
        .status.loading { display: block; background: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }
        .status.warning { display: block; background: #fffbeb; color: #b45309; border: 1px solid #fde68a; }
        .spinner {
            display: inline-block; width: 14px; height: 14px;
            border: 2px solid #2563eb; border-top-color: transparent;
            border-radius: 50%; animation: spin 0.8s linear infinite;
            margin-right: 6px; vertical-align: middle;
        }
        @keyframes spin { to { transform: rotate(360deg); } }

        /* 详细错误列表 */
        .error-panel {
            margin-top: 16px; border: 1px solid #fecaca; border-radius: 10px;
            background: #fef2f2; overflow: hidden; display: none;
        }
        .error-panel.show { display: block; }
        .error-panel-header {
            padding: 12px 16px; font-size: 14px; font-weight: 600;
            color: #dc2626; border-bottom: 1px solid #fecaca; background: #fef2f2;
            display: flex; align-items: center; justify-content: space-between;
        }
        .error-panel-header .count {
            background: #dc2626; color: #fff; font-size: 12px;
            padding: 2px 8px; border-radius: 10px; font-weight: 700;
        }
        .error-list {
            max-height: 260px; overflow-y: auto; padding: 8px 0;
        }
        .error-list .item {
            padding: 6px 16px; font-size: 13px; color: #991b1b;
            line-height: 1.6; border-bottom: 1px solid #fee2e2;
        }
        .error-list .item:last-child { border-bottom: none; }
        .error-list .item .who { font-weight: 600; }

        /* 成功面板 */
        .success-panel {
            margin-top: 16px; border: 1px solid #bbf7d0; border-radius: 10px;
            background: #f0fdf4; overflow: hidden; display: none;
        }
        .success-panel.show { display: block; }
        .success-panel .header {
            padding: 12px 16px; font-size: 14px; font-weight: 600; color: #16a34a;
        }
        .success-panel .worker-list {
            padding: 4px 16px 12px; font-size: 13px; color: #166534; line-height: 1.8;
        }
        .success-panel .worker-list .tag {
            display: inline-block; background: #dcfce7; border-radius: 4px;
            padding: 1px 8px; margin: 2px 4px 2px 0; font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>輸入勞工終止僱傭合約通知書</h1>
        <p class="subtitle">批量生成 PDF — 上傳 Excel，輸入日期，一鍵產出</p>

        <form id="mainForm">
            <div class="form-group">
                <label>Excel 檔案 <span class="hint">（包含勞工名單的 .xlsx 文件）</span></label>
                <label class="file-upload" id="fileLabel" for="excelFile">
                    <div class="icon">📄</div>
                    <div class="text" id="fileText">點擊選擇 Excel 檔案</div>
                    <div class="filename" id="fileName" style="display:none"></div>
                </label>
                <input type="file" id="excelFile" name="excelFile" accept=".xlsx,.xls">
            </div>

            <div class="form-group">
                <label>簽署日期 <span class="hint">（通知書底部的 Date）</span></label>
                <input type="date" id="signDate" name="signDate" required>
            </div>

            <button type="submit" class="btn" id="submitBtn" disabled>生成並下載 PDF</button>

            <div class="status" id="status"></div>
            <div class="error-panel" id="errorPanel">
                <div class="error-panel-header">
                    <span>⚠ 數據驗證問題</span>
                    <span class="count" id="errorCount">0</span>
                </div>
                <div class="error-list" id="errorList"></div>
            </div>
            <div class="success-panel" id="successPanel">
                <div class="header">✅ 生成完成</div>
                <div class="worker-list" id="workerList"></div>
            </div>
        </form>
    </div>

    <script>
        const $ = id => document.getElementById(id);
        const fileInput = $('excelFile');
        const fileLabel = $('fileLabel');
        const fileText = $('fileText');
        const fileName = $('fileName');
        const signDate = $('signDate');
        const submitBtn = $('submitBtn');
        const statusDiv = $('status');
        const errorPanel = $('errorPanel');
        const errorList = $('errorList');
        const errorCount = $('errorCount');
        const successPanel = $('successPanel');
        const workerList = $('workerList');

        // 默认今天
        signDate.value = new Date().toISOString().split('T')[0];

        fileInput.addEventListener('change', () => {
            if (fileInput.files.length > 0) {
                fileText.style.display = 'none';
                fileName.textContent = fileInput.files[0].name;
                fileName.style.display = 'block';
                fileLabel.classList.add('has-file');
            } else {
                fileText.style.display = 'block';
                fileName.style.display = 'none';
                fileLabel.classList.remove('has-file');
            }
            clearResults();
            checkReady();
        });

        signDate.addEventListener('change', () => { clearResults(); checkReady(); });

        function checkReady() {
            submitBtn.disabled = !(fileInput.files.length > 0 && signDate.value);
        }
        checkReady();

        function clearResults() {
            statusDiv.className = 'status';
            statusDiv.innerHTML = '';
            errorPanel.classList.remove('show');
            successPanel.classList.remove('show');
        }

        function showErrors(errors) {
            errorCount.textContent = errors.length;
            errorList.innerHTML = errors.map(e => {
                // 高亮 [姓名] 部分
                const match = e.match(/^(.*?\[.*?\])(.*)/);
                if (match) {
                    return `<div class="item"><span class="who">${esc(match[1])}</span>${esc(match[2])}</div>`;
                }
                return `<div class="item">${esc(e)}</div>`;
            }).join('');
            errorPanel.classList.add('show');
        }

        function showSuccess(workers) {
            workerList.innerHTML = workers.map(w =>
                `<span class="tag">${esc(w)}</span>`
            ).join('');
            successPanel.classList.add('show');
        }

        function esc(s) {
            const d = document.createElement('div');
            d.textContent = s;
            return d.innerHTML;
        }

        $('mainForm').addEventListener('submit', async (e) => {
            e.preventDefault();
            clearResults();
            submitBtn.disabled = true;
            statusDiv.className = 'status loading';
            statusDiv.innerHTML = '<span class="spinner"></span>正在生成 PDF，請稍候...';

            const formData = new FormData();
            formData.append('excelFile', fileInput.files[0]);

            // 转换日期: 2026-06-15 → 15/6/2026
            const parts = signDate.value.split('-');
            formData.append('signDate', `${parseInt(parts[2])}/${parseInt(parts[1])}/${parts[0]}`);

            try {
                const resp = await fetch('/generate', { method: 'POST', body: formData });
                const contentType = resp.headers.get('content-type') || '';

                if (!resp.ok) {
                    // 后端返回 JSON 错误
                    const data = await resp.json();
                    if (data.errors && data.errors.length > 0) {
                        // 验证错误（有详细列表）
                        statusDiv.className = 'status error';
                        statusDiv.textContent = `發現 ${data.errors.length} 個數據問題，請修正 Excel 後重試。`;
                        showErrors(data.errors);
                    } else {
                        statusDiv.className = 'status error';
                        statusDiv.textContent = '錯誤: ' + (data.error || '生成失敗');
                    }
                    return;
                }

                // 检查是否有 warnings header
                const warnings = resp.headers.get('X-Warnings');
                const workersHeader = resp.headers.get('X-Workers');

                const blob = await resp.blob();
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = '勞工終止合約通知書.zip';
                a.click();
                URL.revokeObjectURL(url);

                if (warnings) {
                    const warnList = JSON.parse(decodeURIComponent(warnings));
                    statusDiv.className = 'status warning';
                    statusDiv.textContent = '已生成並下載，但部分勞工有數據問題（仍嘗試生成）。';
                    showErrors(warnList);
                } else {
                    statusDiv.className = 'status success';
                    statusDiv.textContent = '生成完成！已開始下載 ZIP 檔案。';
                }

                if (workersHeader) {
                    showSuccess(JSON.parse(decodeURIComponent(workersHeader)));
                }

            } catch (err) {
                statusDiv.className = 'status error';
                statusDiv.textContent = '請求失敗: ' + err.message;
            } finally {
                submitBtn.disabled = false;
            }
        });
    </script>
</body>
</html>
"""


@app.route("/")
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route("/generate", methods=["POST"])
def generate():
    try:
        excel_file = request.files.get("excelFile")
        if not excel_file:
            return jsonify({"error": "请上传Excel文件"}), 400

        signing_date = request.form.get("signDate", "")
        if not signing_date:
            return jsonify({"error": "请输入签署日期"}), 400

        if not os.path.exists(TEMPLATE_PDF):
            return jsonify({"error": f"模板PDF不存在，请将「勞工終止合約通知書-.pdf」放到程序同目录下"}), 500

        # 保存上传文件
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            excel_file.save(tmp.name)
            tmp_path = tmp.name

        try:
            workers = read_workers_from_excel(tmp_path)
            if not workers:
                return jsonify({"error": "Excel中没有找到勞工数据，请确认表头包含「中文姓名」"}), 400

            # ===== 验证所有勞工数据 =====
            all_errors = []
            for w in workers:
                all_errors.extend(validate_worker(w))

            # 如果有严重错误（缺少必填字段），直接拒绝
            critical_errors = [e for e in all_errors if "缺少" in e or "格式錯誤" in e or "無法解析" in e]
            if critical_errors:
                return jsonify({"errors": all_errors}), 400

            # 非严重警告（如姓名过短），仍然生成但附带警告
            warnings = [e for e in all_errors if e not in critical_errors]

            # ===== 生成PDF =====
            zip_buffer = io.BytesIO()
            generated_names = []
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    for worker in workers:
                        name_cn = worker["中文姓名"]
                        filename = f"勞工終止合約通知書-{name_cn}.pdf"
                        output_path = os.path.join(tmp_dir, filename)
                        fill_pdf(TEMPLATE_PDF, worker, output_path, signing_date)
                        zf.write(output_path, filename)
                        generated_names.append(name_cn)

            zip_buffer.seek(0)

            resp = send_file(
                zip_buffer,
                mimetype="application/zip",
                as_attachment=True,
                download_name="勞工終止合約通知書.zip",
            )

            # 通过 header 传递附加信息
            if warnings:
                resp.headers["X-Warnings"] = _safe_header_json(warnings)
            resp.headers["X-Workers"] = _safe_header_json(generated_names)
            return resp

        finally:
            os.unlink(tmp_path)

    except ValueError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"生成失败: {str(e)}"}), 500


def _safe_header_json(data):
    """将数据编码为安全的 HTTP header 值（ASCII）"""
    import json
    import urllib.parse
    raw = json.dumps(data, ensure_ascii=False)
    return urllib.parse.quote(raw)


@app.route("/decode_header")
def decode_header():
    """辅助：解码 header（前端用 JS decodeURIComponent 即可，此路由仅做测试用）"""
    import urllib.parse
    val = request.args.get("v", "")
    return urllib.parse.unquote(val)


def open_browser_delayed():
    """延迟 1.5 秒后打开浏览器"""
    import time
    time.sleep(1.5)
    webbrowser.open("http://localhost:5000")


if __name__ == "__main__":
    print("=" * 50)
    print("  勞工終止合約通知書 - 批量生成系統")
    print("  正在啟動，請稍候...")
    print("=" * 50)

    # 自动打开浏览器
    threading.Thread(target=open_browser_delayed, daemon=True).start()

    app.run(host="127.0.0.1", port=5000, debug=False)
