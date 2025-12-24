import json
import io
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

def set_font(run, font_name='微軟正黑體', size=11, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold

def generate_docx(data):
    print("--- 開始進入生成邏輯 ---")
    
    # 【核心修正】強效解析邏輯：處理 List、String 或 Dict
    current_data = data
    
    # 1. 如果資料被包在 List 裡，先拆開
    if isinstance(current_data, list) and len(current_data) > 0:
        current_data = current_data[0]
        print("已拆解 List 包裝")

    # 2. 如果資料是字串 (String)，將其轉為字典 (Dict)
    if isinstance(current_data, str):
        try:
            print("偵測到字串格式，嘗試進行 JSON 解析...")
            current_data = json.loads(current_data)
        except Exception as e:
            print(f"JSON 解析失敗: {e}")
            return None, "JSON Parsing Error"

    # 3. 確保最終是 Dict
    if not isinstance(current_data, dict):
        print(f"錯誤：最終格式不符，類型為 {type(current_data)}")
        return None, "Invalid Final Structure"

    # --- 開始建立 Word ---
    doc = Document()
    doc_title = current_data.get('document_title', '文件提取結果')
    pages = current_data.get('pages', [])

    for i, page in enumerate(pages):
        section_title = page.get('section_title')
        if not section_title or section_title in ["null", "", None]:
            section_title = doc_title
            
        st_p = doc.add_paragraph()
        st_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st_run = st_p.add_run(str(section_title))
        set_font(st_run, size=16, bold=True)

        page_label = page.get('page_label', f"p.{i+1}")
        pl_p = doc.add_paragraph()
        pl_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        pl_run = pl_p.add_run(f"頁面：{page_label}")
        set_font(pl_run, size=9)

        headers = page.get('headers', ["成語", "解釋"])
        table_data = page.get('data', [])
        
        table = doc.add_table(rows=1, cols=len(headers))
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        for idx, h_text in enumerate(headers):
            run = hdr_cells[idx].paragraphs[0].add_run(str(h_text))
            set_font(run, size=12, bold=True)

        for row_data in table_data:
            row_cells = table.add_row().cells
            for col_idx, cell_value in enumerate(row_data):
                if col_idx < len(row_cells):
                    is_bold = True if col_idx == 0 else False
                    run = row_cells[col_idx].paragraphs[0].add_run(str(cell_value))
                    set_font(run, bold=is_bold)

        if i < len(pages) - 1:
            doc.add_page_break()

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    print(f"--- Word 生成成功 ---")
    return file_stream, doc_title

@app.route('/api/test', methods=['GET'])
def api_test():
    return "API 伺服器正常運作中！", 200

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'json_file' not in request.files:
        return "未選擇檔案", 400
    file = request.files['json_file']
    try:
        content = file.read().decode('utf-8')
        data = json.loads(content)
        docx_file, doc_title = generate_docx(data)
        return send_file(docx_file, as_attachment=True, download_name=f"{doc_title}.docx", mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return f"網頁版錯誤: {str(e)}", 500

@app.route('/api/convert', methods=['POST'])
def api_convert():
    print(">>> 收到 Dify API 請求 <<<")
    try:
        # 嘗試讀取 JSON Body
        data = request.get_json(silent=True)
        
        # 如果 request.get_json() 失敗（例如沒設 Header），手動從 Data 讀取
        if data is None:
            print("無法直接取得 JSON，嘗試讀取原始數據...")
            data = json.loads(request.data.decode('utf-8'))

        docx_file, doc_title = generate_docx(data)
        
        if docx_file is None:
            return jsonify({"error": "Data structure error"}), 400

        return send_file(
            docx_file, 
            as_attachment=True, 
            download_name=f"{doc_title}.docx", 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f"API 進入點嚴重錯誤: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)