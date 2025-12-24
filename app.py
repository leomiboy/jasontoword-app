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
    print("--- 開始生成 Word 檔案 ---")
    
    # 【核心優化】深度解包：不論包了幾層 List，一路拆到看到 Dict 為止
    current_data = data
    while isinstance(current_data, list):
        if len(current_data) > 0:
            print(f"解開一層清單包裝，剩餘內容類型: {type(current_data[0])}")
            current_data = current_data[0]
        else:
            print("錯誤：收到空的資料清單")
            return None, "Empty List"
    
    if not isinstance(current_data, dict):
        print(f"嚴重錯誤：最終解析結果不是字典，而是 {type(current_data)}")
        return None, "Invalid Data Structure"

    doc = Document()
    doc_title = current_data.get('document_title', '文件提取結果')
    pages = current_data.get('pages', [])

    if not pages:
        print("警告：解析出的字典中沒有頁面資料")

    for i, page in enumerate(pages):
        section_title = page.get('section_title')
        if not section_title or section_title in ["null", ""]:
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

        headers = page.get('headers', ["欄位1", "欄位2"])
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
    print(f"--- Word 生成成功：{doc_title} ---")
    return file_stream, doc_title

@app.route('/api/test', methods=['GET'])
def api_test():
    return "API 伺服器運作中！", 200

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'json_file' not in request.files:
        return "未選擇檔案", 400
    file = request.files['json_file']
    try:
        data = json.loads(file.read().decode('utf-8'))
        docx_file, doc_title = generate_docx(data)
        return send_file(docx_file, as_attachment=True, download_name=f"{doc_title}.docx", mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    except Exception as e:
        return f"解析錯誤: {str(e)}", 500

@app.route('/api/convert', methods=['POST'])
def api_convert():
    print(">>> 收到 Dify API 請求 <<<")
    try:
        data = request.get_json()
        print(f"原始數據類型: {type(data)}")
        
        docx_file, doc_title = generate_docx(data)
        if docx_file is None:
            return jsonify({"error": "Data structure error", "details": doc_title}), 400

        return send_file(
            docx_file, 
            as_attachment=True, 
            download_name=f"{doc_title}.docx", 
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f"API 進入點發生錯誤: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)