import json
import io
from flask import Flask, render_template, request, send_file, jsonify
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = Flask(__name__)

# 設定字體函數：確保繁體中文美觀，並支援動態粗體
def set_font(run, font_name='微軟正黑體', size=11, bold=False):
    run.font.name = font_name
    # 針對中文字體的特殊設定
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size)
    run.font.bold = bold

# 核心轉檔邏輯
def generate_docx(data):
    doc = Document()
    
    # 1. 取得文件總標題
    doc_title = data.get('document_title', '文件提取結果')
    
    # 2. 處理每一頁資料
    pages = data.get('pages', [])
    for i, page in enumerate(pages):
        # 加入該頁的章節標題 (如果該頁沒標題就用總標題)
        section_title = page.get('section_title', doc_title)
        st_p = doc.add_paragraph()
        st_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st_run = st_p.add_run(section_title)
        set_font(st_run, size=16, bold=True)

        # 加入頁碼標示
        page_label = page.get('page_label', '')
        if page_label:
            pl_p = doc.add_paragraph()
            pl_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            pl_run = pl_p.add_run(f"頁面：{page_label}")
            set_font(pl_run, size=9)

        # 3. 動態建立表格
        headers = page.get('headers', [])
        table_data = page.get('data', [])
        
        if headers:
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            # 設定動態標頭
            hdr_cells = table.rows[0].cells
            for idx, h_text in enumerate(headers):
                run = hdr_cells[idx].paragraphs[0].add_run(h_text)
                set_font(run, size=12, bold=True)
                hdr_cells[idx].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
            # 填入內容資料
            for row_data in table_data:
                row_cells = table.add_row().cells
                for col_idx, cell_value in enumerate(row_data):
                    # 第一欄通常是重點，設為粗體
                    is_bold = True if col_idx == 0 else False
                    run = row_cells[col_idx].paragraphs[0].add_run(str(cell_value))
                    set_font(run, bold=is_bold)

        # 若非最後一頁，加入分頁符號
        if i < len(pages) - 1:
            doc.add_page_break()

    # 將文件存入記憶體流
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream, doc_title

# 路由 1：首頁 (網頁介面)
@app.route('/')
def index():
    return render_template('index.html')

# 路由 2：手動上傳 (網頁 Form 表單使用)
@app.route('/convert', methods=['POST'])
def convert():
    if 'json_file' not in request.files:
        return "未選擇檔案", 400
    
    file = request.files['json_file']
    if file.filename == '':
        return "檔名為空", 400

    try:
        content = file.read().decode('utf-8')
        data = json.loads(content)
        docx_file, doc_title = generate_docx(data)
        
        return send_file(
            docx_file,
            as_attachment=True,
            download_name=f"{doc_title}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return f"解析錯誤: {str(e)}", 500

# 路由 3：API 接口 (Dify 自動化串接使用)
@app.route('/api/convert', methods=['POST'])
def api_convert():
    try:
        # 直接從 HTTP Body 讀取 JSON
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON data received"}), 400
            
        docx_file, doc_title = generate_docx(data)
        
        # 回傳 Word 檔案給 Dify
        return send_file(
            docx_file,
            as_attachment=True,
            download_name=f"{doc_title}.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        print(f"API Error: {str(e)}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)