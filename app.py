import json
import io
from flask import Flask, render_template, request, send_file
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
    doc = Document()
    
    # 標題
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(data.get('document_title', '成語彙編'))
    set_font(title_run, size=18, bold=True)

    # 內容
    for i, page in enumerate(data.get('pages', [])):
        page_p = doc.add_paragraph()
        page_run = page_p.add_run(f"頁碼：{page['page_label']}")
        set_font(page_run, size=14, bold=True)

        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        headers = data.get('table_headers', ['成語', '解釋'])
        for idx, h_text in enumerate(headers):
            run = hdr_cells[idx].paragraphs[0].add_run(h_text)
            set_font(run, size=12, bold=True)

        for row_data in page.get('data', []):
            row_cells = table.add_row().cells
            # 成語
            run0 = row_cells[0].paragraphs[0].add_run(row_data[0])
            set_font(run0, bold=True)
            # 解釋
            run1 = row_cells[1].paragraphs[0].add_run(row_data[1])
            set_font(run1)

        if i < len(data['pages']) - 1:
            doc.add_page_break()

    # 將 Word 存在記憶體
    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'json_file' not in request.files:
        return "未選擇檔案", 400
    
    file = request.files['json_file']
    if file.filename == '':
        return "檔名為空", 400

    try:
        content = file.read()
        data = json.loads(content)
        docx_file = generate_docx(data)
        
        return send_file(
            docx_file,
            as_attachment=True,
            download_name="成語彙編.docx",
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return f"發生錯誤: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)