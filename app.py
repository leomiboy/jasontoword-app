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
    
    # 1. 取得文件總標題 (作為首頁標題)
    doc_title = data.get('document_title', '文件提取結果')
    
    # 2. 處理每一頁資料
    for i, page in enumerate(data.get('pages', [])):
        # 加入該頁的章節標題 (Section Title)
        section_title = page.get('section_title', doc_title)
        st_p = doc.add_paragraph()
        st_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        st_run = st_p.add_run(section_title)
        set_font(st_run, size=16, bold=True)

        # 加入頁碼標示
        page_label = page.get('page_label', '')
        pl_p = doc.add_paragraph()
        pl_run = pl_p.add_run(f"頁面：{page_label}")
        set_font(pl_run, size=10, bold=False)

        # 3. 動態建立表格
        headers = page.get('headers', [])
        if headers:
            table = doc.add_table(rows=1, cols=len(headers))
            table.style = 'Table Grid'
            
            # 設定動態標頭
            hdr_cells = table.rows[0].cells
            for idx, h_text in enumerate(headers):
                run = hdr_cells[idx].paragraphs[0].add_run(h_text)
                set_font(run, size=12, bold=True)
        
            # 填入資料
            for row_data in page.get('data', []):
                row_cells = table.add_row().cells
                for col_idx, cell_value in enumerate(row_data):
                    # 判斷是否為第一欄 (通常是重點關鍵字)，設為粗體
                    is_bold = True if col_idx == 0 else False
                    run = row_cells[col_idx].paragraphs[0].add_run(str(cell_value))
                    set_font(run, bold=is_bold)

        if i < len(data['pages']) - 1:
            doc.add_page_break()

    file_stream = io.BytesIO()
    doc.save(file_stream)
    file_stream.seek(0)
    return file_stream, doc_title

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
        
        # 根據文件標題動態命名下載檔案
        download_name = f"{doc_title}.docx"
        
        return send_file(
            docx_file,
            as_attachment=True,
            download_name=download_name,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        return f"解析錯誤: {str(e)}", 500

if __name__ == '__main__':
    app.run(debug=True)