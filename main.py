from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import re
from docx import Document
import fitz  # PyMuPDF
import io
import os

app = Flask(__name__)
CORS(app)  # Allow cross-origin requests

# --- Helper Functions from the original script ---

def extract_korean(text):
    """주어진 텍스트에서 한글만 추출하여 공백으로 연결합니다."""
    korean_chars = re.findall(r'[가-힣]+', text)
    return " ".join(korean_chars)

def process_name_line_s1(line):
    """'*'로 시작하는 줄을 분석하여 이름 형식을 변경합니다."""
    if not line.startswith('*'):
        return None
    text = line.lstrip('*').strip()
    match = re.match(r'([가-힣\s]+)([a-zA-Z\s]+)', text)
    if not match:
        return None
    korean_part, english_part = match.group(1).strip(), match.group(2).strip()
    korean_names = korean_part.split()
    reformatted_korean = f"{korean_names[-1]}, {' '.join(korean_names[:-1])}" if len(korean_names) >= 2 else korean_part
    english_names = english_part.split()
    reformatted_english = f"{english_names[-1]}, {' '.join(english_names[:-1])}" if len(english_names) >= 2 else english_part
    return f"{reformatted_korean}{reformatted_english}"

# --- Processing Logic for Each Tab ---

def handle_tab1(file, regex_pattern, decorator_str):
    """색인 추출 (Tab 1) 로직"""
    doc = Document(file)
    full_text = "\n".join([para.text for para in doc.paragraphs])
    matches = re.findall(regex_pattern, full_text)

    if not matches:
        raise ValueError("정규식과 일치하는 내용을 찾지 못했습니다.")

    prefix, suffix = ('', '')
    if len(decorator_str) == 1:
        prefix = suffix = decorator_str
    elif len(decorator_str) >= 2:
        prefix, suffix = decorator_str[0], decorator_str[1]

    string_matches = [f"{t[0]}@{t[1]}@" if isinstance(t, tuple) and len(t) >= 2 else str(t) for t in matches]
    unique_matches = list(dict.fromkeys(string_matches))
    
    formatted_matches = []
    for item in unique_matches:
        parts = item.strip('@').split('@')
        if len(parts) == 2 and parts[0] and parts[1]:
            formatted_matches.append(f"{parts[0]}{prefix}{parts[1]}{suffix}")
        else:
            formatted_matches.append(item)

    new_doc = Document()
    new_doc.add_heading("추출된 색인 목록", level=1)
    for item in formatted_matches:
        new_doc.add_paragraph(item)
    
    # Save to an in-memory stream
    file_stream = io.BytesIO()
    new_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def handle_tab2(file):
    """인명 정리 (Tab 2) 로직"""
    source_doc = Document(file)
    result_doc = Document()
    table = result_doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text, hdr_cells[1].text = '원본 내용', '변경 내용'
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].font.bold = True
    
    for para in source_doc.paragraphs:
        original_text = para.text.strip()
        if not original_text:
            continue
        changed_text = process_name_line_s1(original_text)
        row_cells = table.add_row().cells
        row_cells[0].text, row_cells[1].text = original_text, changed_text if changed_text else ""
        
    file_stream = io.BytesIO()
    result_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def handle_tab3(file):
    """한글 추출 (Tab 3) 로직"""
    source_doc = Document(file)
    if not source_doc.tables:
        raise ValueError("선택한 파일에 표(테이블)가 없습니다.")
    
    source_table = source_doc.tables[0]
    result_doc = Document()
    result_table = result_doc.add_table(rows=1, cols=3)
    result_table.style = 'Table Grid'
    hdr_cells = result_table.rows[0].cells
    hdr_cells[0].text, hdr_cells[1].text, hdr_cells[2].text = '원본 내용 (*제거)', '한글 추출', '변경 내용'
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].font.bold = True
        
    for row in source_table.rows[1:]:
        col1_text, col2_text = row.cells[0].text, row.cells[1].text
        row_cells = result_table.add_row().cells
        row_cells[0].text = col1_text.lstrip('*').strip()
        row_cells[1].text = extract_korean(col1_text)
        row_cells[2].text = col2_text
        
    file_stream = io.BytesIO()
    result_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def handle_tab4(pdf_file, docx_file, page_range_str):
    """PDF 페이지 찾기 (Tab 4) 로직"""
    index_doc = Document(docx_file)
    if not index_doc.tables:
        raise ValueError("색인 파일에 테이블이 없습니다.")
    
    table = index_doc.tables[0]
    search_terms = [row.cells[1].text.strip() for i, row in enumerate(table.rows) if i > 0 and len(row.cells) >= 2 and row.cells[1].text.strip()]
    if not search_terms:
        raise ValueError("색인 테이블의 2열에서 검색할 단어를 찾을 수 없습니다.")

    pdf_doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    # Parse page range
    total_pages = pdf_doc.page_count
    page_range_iterator = range(total_pages) # Default to all pages
    if page_range_str:
        match = re.match(r'^\s*(\d+)\s*-\s*(\d+)\s*$', page_range_str)
        if match:
            start_page = int(match.group(1)) - 1
            end_page = int(match.group(2))
            if 0 <= start_page < end_page <= total_pages:
                page_range_iterator = range(start_page, end_page)

    normalized_terms_map = {term: "".join(term.split()) for term in search_terms}
    results = {term: [] for term in search_terms}
    
    for page_num in page_range_iterator:
        page = pdf_doc.load_page(page_num)
        page_text_normalized = "".join(page.get_text("text").split())
        if not page_text_normalized: continue
        for original_term, normalized_term in normalized_terms_map.items():
            if normalized_term in page_text_normalized and (page_num + 1) not in results[original_term]:
                results[original_term].append(page_num + 1)
    pdf_doc.close()

    new_doc = Document()
    new_doc.add_heading("PDF 색인 분석 결과", level=1)
    original_table = index_doc.tables[0]
    new_table = new_doc.add_table(rows=0, cols=len(original_table.columns) + 1)
    new_table.style = 'Table Grid'

    for i, row in enumerate(original_table.rows):
        new_row_cells = new_table.add_row().cells
        for j, cell in enumerate(row.cells):
            new_row_cells[j].text = cell.text
        
        if i == 0:
            new_row_cells[-1].text = "페이지"
        else:
            term_to_find = row.cells[1].text.strip()
            found_pages = results.get(term_to_find)
            new_row_cells[-1].text = ", ".join(map(str, sorted(found_pages))) if found_pages else "PDF에서 찾을 수 없음"

    file_stream = io.BytesIO()
    new_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream

def handle_tab5(file):
    """테이블 조합 (Tab 5) 로직"""
    doc = Document(file)
    if not doc.tables:
        raise ValueError("업로드된 워드 파일에 테이블이 없습니다.")
    table = doc.tables[0]
    if len(table.columns) < 4:
        raise ValueError("테이블이 4열 이상으로 구성되어야 합니다.")
        
    processed_lines = []
    for row in table.rows[1:]:
        col2_text, col3_text, col4_text = row.cells[1].text.strip(), row.cells[2].text.strip(), row.cells[3].text.strip()
        if col2_text and not col3_text:
            processed_lines.append(f"{col2_text} {col4_text}")
        elif col2_text and col3_text:
            processed_lines.append(f"{col3_text} {col4_text}")
    
    if not processed_lines:
        raise ValueError("테이블에서 처리할 데이터를 찾지 못했습니다.")

    unique_sorted_lines = sorted(list(set(processed_lines)))
    
    new_doc = Document()
    new_doc.add_heading("테이블 정리 결과 (가나다순 정렬)", level=1)
    for line in unique_sorted_lines:
        new_doc.add_paragraph(line)
        
    file_stream = io.BytesIO()
    new_doc.save(file_stream)
    file_stream.seek(0)
    return file_stream


# --- API Endpoint ---

@app.route('/api/process/<task_name>', methods=['POST'])
def process_task(task_name):
    try:
        if task_name == 'tab1':
            file = request.files.get('file')
            regex = request.form.get('regex')
            decorator = request.form.get('decorator')
            if not file:
                return jsonify({"message": "파일이 없습니다."}), 400
            result_stream = handle_tab1(file, regex, decorator)
            filename = f"result_{task_name}.docx"

        elif task_name in ['tab2', 'tab3', 'tab5']:
            file = request.files.get('file')
            if not file:
                return jsonify({"message": "파일이 없습니다."}), 400
            
            if task_name == 'tab2':
                result_stream = handle_tab2(file)
            elif task_name == 'tab3':
                result_stream = handle_tab3(file)
            else: # tab5
                result_stream = handle_tab5(file)
            filename = f"result_{task_name}.docx"

        elif task_name == 'tab4':
            pdf_file = request.files.get('pdf_file')
            docx_file = request.files.get('docx_file')
            page_range = request.form.get('page_range')
            if not pdf_file or not docx_file:
                return jsonify({"message": "PDF 또는 DOCX 파일이 없습니다."}), 400
            result_stream = handle_tab4(pdf_file, docx_file, page_range)
            filename = f"result_{task_name}.docx"

        else:
            return jsonify({"message": "알 수 없는 작업입니다."}), 404

        return send_file(
            result_stream,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except ValueError as e:
        return jsonify({"message": str(e)}), 400
    except Exception as e:
        print(f"An error occurred: {e}")
        return jsonify({"message": "서버 내부 오류가 발생했습니다."}), 500
