import re
from difflib import SequenceMatcher
import xlsxwriter
from docx import Document
import os

def _extract_paragraphs_from_docx(file_path):
    """
    Extract paragraphs and table cell texts from a .docx file using python-docx.
    Note: This is a fallback and does not capture automatic numbering or perfect order 
    if tables are interleaved with paragraphs in complex ways.
    """
    try:
        doc = Document(file_path)
        all_texts = []
        
        # In python-docx, to get interleaved order, we need to iterate over document elements.
        # But for simplicity, we keep the previous logic as fallback, 
        # while primary path will be via Win32 COM in main.py.
        for p in doc.paragraphs:
            if p.text.strip():
                all_texts.append(p.text.strip())

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if paragraph.text.strip():
                            all_texts.append(paragraph.text.strip())
        return all_texts
    except Exception as e:
        raise Exception(f"Failed to parse DOCX file {file_path}: {str(e)}")

def create_excel_report(before_file_path, after_file_path, excel_save_path, log_callback, paras_before=None, paras_after=None):
    """
    Compares two Word documents and generates an Excel report with changes.
    The report formats changed words and preserves document structure.
    """
    if log_callback:
        log_callback(f"-> Excel 보고서 생성 중...")

    # If paragraph lists are not provided, extract them using fallback method
    if paras_before is None or paras_after is None:
        if log_callback:
            log_callback("-> '전'/'후' 문서에서 텍스트를 추출합니다 (기본 방식)...")
        try:
            if paras_before is None:
                paras_before = _extract_paragraphs_from_docx(before_file_path)
            if paras_after is None:
                paras_after = _extract_paragraphs_from_docx(after_file_path)
        except Exception as e:
            if log_callback:
                log_callback(f"오류: 문서 파일을 읽는 중 문제가 발생했습니다. {e}")
            return

    workbook = xlsxwriter.Workbook(excel_save_path)
    worksheet = workbook.add_worksheet("변경 내용")

    # Formats
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'bg_color': '#D3D3D3'})
    deleted_format = workbook.add_format({'font_color': 'blue', 'font_strikeout': True, 'valign': 'vcenter', 'text_wrap': True})
    inserted_format = workbook.add_format({'font_color': 'red', 'bold': True, 'valign': 'vcenter', 'text_wrap': True})
    default_format = workbook.add_format({'valign': 'vcenter', 'text_wrap': True})

    # Setup header
    worksheet.write('A1', '위치', header_format)
    worksheet.write('B1', '수정 전', header_format)
    worksheet.write('C1', '수정 후', header_format)
    worksheet.set_column('A:A', 15, default_format)
    worksheet.set_column('B:B', 60, default_format)
    worksheet.set_column('C:C', 60, default_format)
    worksheet.freeze_panes(1, 0)
    
    excel_row = 1
    if log_callback:
        log_callback("-> 문단 변경 사항을 비교합니다...")

    # Compare at paragraph level
    matcher = SequenceMatcher(None, paras_before, paras_after, autojunk=False)
    
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            continue
            
        # For non-equal blocks (replace, delete, insert), create a row in Excel.
        # i1:i2 are indices in paras_before, j1:j2 are indices in paras_after.
        
        location_str = ""
        if tag == 'delete':
            location_str = f"{i1 + 1}~{i2}행 (삭제됨)" if i2 - i1 > 1 else f"{i1 + 1}행 (삭제됨)"
        elif tag == 'insert':
            location_str = f"{j1 + 1}~{j2}행 (추가됨)" if j2 - j1 > 1 else f"{j1 + 1}행 (추가됨)"
        elif tag == 'replace':
            loc_b = f"{i1 + 1}~{i2}" if i2 - i1 > 1 else f"{i1 + 1}"
            loc_a = f"{j1 + 1}~{j2}" if j2 - j1 > 1 else f"{j1 + 1}"
            location_str = f"{loc_b} -> {loc_a}행"

        content_b = "\n".join(paras_before[i1:i2])
        content_a = "\n".join(paras_after[j1:j2])
        
        # Tokenize preserving all whitespace and newlines
        words_b = [w for w in re.split(r'(\s+)', content_b) if w]
        words_a = [w for w in re.split(r'(\s+)', content_a) if w]
        
        word_matcher = SequenceMatcher(None, words_b, words_a, autojunk=False)

        rich_b = []
        for w_tag, w_i1, w_i2, w_j1, w_j2 in word_matcher.get_opcodes():
            text_fragment = "".join(words_b[w_i1:w_i2])
            if not text_fragment:
                continue
            
            if w_tag == 'equal':
                rich_b.extend([default_format, text_fragment])
            elif w_tag == 'delete' or w_tag == 'replace':
                rich_b.extend([deleted_format, text_fragment])
        
        rich_a = []
        for w_tag, w_i1, w_i2, w_j1, w_j2 in word_matcher.get_opcodes():
            text_fragment = "".join(words_a[w_j1:w_j2])
            if not text_fragment:
                continue

            if w_tag == 'equal':
                rich_a.extend([default_format, text_fragment])
            elif w_tag == 'insert' or w_tag == 'replace':
                rich_a.extend([inserted_format, text_fragment])

        worksheet.write(excel_row, 0, location_str)
        
        # Write rich strings if they are not empty and within Excel limits
        # Excel rich strings have a limit of 255 fragments.
        if rich_b:
            try:
                if len(rich_b) > 500: # 2 items per fragment (format, text)
                    worksheet.write(excel_row, 1, content_b, default_format)
                else:
                    worksheet.write_rich_string(excel_row, 1, *rich_b)
            except Exception:
                worksheet.write(excel_row, 1, content_b, default_format)
        
        if rich_a:
            try:
                if len(rich_a) > 500:
                    worksheet.write(excel_row, 2, content_a, default_format)
                else:
                    worksheet.write_rich_string(excel_row, 2, *rich_a)
            except Exception:
                worksheet.write(excel_row, 2, content_a, default_format)
                
        excel_row += 1
    
    workbook.close()
    if log_callback:
        if excel_row > 1:
            log_callback(f"-> Excel 보고서 저장 완료: {excel_save_path}")
        else:
            log_callback("-> 텍스트 변경 사항이 없어 Excel 보고서를 생성하지 않습니다.")
            try:
                if os.path.exists(excel_save_path):
                    os.remove(excel_save_path)
            except OSError:
                pass
