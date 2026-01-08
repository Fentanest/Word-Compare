
import re
from difflib import SequenceMatcher
import xlsxwriter
from docx import Document
import os
import os

def _extract_paragraphs_from_docx(file_path):
    """Extract paragraphs from a .docx file."""
    try:
        doc = Document(file_path)
        return [p.text for p in doc.paragraphs if p.text.strip()]
    except Exception as e:
        raise Exception(f"Failed to parse DOCX file {file_path}: {str(e)}")

def create_excel_report(before_file_path, after_file_path, excel_save_path, log_callback):
    """
    Compares two Word documents and generates an Excel report with changes.
    The report includes context (2 paragraphs before and after) and formats changed words.
    """
    if log_callback:
        log_callback(f"-> Excel 보고서 생성 중...")
        log_callback("-> '전'/'후' 문서에서 텍스트를 추출합니다...")

    try:
        paras_before = _extract_paragraphs_from_docx(before_file_path)
        paras_after = _extract_paragraphs_from_docx(after_file_path)
    except Exception as e:
        if log_callback:
            log_callback(f"오류: 문서 파일을 읽는 중 문제가 발생했습니다. {e}")
        return

    workbook = xlsxwriter.Workbook(excel_save_path)
    worksheet = workbook.add_worksheet("변경 내용")

    # Formats
    header_format = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
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

    matcher = SequenceMatcher(None, paras_before, paras_after)
    
    # We process opcodes to find changes and avoid redundant reports for adjacent changes.
    # We'll merge consecutive changed blocks.
    opcodes = matcher.get_opcodes()
    i = 0
    while i < len(opcodes):
        tag, i1, i2, j1, j2 = opcodes[i]
        if tag != 'equal':
            # This is the start of a changed block
            start_i1, start_j1 = i1, j1
            end_i2, end_j2 = i2, j2
            
            # See if the next opcodes are also changes, to merge them
            while i + 1 < len(opcodes) and opcodes[i+1][0] != 'equal':
                i += 1
                _, _, end_i2, _, end_j2 = opcodes[i]
            
            # Now we have a merged block of changes from (start_i1, start_j1) to (end_i2, end_j2)
            location_str = f"{start_i1 + 1}~{end_i2}줄"

            # --- Gather context and content for "Before" and "After" cells ---
            
            # Context before
            context_start_i = max(0, start_i1 - 1)
            context_start_j = max(0, start_j1 - 1)
            context_before_str_b = "\n".join(paras_before[context_start_i:start_i1])
            context_before_str_a = "\n".join(paras_after[context_start_j:start_j1])

            # Changed content
            content_b = " ".join(paras_before[start_i1:end_i2])
            content_a = " ".join(paras_after[start_j1:end_j2])

            # Context after
            context_end_i = min(len(paras_before), end_i2 + 1)
            context_end_j = min(len(paras_after), end_j2 + 1)
            context_after_str_b = "\n".join(paras_before[end_i2:context_end_i])
            context_after_str_a = "\n".join(paras_after[end_j2:context_end_j])
            
            # --- Perform word-level diff on changed content ---
            words_b = content_b.split()
            words_a = content_a.split()
            word_matcher = SequenceMatcher(None, words_b, words_a)

            # --- Build rich string for '수정 전' (Before) ---
            rich_b = []
            if context_before_str_b:
                rich_b.extend([default_format, context_before_str_b + "\n"])
            
            for w_tag, w_i1, w_i2, w_j1, w_j2 in word_matcher.get_opcodes():
                if w_tag == 'equal':
                    rich_b.extend([default_format, " ".join(words_b[w_i1:w_i2]), " "])
                elif w_tag == 'delete' or w_tag == 'replace':
                    rich_b.extend([deleted_format, " ".join(words_b[w_i1:w_i2]), " "])
            
            if context_after_str_b:
                rich_b.extend([default_format, "\n" + context_after_str_b])

            # --- Build rich string for '수정 후' (After) ---
            rich_a = []
            if context_before_str_a:
                rich_a.extend([default_format, context_before_str_a + "\n"])

            for w_tag, w_i1, w_i2, w_j1, w_j2 in word_matcher.get_opcodes():
                if w_tag == 'equal':
                    rich_a.extend([default_format, " ".join(words_a[w_j1:w_j2]), " "])
                elif w_tag == 'insert' or w_tag == 'replace':
                    rich_a.extend([inserted_format, " ".join(words_a[w_j1:w_j2]), " "])

            if context_after_str_a:
                rich_a.extend([default_format, "\n" + context_after_str_a])

            # --- Write to Excel ---
            worksheet.write(excel_row, 0, location_str)
            if rich_b:
                # xlsxwriter needs pairs of (format, string)
                worksheet.write_rich_string(excel_row, 1, *rich_b)
            if rich_a:
                worksheet.write_rich_string(excel_row, 2, *rich_a)
            excel_row += 1
        
        i += 1
    
    workbook.close()
    if log_callback:
        if excel_row > 1:
            log_callback(f"-> Excel 보고서 저장 완료: {excel_save_path}")
        else:
            log_callback("-> 텍스트 변경 사항이 없어 Excel 보고서를 생성하지 않습니다.")
            # If no changes were written, delete the empty file
            try:
                os.remove(excel_save_path)
            except OSError:
                pass