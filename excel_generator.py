import re
from difflib import SequenceMatcher
import xlsxwriter
from docx import Document
import os
import concurrent.futures

def _run_comparison_task(args):
    data_b, data_a = args
    matcher = SequenceMatcher(None, data_b, data_a, autojunk=False)
    return matcher.get_opcodes()

def create_excel_report(before_file_path, after_file_path, excel_save_path, log_callback, 
                        paras_before=None, paras_after=None, get_loc_cb=None,
                        flags_b=None, flags_a=None, tables_before=None, tables_after=None):
    """
    [Strategy 15] Safe Data Writing Engine.
    Fixes the 'empty cell' bug caused by incorrect xlsxwriter rich_string usage.
    """
    if log_callback: log_callback(f"-> Excel 보고서(데이터 쓰기 안정성 강화) 생성 중...")

    workbook = xlsxwriter.Workbook(excel_save_path)
    header_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D3D3D3'})
    del_fmt = workbook.add_format({'font_color': 'blue', 'font_strikeout': True, 'valign': 'vcenter', 'text_wrap': True})
    ins_fmt = workbook.add_format({'font_color': 'red', 'bold': True, 'valign': 'vcenter', 'text_wrap': True})
    def_fmt = workbook.add_format({'valign': 'vcenter', 'text_wrap': True})
    loc_fmt = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    table_cell_fmt = workbook.add_format({'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    table_ins_fmt = workbook.add_format({'font_color': 'red', 'bold': True, 'valign': 'vcenter', 'border': 1, 'text_wrap': True})
    table_del_fmt = workbook.add_format({'font_color': 'blue', 'font_strikeout': True, 'valign': 'vcenter', 'border': 1, 'text_wrap': True})

    def get_rich_diff(text_b, text_a):
        if not text_b: return [], [ins_fmt, text_a]
        words_b = [w for w in re.split(r'(\s+)', text_b) if w]; words_a = [w for w in re.split(r'(\s+)', text_a) if w]
        w_matcher = SequenceMatcher(None, words_b, words_a, autojunk=False)
        rich_b, rich_a = [], []
        for w_tag, w_i1, w_i2, w_j1, w_j2 in w_matcher.get_opcodes():
            fb, fa = "".join(words_b[w_i1:w_i2]), "".join(words_a[w_j1:w_j2])
            if w_tag == 'equal':
                if fb: rich_b.extend([def_fmt, fb]); rich_a.extend([def_fmt, fa])
            elif w_tag == 'delete': rich_b.extend([del_fmt, fb])
            elif w_tag == 'insert': rich_a.extend([ins_fmt, fa])
            elif w_tag == 'replace': rich_b.extend([del_fmt, fb]); rich_a.extend([ins_fmt, fa])
        return rich_b, rich_a

    # (1) 일반 본문 시트
    main_ws = workbook.add_worksheet("변경 내용(일반)")
    main_ws.write_row('A1', ['위치', '수정 전', '수정 후'], header_fmt)
    main_ws.set_column('A:A', 25, loc_fmt); main_ws.set_column('B:C', 60, def_fmt)
    main_ws.freeze_panes(1, 0)
    p_b_filtered = [p for i, p in enumerate(paras_before) if not (flags_b and flags_b[i])]
    p_a_filtered = [p for i, p in enumerate(paras_after) if not (flags_a and flags_a[i])]
    tasks = [(p_b_filtered, p_a_filtered)]
    max_tables = max(len(tables_before or []), len(tables_after or []))
    for t_idx in range(max_tables):
        tb, ta = (tables_before[t_idx] if t_idx < len(tables_before) else []), (tables_after[t_idx] if t_idx < len(tables_after) else [])
        tasks.append(([str(r[0]).strip() if r else "" for r in tb], [str(r[0]).strip() if r else "" for r in ta]))
        tasks.append(([str(c).strip() for c in tb[0]] if tb and tb[0] else [], [str(c).strip() for c in ta[0]] if ta and ta[0] else []))
    try:
        with concurrent.futures.ProcessPoolExecutor() as executor: all_results = list(executor.map(_run_comparison_task, tasks))
    except: all_results = [_run_comparison_task(t) for t in tasks]
    main_opcodes = all_results[0]; excel_row = 1; orig_idx_b, orig_idx_a = [i for i, f in enumerate(flags_b) if not f], [i for i, f in enumerate(flags_a) if not f]
    for tag, i1, i2, j1, j2 in main_opcodes:
        if tag == 'equal': continue
        content_b, content_a = "\n".join(p_b_filtered[i1:i2]).strip(), "\n".join(p_a_filtered[j1:j2]).strip()
        if not content_b and not content_a: continue
        rb, ra = get_rich_diff(content_b, content_a)
        main_ws.write(excel_row, 0, f"{get_loc_cb(orig_idx_b[i1], True) if get_loc_cb else '문단'}", loc_fmt)
        for col, r_data, f_text in [(1, rb, content_b), (2, ra, content_a)]:
            # [SAFE WRITE] Rich String은 최소 3개의 인자가 필요함 (포맷, 텍스트, 포맷...)
            if len(r_data) >= 3 and len(r_data) <= 500:
                try: main_ws.write_rich_string(excel_row, col, *r_data, def_fmt)
                except: main_ws.write(excel_row, col, f_text, def_fmt)
            else:
                main_ws.write(excel_row, col, f_text, (ins_fmt if col==2 and r_data else (del_fmt if col==1 and r_data else def_fmt)))
        excel_row += 1

    # (2) 표 시트
    for t_idx in range(max_tables):
        sheet_name = f"표 {t_idx + 1}"
        ws = workbook.add_worksheet(sheet_name[:31])
        tb, ta = (tables_before[t_idx] if t_idx < len(tables_before) else []), (tables_after[t_idx] if t_idx < len(tables_after) else [])
        max_cols_b = max([len(r) for r in tb]) if tb else 0
        after_start_col = max_cols_b + 1 if max_cols_b > 0 else 0
        ws.write(0, 0, "수정 전", header_fmt); ws.write(0, after_start_col, "수정 후", header_fmt)
        row_opcodes, col_opcodes = all_results[1 + t_idx*2], all_results[2 + t_idx*2]
        row_map = {j_idx: i_idx for tag, i1, i2, j1, j2 in row_opcodes if tag in ('equal', 'replace') for i_idx, j_idx in zip(range(i1, i2), range(j1, j2))}
        col_map = {j_idx: i_idx for tag, i1, i2, j1, j2 in col_opcodes if tag in ('equal', 'replace') for i_idx, j_idx in zip(range(i1, i2), range(j1, j2))}
        
        for r_idx, row in enumerate(tb):
            for c_idx, val in enumerate(row): ws.write(r_idx + 2, c_idx, val, table_cell_fmt)

        for r_idx, row in enumerate(ta):
            for c_idx, val_a in enumerate(row):
                orig_r = row_map.get(r_idx); orig_c = col_map.get(c_idx)
                is_changed = True
                if orig_r is not None and orig_c is not None and orig_r < len(tb) and orig_c < len(tb[orig_r]):
                    val_b = tb[orig_r][orig_c]
                    if val_b == val_a: is_changed = False
                    else:
                        _, ra = get_rich_diff(val_b, val_a)
                        # [SAFE WRITE] Rich String 안정성 검사
                        if len(ra) >= 3 and len(ra) <= 500:
                            try:
                                ws.write_rich_string(r_idx + 2, c_idx + after_start_col, *ra, table_cell_fmt)
                                continue
                            except: pass
                
                # Rich String이 불가능하거나 데이터가 변경된 경우 일반 쓰기
                ws.write(r_idx + 2, c_idx + after_start_col, val_a, table_ins_fmt if is_changed else table_cell_fmt)

    workbook.close()
    if log_callback: log_callback(f"-> 무결성 데이터 쓰기 완료: {excel_save_path}")
