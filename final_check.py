from openpyxl import load_workbook

wb = load_workbook('result/keywordList_all.xlsx')
ws = wb['recent30days_sorted']
print('=== recent30days_sorted 시트 최종 결과 ===')
print('총 행 수: {}'.format(ws.max_row))
print('총 열 수: {}'.format(ws.max_column))

print('\n헤더:')
for col in range(1, ws.max_column + 1):
    header = ws.cell(row=1, column=col).value
    col_letter = chr(64+col)
    print('  {}열: {}'.format(col_letter, header))

print('\n처음 10행 데이터:')
for row in range(2, min(12, ws.max_row + 1)):
    c_val = ws.cell(row=row, column=3).value
    d_val = ws.cell(row=row, column=4).value
    e_val = ws.cell(row=row, column=5).value
    f_val = ws.cell(row=row, column=6).value
    print('행{}: C={}, D={}, E={}, F={}'.format(row, c_val, d_val, e_val, f_val))

print('\n노란색 강조된 행들 (처음 5개):')
yellow_count = 0
for row in range(2, ws.max_row + 1):
    cell_fill = ws.cell(row=row, column=1).fill
    if cell_fill and hasattr(cell_fill, 'start_color') and cell_fill.start_color and cell_fill.start_color.rgb == 'FFFFFF00':
        c_val = ws.cell(row=row, column=3).value
        d_val = ws.cell(row=row, column=4).value
        f_val = ws.cell(row=row, column=6).value
        print('행{}: C={}, D={}, F={}'.format(row, c_val, d_val, f_val))
        yellow_count += 1
        if yellow_count >= 5:
            break

print('\n총 노란색 강조 행 수: {}'.format(yellow_count))

# E열 값으로 정렬 확인
print('\nE열 값 분포 (처음 10개):')
e_values = []
for row in range(2, ws.max_row + 1):
    e_val = ws.cell(row=row, column=5).value
    if e_val is not None:
        e_values.append(float(e_val))

e_values.sort()
print('정렬된 E열 값들: {}'.format(e_values[:10]))