import json, collections

with open('tools/ref_style_analysis.json', encoding='utf-8') as f:
    data = json.load(f)

ws = data['Feuille 1']

print('=== SHEET OVERVIEW ===')
print(f'  freeze_panes: {ws["freeze_panes"]}')
print(f'  dimensions: {ws["dimensions"]}')
print(f'  print_title_rows: {ws.get("print_title_rows")}')
print(f'  auto_filter: {ws.get("auto_filter")}')

print()
print('=== PAGE SETUP ===')
for k, v in ws.get('page_setup', {}).items():
    print(f'  {k}: {v}')

print()
print('=== MARGINS ===')
for k, v in ws.get('margins', {}).items():
    print(f'  {k}: {v}')

print()
print('=== COLUMN WIDTHS ===')
for col, info in sorted(ws.get('column_dimensions', {}).items()):
    print(f'  {col}: width={info["width"]}, hidden={info["hidden"]}')

print()
print('=== ROW HEIGHTS (non-default) ===')
for row_idx, info in sorted(ws.get('row_dimensions', {}).items(), key=lambda x: int(x[0])):
    if info.get('height') is not None:
        print(f'  Row {row_idx}: height={info["height"]}, hidden={info["hidden"]}')

print()
print('=== MERGED CELLS ===')
for mc in ws.get('merged_cells', []):
    print(f'  {mc}')

print()
print('=== CELL STYLES - FIRST 5 ROWS ===')
cell_styles = ws.get('cell_styles', {})
# Show all cells in rows 1-5
for row_idx in range(1, 6):
    for col in 'ABCDEFGHIJKLMN':
        key = f'{col}{row_idx}'
        if key in cell_styles:
            cs = cell_styles[key]
            print(f'  {key}: val={repr(cs["value"])[:40]} | font={cs["font"]["name"]} sz={cs["font"]["size"]} bold={cs["font"]["bold"]} color={cs["font"]["color"]} | fill={cs["fill"]["fg_color"]} | align={cs["alignment"]["horizontal"]} wrap={cs["alignment"]["wrap_text"]} indent={cs["alignment"]["indent"]} | fmt={cs["number_format"]}')

print()
print('=== UNIQUE FILL COLORS (all rows) ===')
fill_colors = collections.Counter()
for key, cs in cell_styles.items():
    fg = cs['fill']['fg_color']
    if fg not in ('NONE', 'None', '00000000', None) and cs['fill']['pattern_type'] not in (None, 'none'):
        fill_colors[fg] += 1
for color, count in fill_colors.most_common(20):
    print(f'  {color}: {count} cells')

print()
print('=== UNIQUE FONTS (name, size, bold combinations) ===')
font_combos = collections.Counter()
for key, cs in cell_styles.items():
    f = cs['font']
    combo = (f['name'], f['size'], f['bold'], f['color'])
    font_combos[combo] += 1
for combo, count in font_combos.most_common(15):
    print(f'  name={combo[0]} sz={combo[1]} bold={combo[2]} color={combo[3]}: {count}')

print()
print('=== NUMBER FORMATS USED ===')
fmt_counter = collections.Counter()
for key, cs in cell_styles.items():
    fmt = cs['number_format']
    if fmt and fmt != 'General':
        fmt_counter[fmt] += 1
for fmt, count in fmt_counter.most_common(15):
    print(f'  {repr(fmt)}: {count}')

print()
print('=== BORDERS SUMMARY ===')
has_border = 0
no_border = 0
border_styles = collections.Counter()
for key, cs in cell_styles.items():
    b = cs['border']
    for side in ['top','bottom','left','right']:
        if b.get(side):
            has_border += 1
            border_styles[b[side]['style']] += 1
        else:
            no_border += 1
print(f'  Cells with borders: {has_border}, without: {no_border}')
for style, count in border_styles.most_common():
    print(f'  Border style {style!r}: {count}')

print()
print('=== SAMPLE ROWS BY TYPE (detect section/article/total rows) ===')
# Show rows 1-30 with their key cell styles
for row_idx in range(1, 31):
    row_cells = {k: v for k, v in cell_styles.items() if k[1:] == str(row_idx) or (len(k) > 2 and k[2:] == str(row_idx))}
    if row_cells:
        a_key = f'A{row_idx}'
        b_key = f'B{row_idx}'
        f_key = f'F{row_idx}'
        a_info = cell_styles.get(a_key, {})
        b_info = cell_styles.get(b_key, {})
        f_info = cell_styles.get(f_key, {})
        a_val = a_info.get('value', '')
        b_val = b_info.get('value', '')
        f_val = f_info.get('value', '')
        a_fill = a_info.get('fill', {}).get('fg_color', '')
        b_bold = b_info.get('font', {}).get('bold', False)
        b_sz = b_info.get('font', {}).get('size', '')
        print(f'  R{row_idx}: A={repr(str(a_val)[:15])} B={repr(str(b_val)[:25])} F={repr(str(f_val)[:15])} fill={a_fill} bold={b_bold} sz={b_sz}')
