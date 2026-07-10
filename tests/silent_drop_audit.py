"""Behavioral silent-drop audit: apply EACH feature on EACH of the 4 arrow
paths and verify it actually took effect in the output. Reports any path where
a feature is silently ignored."""
import jetxl, pyarrow as pa, openpyxl, io, zipfile, tempfile, os, re

def tbl(n=20):
    cats=['North','South','East','West','Central']
    return pa.table({'id':pa.array(range(n),pa.int64()),
                     'region':pa.array([cats[i%5] for i in range(n)]),
                     'amount':pa.array([i*1.5 for i in range(n)],pa.float64())})

def write(path, arrow, kw):
    if path=='single_file':
        f=tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False); f.close()
        jetxl.write_sheet_arrow(arrow, f.name, **kw); d=open(f.name,'rb').read(); os.unlink(f.name); return d
    if path=='single_bytes':
        return jetxl.write_sheet_arrow_to_bytes(arrow, **kw)
    if path=='multi_file':
        f=tempfile.NamedTemporaryFile(suffix='.xlsx',delete=False); f.close()
        s={'data':arrow,'name':'S1'}; s.update(kw)
        jetxl.write_sheets_arrow([s], f.name, 1); d=open(f.name,'rb').read(); os.unlink(f.name); return d
    if path=='multi_bytes':
        s={'data':arrow,'name':'S1'}; s.update(kw)
        return jetxl.write_sheets_arrow_to_bytes([s], 1)

PATHS=['single_file','single_bytes','multi_file','multi_bytes']

# (feature name, kwargs, verifier(bytes)->bool)
def sheetxml(b): return zipfile.ZipFile(io.BytesIO(b)).read('xl/worksheets/sheet1.xml').decode()
def wb(b): return openpyxl.load_workbook(io.BytesIO(b))
def styles(b): return zipfile.ZipFile(io.BytesIO(b)).read('xl/styles.xml').decode()
def names(b): return zipfile.ZipFile(io.BytesIO(b)).namelist()

CHECKS = [
 ('auto_filter', {'auto_filter':True}, lambda b: wb(b).active.auto_filter.ref is not None),
 ('freeze_rows', {'freeze_rows':1}, lambda b: wb(b).active.freeze_panes is not None),
 ('freeze_cols', {'freeze_cols':2}, lambda b: wb(b).active.freeze_panes is not None),
 ('styled_headers', {'styled_headers':True}, lambda b: '<c r="A1" s=' in sheetxml(b)),
 ('write_header_row=False', {'write_header_row':False}, lambda b: wb(b).active['A1'].value==0),
 ('column_widths(name)', {'column_widths':{'amount':22.0}}, lambda b: wb(b).active.column_dimensions['C'].width is not None),
 ('column_widths(index)', {'column_widths':{2:22.0}}, lambda b: wb(b).active.column_dimensions['C'].width is not None),
 ('column_formats(name)', {'column_formats':{'amount':'currency'}}, lambda b: '$' in wb(b).active.cell(row=2,column=3).number_format),
 ('column_formats(index)', {'column_formats':{2:'currency'}}, lambda b: '$' in wb(b).active.cell(row=2,column=3).number_format),
 ('auto_width', {'auto_width':True}, lambda b: any(cd.width for cd in wb(b).active.column_dimensions.values())),
 ('merge_cells', {'merge_cells':[(1,0,1,2)]}, lambda b: 'A1:C1' in {str(r) for r in wb(b).active.merged_cells.ranges}),
 ('data_validations', {'data_validations':[{'start_row':2,'start_col':0,'end_row':20,'end_col':0,'type':'whole_number','min':0,'max':100}]}, lambda b: len(wb(b).active.data_validations.dataValidation)>=1),
 ('hyperlinks', {'hyperlinks':[(2,0,'https://x.com','h')]}, lambda b: 'hyperlink' in (zipfile.ZipFile(io.BytesIO(b)).read('xl/worksheets/_rels/sheet1.xml.rels').decode().lower() if 'xl/worksheets/_rels/sheet1.xml.rels' in names(b) else '')),
 ('row_heights', {'row_heights':{1:33.0}}, lambda b: wb(b).active.row_dimensions[1].height==33.0),
 ('cell_styles(font)', {'cell_styles':[{'row':2,'col':0,'font':{'bold':True}}]}, lambda b: wb(b).active.cell(row=2,column=1).font.bold is True),
 ('cell_styles(border)', {'cell_styles':[{'row':2,'col':0,'border':{'bottom':{'style':'thick'}}}]}, lambda b: wb(b).active.cell(row=2,column=1).border.bottom.style=='thick'),
 ('cell_styles(fill)', {'cell_styles':[{'row':2,'col':0,'fill':{'pattern':'solid','fg_color':'FFFFFF00'}}]}, lambda b: wb(b).active.cell(row=2,column=1).fill.patternType=='solid'),
 ('cell_styles(align+rot)', {'cell_styles':[{'row':2,'col':0,'alignment':{'horizontal':'center','text_rotation':45}}]}, lambda b: wb(b).active.cell(row=2,column=1).alignment.textRotation==45),
 ('formulas', {'formulas':[(5,2,'SUM(C2:C4)','9')]}, lambda b: str(wb(b).active.cell(row=5,column=3).value).startswith('=')),
 ('conditional_formats', {'conditional_formats':[{'start_row':2,'start_col':2,'end_row':20,'end_col':2,'rule_type':'cell_value','operator':'greater_than','value':'10'}]}, lambda b: '<conditionalFormatting' in sheetxml(b)),
 ('tables', {'tables':[{'name':'T1','start_row':0,'start_col':0,'end_row':20,'end_col':2}]}, lambda b: any('tables/table' in n for n in names(b))),
 ('charts', {'charts':[{'chart_type':'column','data_range':(0,2,20,2)}]}, lambda b: any('chart' in n and n.endswith('.xml') for n in names(b))),
 ('gridlines_visible=False', {'gridlines_visible':False}, lambda b: wb(b).active.sheet_view.showGridLines is False),
 ('zoom_scale', {'zoom_scale':140}, lambda b: wb(b).active.sheet_view.zoomScale==140),
 ('tab_color', {'tab_color':'FFFF0000'}, lambda b: wb(b).active.sheet_properties.tabColor is not None),
 ('default_row_height', {'default_row_height':18.0}, lambda b: wb(b).active.sheet_format.defaultRowHeight==18.0),
 ('hidden_columns(name)', {'hidden_columns':['region']}, lambda b: wb(b).active.column_dimensions['B'].hidden),
 ('hidden_columns(index)', {'hidden_columns':[1]}, lambda b: wb(b).active.column_dimensions['B'].hidden),
 ('hidden_rows', {'hidden_rows':[3]}, lambda b: wb(b).active.row_dimensions[3].hidden),
 ('right_to_left', {'right_to_left':True}, lambda b: wb(b).active.sheet_view.rightToLeft is True),
]

drops=[]
print(f"{'feature':<26}{'single_file':<14}{'single_bytes':<14}{'multi_file':<14}{'multi_bytes':<14}")
print("-"*82)
for name, kw, verify in CHECKS:
    row=f"{name:<26}"
    for path in PATHS:
        try:
            b=write(path, tbl(), kw)
            r=verify(b)
            row += f"{'ok' if r else 'DROP':<14}"
            if not r: drops.append((name,path))
        except Exception as e:
            row += f"{'ERR:'+type(e).__name__:<14}"
            drops.append((name,path,repr(e)))
    print(row)

print("\n"+"="*82)
if drops:
    print(f"SILENT DROPS / ERRORS FOUND: {len(drops)}")
    for d in drops: print("  ", d)
else:
    print("NO SILENT DROPS — every feature took effect on every path")
