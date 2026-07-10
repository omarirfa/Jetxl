import jetxl, pyarrow as pa, openpyxl, zipfile, os, traceback

def mk_table():
    return pa.table({
        "id":     pa.array(list(range(1000)), pa.int64()),
        "region": pa.array([["North","South","East","West","Central"][i%5] for i in range(1000)]),  # low-card -> shared
        "note":   pa.array([f"unique note {i}" for i in range(1000)]),  # high-card -> inline
        "amount": pa.array([i*1.5 for i in range(1000)], pa.float64()),
    })

passed=failed=0
def check(name, cond):
    global passed, failed
    if cond: passed+=1; print(f"  PASS  {name}")
    else: failed+=1; print(f"  FAIL  {name}")

print("=== 1. Column ref by NAME ===")
jetxl.write_sheet_arrow(mk_table(), "/tmp/t_name.xlsx",
    column_widths={"region": 25.0, "amount": "auto"},
    column_formats={"amount": "currency"},
    hidden_columns=["note"])
wb=openpyxl.load_workbook("/tmp/t_name.xlsx"); ws=wb.active
check("opens with name refs", ws["A1"].value=="id")
check("hidden col by name applied", ws.column_dimensions["C"].hidden)

print("=== 2. Column ref by INDEX ===")
jetxl.write_sheet_arrow(mk_table(), "/tmp/t_idx.xlsx",
    column_widths={1: 25.0, 3: "auto"},   # region=1, amount=3
    column_formats={3: "currency"},
    hidden_columns=[2])                    # note=2
wb=openpyxl.load_workbook("/tmp/t_idx.xlsx"); ws=wb.active
check("opens with index refs", ws["A1"].value=="id")
check("hidden col by index applied", ws.column_dimensions["C"].hidden)

print("=== 3. MIXED index + name ===")
jetxl.write_sheet_arrow(mk_table(), "/tmp/t_mix.xlsx",
    column_widths={0: 10.0, "region": 25.0},
    hidden_columns=[2, "amount"])
wb=openpyxl.load_workbook("/tmp/t_mix.xlsx"); ws=wb.active
check("opens with mixed refs", ws["A1"].value=="id")

print("=== 4. ERROR: unknown name (should raise ValueError) ===")
try:
    jetxl.write_sheet_arrow(mk_table(), "/tmp/t_err.xlsx", column_widths={"nonexistent": 10.0})
    check("unknown name raises", False)
except ValueError as e:
    check("unknown name raises ValueError", "nonexistent" in str(e))
except Exception as e:
    check(f"unknown name raises ValueError (got {type(e).__name__})", False)

print("=== 5. ERROR: out-of-range index (should raise IndexError) ===")
try:
    jetxl.write_sheet_arrow(mk_table(), "/tmp/t_err2.xlsx", hidden_columns=[99])
    check("oob index raises", False)
except IndexError as e:
    check("oob index raises IndexError", "99" in str(e))
except Exception as e:
    check(f"oob index raises IndexError (got {type(e).__name__})", False)

print("=== 6. SHARED STRINGS active + correct ===")
jetxl.write_sheet_arrow(mk_table(), "/tmp/t_sst.xlsx")
with zipfile.ZipFile("/tmp/t_sst.xlsx") as z:
    names=z.namelist()
check("sharedStrings.xml present", "xl/sharedStrings.xml" in names)
wb=openpyxl.load_workbook("/tmp/t_sst.xlsx"); ws=wb.active
cats=["North","South","East","West","Central"]
errs=sum(1 for r in range(2,1002) if ws.cell(row=r,column=2).value!=cats[(r-2)%5])
check("shared region column resolves correctly", errs==0)
# high-card note column should still be correct
nerrs=sum(1 for r in range(2,1002) if ws.cell(row=r,column=3).value!=f"unique note {r-2}")
check("inline note column resolves correctly", nerrs==0)

print("=== 7. ERROR: hyperlink with & (corruption fix) ===")
jetxl.write_sheet_arrow(mk_table(), "/tmp/t_url.xlsx",
    hyperlinks=[(1,0,"https://example.com/?x=1&y=2","link")])
wb=openpyxl.load_workbook("/tmp/t_url.xlsx"); ws=wb.active
check("file with &-URL opens (no corruption)", ws["A1"].value=="id")

print(f"\n=== RESULTS: {passed} passed, {failed} failed ===")

# ============================================================================
# fastzip ZIP-writer validation (added when jetxl replaced mtzip with its own
# in-memory DEFLATE ZIP writer). These guard the hand-rolled ZIP structure.
# ============================================================================
def test_fastzip():
    import zipfile, io
    print("=== 8. fastzip ZIP structure ===")
    p=f=0
    def ck(n,c):
        nonlocal p,f
        if c: p+=1; print(f"  PASS  {n}")
        else: f+=1; print(f"  FAIL  {n}")

    t=mk_table()
    by=jetxl.write_sheet_arrow_to_bytes(t)
    z=zipfile.ZipFile(io.BytesIO(by))
    ck("zipfile opens archive", True)
    ck("testzip (all CRCs valid)", z.testzip() is None)
    ck("has [Content_Types].xml", "[Content_Types].xml" in z.namelist())
    ck("has workbook", "xl/workbook.xml" in z.namelist())

    # tiny file: exercises Store-fallback path safely
    tiny=pa.table({"a": pa.array([1], pa.int64())})
    bt=jetxl.write_sheet_arrow_to_bytes(tiny)
    ck("tiny file valid CRC", zipfile.ZipFile(io.BytesIO(bt)).testzip() is None)

    # reproducibility: same input -> byte-identical archive (fixed timestamps)
    b1=jetxl.write_sheet_arrow_to_bytes(mk_table())
    b2=jetxl.write_sheet_arrow_to_bytes(mk_table())
    ck("reproducible (byte-identical output)", b1==b2)

    print(f"\n=== fastzip: {p} passed, {f} failed ===")
    return f==0

test_fastzip()