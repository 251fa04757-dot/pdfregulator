import docx

doc = docx.Document("uploads/5.Machine Learning(II-II) (4C).docx")
for idx, t in enumerate(doc.tables):
    print(f"=== Table {idx}: {len(t.rows)} rows x {len(t.columns)} cols ===")
    for r_idx, r in enumerate(t.rows):
        cells = [c.text.strip().replace('\n', ' ') for c in r.cells]
        print(f"Row {r_idx}: {cells}")
