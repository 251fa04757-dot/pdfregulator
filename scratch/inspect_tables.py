import docx
from pathlib import Path

uploads_dir = Path("uploads")
for f in uploads_dir.glob("*.docx"):
    print(f"=== {f.name} ===")
    try:
        doc = docx.Document(str(f))
        for idx, t in enumerate(doc.tables):
            print(f"Table {idx}: {len(t.rows)} rows x {len(t.columns)} cols")
            for r_idx, row in enumerate(t.rows):
                cells_text = [c.text.strip().replace("\n", " ") for c in row.cells]
                # show preview
                print(f"  Row {r_idx}: {cells_text}")
    except Exception as e:
        print(f"Error {f.name}: {e}")
    print()
