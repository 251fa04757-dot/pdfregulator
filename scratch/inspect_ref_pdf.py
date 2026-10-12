import pymupdf

doc = pymupdf.open(r"C:\Users\Sasidhar\OneDrive\Documents\112-126.pdf")
with open(r"scratch\ref_pdf_text.txt", "w", encoding="utf-8") as out:
    for i, page in enumerate(doc):
        out.write(f"\n=================== PAGE {i+1} ===================\n")
        out.write(page.get_text())
        img_list = page.get_images()
        out.write(f"\nImages on Page {i+1}: {len(img_list)}\n")
        for img in img_list:
            out.write(f"  xref: {img[0]}, smask: {img[1]}, w: {img[2]}, h: {img[3]}, bpc: {img[4]}, cs: {img[5]}\n")

print("Saved scratch/ref_pdf_text.txt")
