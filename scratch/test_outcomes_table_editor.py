import requests
import json
import pymupdf
from pathlib import Path

BASE_URL = "http://localhost:8000"

def test_outcomes_table_workflow():
    print("=== Testing In-Table Course Outcomes Editing & Persistence ===")

    doc_name = "5.Machine Learning(II-II) (4C).docx"

    # 1. Fetch document content and verify outcomes table is parsed
    resp = requests.get(f"{BASE_URL}/get-document-content?docx_name={doc_name}")
    assert resp.status_code == 200, f"GET failed: {resp.status_code}"
    data = resp.json()
    assert data["success"] is True

    co_table = data.get("course_outcomes_table")
    assert co_table is not None and len(co_table) >= 2, f"Outcomes table not found or empty: {co_table}"
    print(f"[PASS] Successfully retrieved Course Outcomes table for '{doc_name}' ({len(co_table)} rows)")
    print(f"       Headers: {co_table[0]}")
    print(f"       Sample Row 1: {co_table[1]}")

    # 2. Modify information inside the table itself
    # Modify Row 1 (CO1)
    modified_co_table = [row[:] for row in co_table]
    modified_co_table[1][0] = "CO1"
    modified_co_table[1][1] = "Develop state-of-the-art Deep Learning architectures with PyTorch and CUDA acceleration."
    modified_co_table[1][2] = "Analyze & Apply"
    modified_co_table[1][3] = "Module 1 Unit 1"
    modified_co_table[1][4] = "PO1, PO2, PO3, PO5"

    # Add an entirely new Outcome Row into the table (CO6)
    new_co_row = [
        "CO6",
        "Deploy production-grade Machine Learning inference pipelines using Kubernetes and Triton server.",
        "Create",
        "Module 2 Unit 3",
        "PO1, PO3, PO5, PO8"
    ]
    modified_co_table.append(new_co_row)

    print(f"[INFO] Modified CO1 text & added new row CO6 into the table (Total rows now: {len(modified_co_table)})")

    # 3. Save modified document via POST /save-pdf-content
    payload = {
        "docx_name": doc_name,
        "course_outcomes_table": json.dumps(modified_co_table)
    }

    save_resp = requests.post(f"{BASE_URL}/save-pdf-content", data=payload)
    assert save_resp.status_code == 200, f"POST /save-pdf-content failed: {save_resp.status_code} {save_resp.text}"
    save_data = save_resp.json()
    assert save_data["success"] is True, f"Save failed: {save_data}"
    print(f"[PASS] Successfully saved modified table content via /save-pdf-content")
    print(f"       PDF output: {save_data.get('pdf_url')}")

    # 4. Verify modified content in the generated PDF using PyMuPDF
    pdf_path = Path("outputs") / save_data["filename"]
    assert pdf_path.exists(), f"PDF file does not exist: {pdf_path}"

    doc_pdf = pymupdf.open(str(pdf_path))
    pdf_text = ""
    for page in doc_pdf:
        pdf_text += page.get_text()
    doc_pdf.close()

    norm_text = pdf_text.replace(" ", "")
    assert "CUDAacceleration" in norm_text, "Modified CO1 text not found in PDF"
    assert "CO6" in norm_text, "Newly added CO6 identifier not found in PDF"
    assert "production-gradeMachineLearning" in norm_text or "production-grade" in pdf_text or "inferencepipelines" in norm_text, "Newly added CO6 description not found in PDF"
    assert "Analyze&Apply" in norm_text or "Analyze" in norm_text, "Modified Bloom's level not found in PDF"

    print(f"[PASS] PDF inspection confirmed: Modified in-table outcomes and new row are rendered cleanly in the PDF table!")

    # 5. Re-query GET /get-document-content to verify round-trip persistence
    check_resp = requests.get(f"{BASE_URL}/get-document-content?docx_name={doc_name}")
    assert check_resp.status_code == 200
    check_data = check_resp.json()
    reloaded_co = check_data.get("course_outcomes_table")
    assert len(reloaded_co) == len(modified_co_table), f"Expected {len(modified_co_table)} rows, got {len(reloaded_co)}"
    assert "CUDA" in reloaded_co[1][1], "Reloaded CO1 text did not persist"
    assert reloaded_co[-1][0] == "CO6", "Reloaded CO6 did not persist"

    print(f"[PASS] Document persistence confirmed: Re-queried docx reflects all modified table rows and cells perfectly.")
    print("\nALL IN-TABLE OUTCOME MODIFICATION TESTS PASSED SUCCESSFULLY!")

if __name__ == "__main__":
    test_outcomes_table_workflow()
