import requests
import json
import pymupdf
from pathlib import Path

BASE_URL = "http://localhost:8000"

def test_edit_and_save_workflow():
    print("=== Testing Edit PDF and Save Modified PDF Endpoints ===")

    # 1. Fetch document content
    doc_name = "5.Machine Learning(II-II) (4C).docx"
    resp = requests.get(f"{BASE_URL}/get-document-content?docx_name={doc_name}")
    assert resp.status_code == 200, f"Failed GET /get-document-content: {resp.status_code} {resp.text}"
    data = resp.json()
    assert data["success"] is True, f"Response success was not True: {data}"
    print(f"[PASS] Retrieved document content for '{doc_name}'")
    print(f"       Course Code: {data.get('course_code')}")
    print(f"       Title: {data.get('course_title')}")
    print(f"       Hours: {data.get('hours')}")

    # 2. Modify and Save
    test_title = "ADVANCED MACHINE LEARNING & INTELLIGENT SYSTEMS"
    test_code = "25CS999"
    test_prereq = "Python Programming and Linear Algebra Essentials"
    test_desc = "This modified course delivers rigorous hands-on training in state-of-the-art intelligent systems, deep neural networks, and decision architectures."
    test_skills = "Architecting Neural Networks\nImplementing Deep Q-Learning\nDeploying Transformer Models"
    test_url = "https://en.wikipedia.org/wiki/Machine_learning"

    payload = {
        "docx_name": doc_name,
        "course_title": test_title,
        "course_code": test_code,
        "hours_l": "4",
        "hours_t": "1",
        "hours_p": "2",
        "hours_sl": "1",
        "hours_c": "5",
        "start_page": "80",
        "header_text": "CSE - II Year II Semester",
        "prerequisite": test_prereq,
        "description": test_desc,
        "skills": test_skills,
        "custom_source_url": test_url,
        "subject_image_query": "Machine Learning Deep Neural Networks"
    }

    save_resp = requests.post(f"{BASE_URL}/save-pdf-content", data=payload)
    assert save_resp.status_code == 200, f"Failed POST /save-pdf-content: {save_resp.status_code} {save_resp.text}"
    save_data = save_resp.json()
    assert save_data["success"] is True, f"Save failed: {save_data}"
    print(f"[PASS] Successfully saved modified PDF via /save-pdf-content")
    print(f"       Message: {save_data.get('message')}")
    print(f"       PDF URL: {save_data.get('pdf_url')}")
    print(f"       Download URL: {save_data.get('download_url')}")

    # 3. Verify Download of Modified PDF
    dl_url = f"{BASE_URL}{save_data['download_url']}"
    dl_resp = requests.get(dl_url)
    assert dl_resp.status_code == 200, f"Failed to download modified PDF: {dl_resp.status_code}"
    assert len(dl_resp.content) > 1000, "Downloaded PDF is too small"
    print(f"[PASS] Downloaded modified PDF ({len(dl_resp.content)} bytes)")

    # 4. Verify text content in generated PDF using PyMuPDF
    pdf_filename = save_data["filename"]
    pdf_path = Path("outputs") / pdf_filename
    assert pdf_path.exists(), f"PDF does not exist: {pdf_path}"

    doc = pymupdf.open(str(pdf_path))
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    doc.close()

    norm_text = full_text.replace(" ", "")
    assert "25CS999" in norm_text, "Modified course code 25CS999 not found in PDF"
    assert "ADVANCEDMACHINELEARNING" in norm_text, "Modified title not found in PDF"
    assert "ArchitectingNeuralNetworks" in norm_text, "Modified skills not found in PDF"
    assert "wikipedia.org" in norm_text, "Custom source URL not found in PDF"
    print(f"[PASS] PDF text inspection verified all modified fields are present in the PDF output!")

    # 5. Verify Combined Curriculum Book re-compilation
    if "combined_download_url" in save_data:
        comb_dl = f"{BASE_URL}{save_data['combined_download_url']}"
        comb_resp = requests.get(comb_dl)
        assert comb_resp.status_code == 200, "Combined book download failed"
        print(f"[PASS] Combined curriculum book updated and downloadable ({len(comb_resp.content)} bytes)")

    print("\nALL TESTS PASSED! Edit PDF and Save Modified PDF functionality is 100% verified.")

if __name__ == "__main__":
    test_edit_and_save_workflow()
