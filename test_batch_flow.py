import requests
import json
import pymupdf

def run_test():
    url = 'http://127.0.0.1:8000/convert'
    files = [
        ('files', ('Vibe_Coding_Curriculum.docx', open('uploads/Vibe_Coding_Curriculum.docx', 'rb'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')),
        ('files', ('Test_Curriculum.docx', open('uploads/Test_Curriculum.docx', 'rb'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'))
    ]
    data = {
        'start_page': '80',
        'continuous_pages': 'true',
        'header_text': 'CSE - II Year II Semester'
    }

    print("Sending POST /convert with 2 files...")
    resp = requests.post(url, files=files, data=data)
    res_json = resp.json()
    print('Convert response status:', resp.status_code)
    print('Total files converted:', res_json.get('total_files'))
    print('Combined PDF URL:', res_json.get('combined_pdf_url'))
    for d in res_json.get('documents', []):
        print(f"  - {d['docx_name']}: Pages {d['start_page']} to {d['end_page']} ({d['page_count']} pages)")

    # Verify Combined PDF with PyMuPDF
    combined_path = 'outputs/Combined_Curriculum_VFSTR_Formatted.pdf'
    doc = pymupdf.open(combined_path)
    print(f"\nVerifying Combined Curriculum Book: {combined_path}")
    print(f"Total pages in combined book: {len(doc)}")
    
    # Inspect Page 1 (Subject 1)
    p1_text = doc[0].get_text()
    print("Page 1 (Doc 1) snippet:")
    print("  " + "\n  ".join([line for line in p1_text.splitlines() if line.strip()][:4]))
    
    # Inspect Page 5 (First page of Subject 2)
    p5_text = doc[4].get_text()
    print("Page 5 (Doc 2 fresh page start) snippet:")
    print("  " + "\n  ".join([line for line in p5_text.splitlines() if line.strip()][:4]))
    doc.close()

    # Test Modifier Agent
    print("\nTesting AI Modifier Agent: Changing Doc 2 photo to Robotics Automation...")
    mod_url = 'http://127.0.0.1:8000/modify'
    mod_data = {
        'docx_name': 'Test_Curriculum.docx',
        'instruction': 'Change subject photo to Robotics Automation'
    }
    mod_resp = requests.post(mod_url, data=mod_data)
    mod_json = mod_resp.json()
    print('Modifier status:', mod_resp.status_code)
    print('Modifier message:', mod_json.get('message'))
    print('Modifier combined PDF URL:', mod_json.get('combined_pdf_url'))
    print('Modifier active settings:', mod_json.get('active_settings'))

if __name__ == '__main__':
    run_test()
