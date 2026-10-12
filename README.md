# VFSTR Academic Document Formatting & Modifier Agent

An automated, precision document-formatting system that converts Microsoft Word (`.docx`) curriculum and syllabus documents into publication-ready academic PDFs following the **VFSTR CSE II Year II Semester** design standard with 100% faithful content preservation.

---

## 🌟 Key Features

1. **Page 1 (Odd Page) - Left Gray Sidebar & Subject Photo**:
   - Full-height light gray sidebar (`#E6E7E8`, `x = 0 -> 144 pt`).
   - Relevant high-resolution subject photo matching the course topic.
   - **Local Photo File Picker**: Select and upload custom images (`.png`, `.jpg`, `.jpeg`, `.webp`) directly from your local folder.
   - **Custom Source Citation Link**: Paste any custom URL to display formatted and word-wrapped beneath the photo on Page 1.
   - Real-world internet photo search via Wikipedia API, Wikimedia Commons, Unsplash API, and curated fallbacks.

2. **Page 2 (Even Page) - Alternating Right Gray Sidebar & SKILLS**:
   - Full-height light gray sidebar (`#E6E7E8`, `x = 468 -> 612 pt`).
   - `SKILLS:` heading with crisp vector checkmark items (`✓`).

3. **Multi-Document Batch Processing & Combined Curriculum Book**:
   - Select or drag & drop multiple `.docx` files simultaneously.
   - Converts each subject into its individual PDF starting fresh on an odd page.
   - Automatically merges all subjects into a single `Combined_Curriculum_VFSTR_Formatted.pdf`.
   - **Subject Re-Ordering**: Interactively arrange the sequence of subjects in the final curriculum book using ▲ Up / ▼ Down controls. Continuous page numbers dynamically recalculate across the entire book.

4. **Strict Margin Alignment**:
   - All starting margins and horizontal rules perfectly align across all pages:
     - **Odd Pages**: Content frame strictly from `x = 158 pt` to `x = 576 pt` (width = 418 pt).
     - **Even Pages**: Content frame strictly from `x = 36 pt` to `x = 454 pt` (width = 418 pt).
     - Header and footer lines match the content frame width exactly.
     - All tables (CO/PO Matrix, Course Outcomes, SDGs & IKS) fit flush within the 418 pt boundary.

5. **🤖 Dedicated AI Document Modifier Agent**:
   - Interactive Modifier Agent built directly into the Web UI and CLI.
   - Modify the generated PDF using natural language commands (e.g., *"Change course code to 25CS207, set starting page to 112, and set photo to Robotics"*).
   - Instant re-generation with live PDF preview updates in the browser.

6. **100% Content Preservation**:
   - Preserves all Modules, Units, Practices, Tasks, Analysis Tasks, Matrices, Outcomes, SDGs, Textbooks, and References without inventing synthetic badge icons.

---

## 🚀 Quick Start (Web App on Localhost)

### 1. Launch Agent
Double-click `start_agent.bat` or run in terminal:
```powershell
python app.py
```
Open your browser at:
```
http://localhost:8000
```

### 2. Upload and Format
1. Select or drag & drop your `.docx` curriculum file.
2. (Optional) Set Course Code (e.g. `24CS206`), Starting Page (e.g. `80`), or Unsplash API key.
3. Click **Transform to Academic PDF**.
4. Review the PDF in the live browser preview or click **Download PDF**.

### 3. Use the AI Modifier Agent
- Use the **AI Document Modifier Agent** prompt box below the preview to make any adjustments:
  - *"Change course code to 25CS208"*
  - *"Set starting page to 112"*
  - *"Change header text to CSE - III Year I Semester"*
  - *"Change photo to Machine Learning"*
  - *"Update skills list"*
- Or click the quick-action chips or the Fine-Tuning Drawer.

---

## 💻 CLI Usage

You can also run both the Formatter and the Modifier Agent from the command line:

```powershell
# Format a document
python academic_formatter.py "uploads/My_Syllabus.docx" "outputs/Output.pdf"

# Run the Modifier Agent with natural language instruction
python academic_modifier.py --docx "uploads/My_Syllabus.docx" --instruction "Change course code to 25CS301 and start page to 112"
```

---

## 📦 Requirements

- Python 3.9+
- Windows OS (with system Arial and Times New Roman fonts)
- Dependencies installed via:
  ```powershell
  pip install -r requirements.txt
  ```
