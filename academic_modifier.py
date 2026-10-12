"""
AI Academic Document Modifier Agent
Interactive assistant that interprets natural language instructions or parameter overrides
to modify, customize, and regenerate the VFSTR academic curriculum PDF dynamically.
"""

import os
import re
import sys
import json
from pathlib import Path
from typing import Dict, Any, Optional, List, Tuple

from academic_formatter import convert_docx_to_pdf


class AcademicModifierAgent:
    """Interprets modification instructions and regenerates the PDF with updated properties."""

    def __init__(self, docx_path: str, output_pdf_path: Optional[str] = None):
        self.docx_path = docx_path
        self.base_name = Path(docx_path).stem
        self.output_pdf_path = output_pdf_path or f"outputs/{self.base_name}_VFSTR_Formatted.pdf"

        # Default document state
        self.current_settings: Dict[str, Any] = {
            "start_page": 80,
            "header_text": "CSE - II Year II Semester",
            "course_code": "24CS206",
            "subject_image_query": None,
            "custom_image_url": None,
            "custom_source_url": None,
            "unsplash_api_key": None,
            "skills": [
                "Develop functional web and mobile applications using modern AI coding assistants.",
                "Design effective prompt engineering workflows for architecture and UI generation.",
                "Integrate frontend components with backend databases, APIs, and cloud services.",
                "Debug, test, and safely deploy AI-generated application code."
            ]
        }

    def parse_instruction(self, instruction: str) -> Tuple[Dict[str, Any], List[str]]:
        """Parses natural language requests into concrete document settings changes."""
        changes = {}
        explanations = []

        if not instruction:
            return changes, explanations

        inst = instruction.strip()

        # 1. Course Code (e.g. "change course code to 25CS207", "set code 24CS102")
        code_match = re.search(r'\b(?:course\s*code|code)\s*(?:to|is|=)?\s*([0-9]{2}[A-Z]{2,4}[0-9]{3,4}[A-Z]?)\b', inst, re.I)
        if not code_match:
            code_match = re.search(r'\b([0-9]{2}[A-Z]{2,4}[0-9]{3,4}[A-Z]?)\b', inst, re.I)
        if code_match:
            new_code = code_match.group(1).upper()
            changes['course_code'] = new_code
            explanations.append(f"Updated course code to '{new_code}'")

        # 2. Starting Page Number (e.g. "change start page to 112", "page 90", "start at 112")
        page_match = re.search(r'\b(?:start(?:ing)?\s*page(?:\s*number)?|page\s*number|start\s*at)\s*(?:to|is|=)?\s*(\d+)\b', inst, re.I)
        if page_match:
            new_page = int(page_match.group(1))
            changes['start_page'] = new_page
            explanations.append(f"Set starting page number to {new_page}")

        # 3. Header Text (e.g. "change header text to CSE - III Year I Semester", "header: CSE ...")
        header_match = re.search(r'\b(?:header(?:\s*text)?)\s*(?:to|is|=)?\s*[\'"]?([A-Za-z0-9\s\-–—]+(?:Semester|Year|CSE|IT|ECE|EEE)[A-Za-z0-9\s\-–—]*)[\'"]?', inst, re.I)
        if header_match:
            new_header = header_match.group(1).strip()
            changes['header_text'] = new_header
            explanations.append(f"Changed running header to '{new_header}'")

        # 4. Unsplash API Key (e.g. "use unsplash key abc123xyz", "unsplash api key: ...")
        unsplash_match = re.search(r'\b(?:unsplash(?:\s*api)?\s*(?:key|token))\s*(?:to|is|=)?\s*([A-Za-z0-9_\-]+)\b', inst, re.I)
        if unsplash_match:
            key = unsplash_match.group(1).strip()
            changes['unsplash_api_key'] = key
            explanations.append("Applied custom Unsplash API key for real-world image retrieval")

        # 5. Custom Photo URL (e.g. "set photo url https://...", "use image https://...")
        url_match = re.search(r'\b(?:image|photo|picture)\s*(?:url|link)\s*(?:to|is|=)?\s*(https?://[^\s]+)', inst, re.I)
        if url_match:
            img_url = url_match.group(1).strip()
            changes['custom_image_url'] = img_url
            explanations.append(f"Configured custom subject image URL: {img_url[:40]}...")

        # 5b. Source Citation URL (e.g. "source https://...", "citation url https://...")
        src_match = re.search(r'\b(?:source|attribution|citation)\s*(?:url|link)?\s*(?:to|is|=)?\s*(https?://[^\s]+)', inst, re.I)
        if src_match:
            src_url = src_match.group(1).strip()
            changes['custom_source_url'] = src_url
            explanations.append(f"Set photo source citation link to '{src_url[:40]}...'")

        # 6. Subject Image Query (e.g. "change photo to robotics", "search image for machine learning", "subject topic: ...")
        topic_match = re.search(r'\b(?:photo|image|picture)\s*(?:of|for|about|to)\s*[\'"]?([a-zA-Z\s]{3,35})[\'"]?', inst, re.I)
        if topic_match and 'url' not in topic_match.group(1).lower():
            query = topic_match.group(1).strip()
            changes['subject_image_query'] = query
            explanations.append(f"Set subject image search topic to '{query}'")

        # 7. Skills list (e.g. "skills: 1. Python 2. React...", "update skills to: ...")
        if 'skill' in inst.lower() and (':' in inst or '\n' in inst or ';' in inst):
            after_skills = re.split(r'\bskills?\s*[:=]\s*', inst, flags=re.I)[-1]
            raw_items = [s.strip() for s in re.split(r'[\n;]|(?:\d+\.|\b-\s*|\b•\s*)', after_skills) if len(s.strip()) > 5]
            if raw_items:
                changes['skills'] = raw_items
                explanations.append(f"Updated Right-Sidebar SKILLS with {len(raw_items)} custom items")

        return changes, explanations

    def modify(
        self,
        instruction: Optional[str] = None,
        overrides: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """Applies modifications and regenerates the PDF document."""
        explanations = []
        parsed_changes = {}

        # 1. Parse natural language instructions
        if instruction:
            parsed_changes, inst_explanations = self.parse_instruction(instruction)
            self.current_settings.update(parsed_changes)
            explanations.extend(inst_explanations)

        # 2. Apply explicit parameter overrides
        if overrides:
            for k, v in overrides.items():
                if v is not None and v != "":
                    if k == "start_page":
                        self.current_settings["start_page"] = int(v)
                        explanations.append(f"Set starting page to {v}")
                    elif k == "header_text":
                        self.current_settings["header_text"] = str(v)
                        explanations.append(f"Set header to '{v}'")
                    elif k == "course_code":
                        self.current_settings["course_code"] = str(v)
                        explanations.append(f"Set course code to '{v}'")
                    elif k == "subject_image_query":
                        self.current_settings["subject_image_query"] = str(v)
                        explanations.append(f"Set image query to '{v}'")
                    elif k == "custom_image_path":
                        self.current_settings["custom_image_path"] = str(v)
                        explanations.append("Installed custom photo from folder")
                    elif k == "custom_image_url":
                        self.current_settings["custom_image_url"] = str(v)
                        explanations.append("Applied custom image URL")
                    elif k == "custom_source_url":
                        self.current_settings["custom_source_url"] = str(v)
                        explanations.append(f"Set photo source citation URL to '{str(v)[:40]}...'")
                    elif k == "unsplash_api_key":
                        self.current_settings["unsplash_api_key"] = str(v)
                        explanations.append("Applied Unsplash API Key")
                    elif k == "skills":
                        if isinstance(v, list):
                            self.current_settings["skills"] = v
                        elif isinstance(v, str):
                            items = [s.strip() for s in v.split("\n") if s.strip()]
                            self.current_settings["skills"] = items
                        explanations.append("Updated SKILLS list")

        if not explanations:
            explanations.append("Re-compiled document with current active formatting specifications")

        # Determine if photo was modified or requested to change
        force_refresh_photo = any(
            k in (instruction or "").lower() for k in ["photo", "image", "picture"]
        ) or any(
            k in parsed_changes for k in ["subject_image_query", "custom_image_url", "custom_image_path", "unsplash_api_key"]
        ) or any(
            k in (overrides or {}) for k in ["subject_image_query", "custom_image_url", "custom_image_path", "unsplash_api_key"]
        )

        # 3. Regenerate the PDF using the updated settings
        out_pdf = convert_docx_to_pdf(
            docx_path=self.docx_path,
            output_pdf_path=self.output_pdf_path,
            start_page=self.current_settings["start_page"],
            header_text=self.current_settings["header_text"],
            course_code=self.current_settings["course_code"],
            subject_image_query=self.current_settings["subject_image_query"],
            custom_image_path=self.current_settings.get("custom_image_path"),
            custom_image_url=self.current_settings["custom_image_url"],
            custom_source_url=self.current_settings["custom_source_url"],
            unsplash_api_key=self.current_settings["unsplash_api_key"],
            custom_skills=self.current_settings["skills"],
            custom_sdgs=self.current_settings.get("custom_sdgs"),
            force_refresh=force_refresh_photo
        )

        filename = Path(out_pdf).name

        return {
            "success": True,
            "message": " &bull; ".join(explanations),
            "filename": filename,
            "pdf_url": f"/view-pdf/{filename}",
            "download_url": f"/download-pdf/{filename}",
            "active_settings": self.current_settings
        }


# ==============================================================================
# CLI ENTRY POINT
# ==============================================================================
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="AI Academic Document Modifier Agent")
    parser.add_argument("--docx", type=str, required=True, help="Path to input Word document (.docx)")
    parser.add_argument("--instruction", type=str, default="", help="Natural language modification prompt")
    parser.add_argument("--course-code", type=str, help="Course code (e.g. 24CS206)")
    parser.add_argument("--start-page", type=int, help="Starting page number (e.g. 80)")
    parser.add_argument("--header", type=str, help="Running header text")
    parser.add_argument("--image-topic", type=str, help="Subject photo search keyword")
    parser.add_argument("--unsplash-key", type=str, help="Unsplash API access key")

    args = parser.parse_args()

    agent = AcademicModifierAgent(docx_path=args.docx)
    overrides = {}
    if args.course_code:
        overrides["course_code"] = args.course_code
    if args.start_page:
        overrides["start_page"] = args.start_page
    if args.header:
        overrides["header_text"] = args.header
    if args.image_topic:
        overrides["subject_image_query"] = args.image_topic
    if args.unsplash_key:
        overrides["unsplash_api_key"] = args.unsplash_key

    result = agent.modify(instruction=args.instruction, overrides=overrides)
    print("\n[Modifier Agent Result]")
    print(json.dumps(result, indent=2))
