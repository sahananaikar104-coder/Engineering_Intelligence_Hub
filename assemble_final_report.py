"""
assemble_final_report.py - Master Assembly Engine for EIH Major Project Report.
Compiles Preliminaries, Chapters 1-6, and Backmatter into Final_Project_Report.docx.
Department of Artificial Intelligence and Machine Learning,
Basaveshwar Engineering College, Bagalkote.
"""

import os
import sys
import time
import docx
from docx.shared import Inches, Pt, RGBColor
from report_styles import (
    add_page_border, setup_header_footer
)
from report_prelims import build_preliminaries
from report_chapter1 import build_chapter_1
from report_chapter2 import build_chapter_2
from report_chapter3 import build_chapter_3
from report_chapter4 import build_chapter_4
from report_chapter5 import build_chapter_5
from report_chapter6 import build_chapter_6
from report_backmatter import build_backmatter

def assemble_report():
    print("=" * 70)
    print("ENGINEERING INTELLIGENCE HUB (EIH) - MAJOR PROJECT REPORT ASSEMBLER")
    print("=" * 70)
    
    print("\n[Step 1/8] Initializing docx.Document object and base typography...")
    doc = docx.Document()

    # Configure Normal Style defaults
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.font.color.rgb = RGBColor(0x00, 0x00, 0x00)

    # 1. Preliminaries (Section 1: Title Page, Section 2: Roman i-x)
    print("\n[Step 2/8] Generating Preliminary Pages (Title, Certificate, Declaration, Acknowledgement, Abstract, TOC, LOF, LOT, Abbr)...")
    build_preliminaries(doc)

    # 2. Main Body Section (Section 3: Decimal 1-N)
    print("\n[Step 3/8] Configuring Section 3 for Report Body (Arabic numeral decimal pagination)...")
    s3 = doc.add_section()
    s3.top_margin = Inches(1.0)
    s3.bottom_margin = Inches(1.0)
    s3.left_margin = Inches(1.25)
    s3.right_margin = Inches(1.0)
    add_page_border(s3)
    setup_header_footer(s3, is_prelim=False)

    # Chapter 1: Introduction
    print("\n[Step 4/8] Assembling Chapter 1: Introduction (1.1 - 1.10)...")
    build_chapter_1(doc)
    doc.add_page_break()

    # Chapter 2: Literature Survey
    print("\n[Step 5/8] Assembling Chapter 2: Literature Survey (2.1 - 2.6)...")
    build_chapter_2(doc)
    doc.add_page_break()

    # Chapter 3: Requirements and Proposed Methodology
    print("\n[Step 6/8] Assembling Chapter 3: Requirements and Proposed Methodology (3.1 - 3.10)...")
    build_chapter_3(doc)
    doc.add_page_break()

    # Chapter 4: System Design and Implementation
    print("\n[Step 7/8] Assembling Chapter 4: System Design and Implementation (4.1 - 4.10)...")
    build_chapter_4(doc)
    doc.add_page_break()

    # Chapter 5: Results and Discussion
    print("\n[Step 8/8a] Assembling Chapter 5: Results and Discussion (5.1 - 5.15)...")
    build_chapter_5(doc)
    doc.add_page_break()

    # Chapter 6: Conclusion and Future Scope
    print("\n[Step 8/8b] Assembling Chapter 6: Conclusion and Future Scope (6.1 - 6.7)...")
    build_chapter_6(doc)
    doc.add_page_break()

    # Backmatter: References & Appendices A to F
    print("\n[Step 8/8c] Assembling Backmatter: References & Appendices A through F...")
    build_backmatter(doc)

    # Save destinations
    out_dir_1 = r"c:\Users\sahan\Downloads\github_analyzer"
    out_dir_2 = r"c:\Users\sahan\Downloads\github_analyzer\github_analyzer"

    docx_path_1 = os.path.join(out_dir_1, "Final_Project_Report.docx")
    docx_path_2 = os.path.join(out_dir_2, "Final_Project_Report.docx")

    print(f"\nSaving master document to: {docx_path_1}")
    doc.save(docx_path_1)

    print(f"Saving duplicate copy to: {docx_path_2}")
    doc.save(docx_path_2)

    print("\n" + "=" * 70)
    print("REPORT ASSEMBLY COMPLETED SUCCESSFULLY!")
    print(f"Target File 1: {docx_path_1} (Size: {os.path.getsize(docx_path_1):,} bytes)")
    print(f"Target File 2: {docx_path_2} (Size: {os.path.getsize(docx_path_2):,} bytes)")
    print("=" * 70)
    return docx_path_1

if __name__ == "__main__":
    t0 = time.time()
    docx_file = assemble_report()
    print(f"\nExecution finished in {time.time() - t0:.2f} seconds.")
