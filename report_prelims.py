import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from report_styles import (
    add_page_border, setup_header_footer, add_p, add_heading_chapter,
    style_table
)

def build_preliminaries(doc):
    # ----------------------------------------------------
    # SECTION 1: TITLE PAGE
    # ----------------------------------------------------
    s1 = doc.sections[0]
    s1.top_margin = Inches(1.0)
    s1.bottom_margin = Inches(1.0)
    s1.left_margin = Inches(1.25)
    s1.right_margin = Inches(1.0)
    s1.different_first_page_header_footer = True
    add_page_border(s1)
    
    # Institution
    add_p(doc, "Basaveshwar Engineering College, Bagalkote", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=2, bold=True, font_size=15)
    add_p(doc, "[An Autonomous Government Aided College, AICTE approved, Affiliated to VTU, Belagavi]", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=4, italic=True, font_size=11)
    add_p(doc, "Department of Artificial Intelligence and Machine Learning", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=2, space_after=22, bold=True, font_size=13.5)
    
    add_p(doc, "MAJOR PROJECT REPORT (22UAI706P)", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=6, bold=True, font_size=13)
    add_p(doc, "on", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8, italic=True, font_size=12)
    
    add_p(doc, "ENGINEERING INTELLIGENCE HUB: An Autonomous Agentic Platform for Repository-Aware Software Engineering", 
          align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=22, bold=True, font_size=15.5)
    
    add_p(doc, "Submitted by", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=6, italic=True, font_size=12)
    
    # Students Table
    t_team = doc.add_table(rows=5, cols=3)
    headers = ["Sl. No.", "Name", "USN"]
    for j, h in enumerate(headers):
        t_team.rows[0].cells[j].text = h
        t_team.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    data_team = [
        ("01", "Chinmayi S Mullur", "2BA23AI014"),
        ("02", "Ranjita M Julpi", "2BA23AI037"),
        ("03", "Sahana R Naikar", "2BA23AI038"),
        ("04", "Sneha S Patil", "2BA23AI053")
    ]
    for i, row in enumerate(data_team):
        for j, val in enumerate(row):
            cell = t_team.rows[i+1].cells[j]
            cell.text = val
            if j == 0:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif j == 2:
                cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_table(t_team, col_widths=[Inches(1.0), Inches(3.2), Inches(1.8)])
    
    add_p(doc, "", space_after=14)
    add_p(doc, "Under the guidance of", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=3, italic=True, font_size=11.5)
    add_p(doc, "Dr. Vishwanath Kagawade", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2, bold=True, font_size=13.5)
    add_p(doc, "Assistant Professor, Department of AIML", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=18, font_size=11)
    
    add_p(doc, "In partial fulfilment of the requirements of the seventh semester\nBachelor of Engineering in Artificial Intelligence and Machine Learning", 
          align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=16, font_size=11.5)
    
    add_p(doc, "Academic Year 2026-27\nNovember 2026", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, bold=True, font_size=12)

    # ----------------------------------------------------
    # SECTION 2: PRELIMINARY PAGES (ROMAN NUMERALS i to x)
    # ----------------------------------------------------
    s2 = doc.add_section()
    s2.top_margin = Inches(1.0)
    s2.bottom_margin = Inches(1.0)
    s2.left_margin = Inches(1.25)
    s2.right_margin = Inches(1.0)
    add_page_border(s2)
    setup_header_footer(s2, is_prelim=True)
    
    # 1. CERTIFICATE (Page i)
    add_heading_chapter(doc, "CERTIFICATE")
    add_p(doc, "Department of Artificial Intelligence and Machine Learning\nBasaveshwar Engineering College, Bagalkote", 
          align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, space_after=18, font_size=12.5)
    
    cert_text = (
        "This is to certify that the Major Project work (22UAI706P) entitled “ENGINEERING INTELLIGENCE HUB: "
        "An Autonomous Agentic Platform for Repository-Aware Software Engineering” is a bonafide work carried out "
        "by Chinmayi S Mullur (2BA23AI014), Ranjita M Julpi (2BA23AI037), Sahana R Naikar (2BA23AI038), "
        "Sneha S Patil (2BA23AI053), in partial fulfilment of the requirements for the seventh semester of the "
        "Bachelor of Engineering programme in Artificial Intelligence and Machine Learning at Basaveshwar "
        "Engineering College, Bagalkote, affiliated to Visvesvaraya Technological University, Belagavi, during "
        "the academic year 2026-27.\n\n"
        "The report has been approved as it satisfies the academic requirements in respect of the Major Project "
        "prescribed for the programme."
    )
    add_p(doc, cert_text, space_after=40, line_spacing=1.25)
    
    # Signatures
    t_sig = doc.add_table(rows=2, cols=4)
    sigs = [
        ("(Dr. Vishwanath Kagawade)\nProject Guide", "(Prof. J. D. Kallaganiger)\nProject Co-ordinator", 
         "(Dr. B. M. Reshmi)\nHead of Department", "(                       )\nPrincipal")
    ]
    for j in range(4):
        t_sig.rows[0].cells[j].text = sigs[0][j]
        t_sig.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        for p in t_sig.rows[0].cells[j].paragraphs:
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(9.5)
                r.font.bold = True
    style_table(t_sig, header_bg=None, col_widths=[Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)])
    
    add_p(doc, "", space_after=30)
    add_p(doc, "Name of the examiners:", bold=True, space_after=8, font_size=11)
    add_p(doc, "1. __________________________________        Signature: __________________________", space_after=10, font_size=11)
    add_p(doc, "2. __________________________________        Signature: __________________________", space_after=0, font_size=11)
    
    # 2. DECLARATION (Page ii)
    doc.add_page_break()
    add_heading_chapter(doc, "DECLARATION")
    decl_text1 = (
        "We, the undersigned students of the seventh semester of the Bachelor of Engineering programme in "
        "Artificial Intelligence and Machine Learning at Basaveshwar Engineering College, Bagalkote, hereby "
        "declare that the project work entitled “ENGINEERING INTELLIGENCE HUB: An Autonomous Agentic Platform "
        "for Repository-Aware Software Engineering” is an original work carried out by us under the guidance of "
        "Dr. Vishwanath Kagawade, Assistant Professor, Department of AIML, in accordance with the academic "
        "regulations of the institution."
    )
    decl_text2 = (
        "We further declare that this work has not been submitted, in part or in full, to any other university or "
        "institution for the award of any degree or diploma. The experimental data, benchmarks, architectural models, "
        "source code implementations, and application screenshots presented in this report were obtained from the "
        "development and empirical evaluation described herein, and all secondary sources referenced have been "
        "duly acknowledged."
    )
    add_p(doc, decl_text1, space_after=10, line_spacing=1.22)
    add_p(doc, decl_text2, space_after=20, line_spacing=1.22)
    
    t_dec = doc.add_table(rows=5, cols=3)
    dec_headers = ["Name", "USN", "Signature"]
    for j, h in enumerate(dec_headers):
        t_dec.rows[0].cells[j].text = h
        t_dec.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for i, row in enumerate(data_team):
        t_dec.rows[i+1].cells[0].text = row[1]
        t_dec.rows[i+1].cells[1].text = row[2]
        t_dec.rows[i+1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        t_dec.rows[i+1].cells[2].text = "" # Signature space
    style_table(t_dec, col_widths=[Inches(2.5), Inches(1.8), Inches(1.7)])
    
    add_p(doc, "", space_after=25)
    add_p(doc, "Place: Bagalkote", space_after=4, font_size=11)
    add_p(doc, "Date: November 2026", space_after=0, font_size=11)
    
    # 3. ACKNOWLEDGEMENT (Page iii)
    doc.add_page_break()
    add_heading_chapter(doc, "ACKNOWLEDGEMENT")
    ack1 = (
        "We express our profound gratitude to Basaveshwar Engineering College, Bagalkote, and to our respected "
        "Principal for providing the academic infrastructure, computing facilities, and conducive research environment "
        "necessary to bring this major project to successful completion."
    )
    ack2 = (
        "We convey our sincere and heartfelt thanks to Dr. B. M. Reshmi, Professor and Head of the Department of "
        "Artificial Intelligence and Machine Learning, for her constant encouragement, strategic guidance, and for "
        "making the department's technical laboratories and computing servers readily available."
    )
    ack3 = (
        "We are deeply indebted to our esteemed project guide, Dr. Vishwanath Kagawade, Assistant Professor, Department "
        "of AIML, for his invaluable technical insights, patient mentorship, and constructive critiques throughout the "
        "conception, implementation, and evaluation of this project. We also express our sincere appreciation to "
        "Prof. J. D. Kallaganiger, Project Co-ordinator, for his diligent reviews, structured project milestones, and "
        "continuous coordination."
    )
    ack4 = (
        "We extend our gratitude to all the teaching and non-teaching faculty members of the Department of AIML for "
        "their technical guidance and assistance. We also acknowledge the global open-source software and research "
        "communities behind FastAPI, React, Vite, ChromaDB, Sentence-Transformers, and LangChain, whose open tools and "
        "rigorous documentation formed the technical foundation of our platform. Finally, we thank our parents, "
        "family members, and fellow students for their unwavering moral support, patience, and encouragement."
    )
    add_p(doc, ack1, space_after=10, line_spacing=1.22)
    add_p(doc, ack2, space_after=10, line_spacing=1.22)
    add_p(doc, ack3, space_after=10, line_spacing=1.22)
    add_p(doc, ack4, space_after=25, line_spacing=1.22)
    
    add_p(doc, "Chinmayi S Mullur (2BA23AI014)", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2, font_size=11)
    add_p(doc, "Ranjita M Julpi (2BA23AI037)", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2, font_size=11)
    add_p(doc, "Sahana R Naikar (2BA23AI038)", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=2, font_size=11)
    add_p(doc, "Sneha S Patil (2BA23AI053)", align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=0, font_size=11)

    # 4. ABSTRACT (Page iv)
    doc.add_page_break()
    add_heading_chapter(doc, "ABSTRACT")
    abs1 = (
        "Modern software repositories are large, multi-layered digital ecosystems consisting of tens of thousands of lines "
        "of source code, nested architectural boundaries, application programming interfaces (APIs), database entities, "
        "and evolving manifest configurations. While large language model (LLM) coding assistants have achieved widespread "
        "adoption, existing tools typically operate on isolated code snippets and lack holistic context awareness across the "
        "entire codebase. Furthermore, monolithic code generation paradigms frequently generate unverified implementations, "
        "suffer from severe hallucination in repository-level question answering, overwrite active workspaces indiscriminately, "
        "and provide no automated validation or human-in-the-loop review gating. This project presents the Engineering "
        "Intelligence Hub (EIH), an autonomous, agentic software engineering platform engineered to ingest, comprehend, "
        "modify, test, and document complex GitHub repositories with line-level ground truth and verifiable citations."
    )
    abs2 = (
        "The architecture introduces three coordinated execution paths. First, an AST Code Intelligence and Ingestion pipeline "
        "safely clones repositories, filters sensitive credentials, and parses Abstract Syntax Trees to extract classes, functions, "
        "decorators, docstrings, and line bounds into a relational symbol catalog alongside 384-dimensional dense ChromaDB vector "
        "embeddings. Second, a Grounded RAG Chat engine couples hybrid vector retrieval with exact line-range citation linking "
        "(file_path:start_line-end_line) to eliminate hallucinated answers and provide interactive source inspection. Third, an "
        "Autonomous Multi-Agent State Machine coordinates 12 specialized agents—Orchestrator, Requirement, Repository Analysis, "
        "Code Understanding, System Design, Code Generation, Test Generation, Documentation, Diagram Studio, Validation, Review, "
        "and Change Impact—to plan, synthesize, and verify engineering tasks."
    )
    abs3 = (
        "To ensure absolute codebase integrity, the platform enforces two novel governance safeguards: an automated Self-Correction "
        "Diagnostic Loop that captures AST syntax errors or sandbox test failures and iteratively re-prompts the code generator up "
        "to a configured budget of three retries; and a Human-in-the-Loop Change Staging Workspace where all proposed additions and "
        "modifications are presented in side-by-side and unified diff viewers, completely preventing repository mutations until "
        "explicit developer approval is granted."
    )
    abs4 = (
        "The complete system is implemented in Python 3.11 with FastAPI, React 18 with TypeScript, Vite, Tailwind CSS, SQLite, "
        "and ChromaDB, and configured for dual cloud deployment on Render and Vercel. Comprehensive empirical evaluation on five "
        "distinct multi-tier software repositories verified all 140 automated tests passing in 22.79 seconds (100% test pass rate). "
        "Comparative benchmark analysis demonstrated that the multi-agent workflow increases validation pass rates from 88.5% "
        "(single-agent baseline) to 97.2%, enhances code quality scores from 85.0 to 96.0/100, elevates security scores from 88.0 "
        "to 98.0/100, and achieves an 89.5% self-correction recovery rate, establishing a dependable, scalable paradigm for autonomous "
        "AI-assisted software engineering."
    )
    add_p(doc, abs1, space_after=8, line_spacing=1.2)
    add_p(doc, abs2, space_after=8, line_spacing=1.2)
    add_p(doc, abs3, space_after=8, line_spacing=1.2)
    add_p(doc, abs4, space_after=14, line_spacing=1.2)
    
    add_p(doc, "Keywords: repository understanding, code intelligence, Abstract Syntax Tree, multi-agent systems, "
               "Retrieval-Augmented Generation, self-correction loop, human-in-the-loop staging, verifiable citations.", 
          bold=True, italic=True, space_after=0, font_size=10.5)

    # 5. TABLE OF CONTENTS (Pages v - vi)
    doc.add_page_break()
    add_heading_chapter(doc, "TABLE OF CONTENTS")
    
    toc_data = [
        ("CHAPTER 1 – INTRODUCTION", "1", True),
        ("1.1 Background", "1", False),
        ("1.2 Problem Statement", "2", False),
        ("1.3 Motivation", "2", False),
        ("1.4 Objectives", "3", False),
        ("1.5 Scope of the Project", "4", False),
        ("1.6 Innovation and Key Contributions", "5", False),
        ("1.7 Social and Industrial Relevance", "5", False),
        ("1.8 Government and National Priority Alignment", "6", False),
        ("1.9 SDG Alignment", "6", False),
        ("1.10 Organization of the Report", "6", False),
        ("CHAPTER 2 – LITERATURE SURVEY", "8", True),
        ("2.1 Introduction", "8", False),
        ("2.2 Review of Existing Systems and Methods", "8", False),
        ("2.3 Comparison of Existing Approaches", "10", False),
        ("2.4 Research and Technology Gap", "11", False),
        ("2.5 Limitations of Existing Methods", "12", False),
        ("2.6 Summary", "12", False),
        ("CHAPTER 3 – REQUIREMENTS AND PROPOSED METHODOLOGY", "13", True),
        ("3.1 Introduction", "13", False),
        ("3.2 Requirement Collection", "13", False),
        ("3.3 Requirement Analysis", "18", False),
        ("3.4 Proposed Solution", "19", False),
        ("3.5 Proposed Methodology", "19", False),
        ("3.6 Proposed System Architecture", "20", False),
        ("3.7 Algorithms and Techniques", "21", False),
        ("3.8 Project Planning", "26", False),
        ("3.9 Requirement Traceability Matrix", "27", False),
        ("3.10 Summary", "29", False),
        ("CHAPTER 4 – SYSTEM DESIGN AND IMPLEMENTATION", "30", True),
        ("4.1 Introduction", "30", False),
        ("4.2 Detailed System Design", "30", False),
        ("4.3 System Modelling", "31", False),
        ("4.4 Module Design", "39", False),
        ("4.5 Dataset and Data Preparation", "40", False),
        ("4.6 Model Development and Algorithm Design", "42", False),
        ("4.7 System Implementation", "43", False),
        ("4.8 User Interface / Application", "44", False),
        ("4.9 Deployment", "50", False),
        ("4.10 Summary", "51", False),
        ("CHAPTER 5 – TESTING, RESULTS AND DISCUSSION", "52", True),
        ("5.1 Experimental Setup", "52", False),
        ("5.2 Test Plan", "53", False),
        ("5.3 Test Cases", "53", False),
        ("5.4 Functional Testing", "55", False),
        ("5.5 Performance Testing", "56", False),
        ("5.6 Evaluation Metrics", "57", False),
        ("5.7 Experimental Results", "58", False),
        ("5.8 Comparative Analysis", "62", False),
        ("5.9 Results – Graphs and Tables", "62", False),
        ("5.10 Screenshots / Outputs", "63", False),
        ("5.11 Error Analysis", "64", False),
        ("5.12 Requirement–Implementation–Testing Traceability", "65", False),
        ("5.13 Discussion of Results", "66", False),
        ("5.14 Reproducibility Details", "67", False),
        ("5.15 Summary", "67", False),
        ("CHAPTER 6 – CONCLUSION AND FUTURE SCOPE", "68", True),
        ("6.1 Conclusion", "68", False),
        ("6.2 Major Findings", "68", False),
        ("6.3 Achievement of Objectives", "68", False),
        ("6.4 Limitations", "69", False),
        ("6.5 Future Enhancements", "70", False),
        ("6.6 Sustainability / Social Impact", "70", False),
        ("6.7 Scope for Further Research", "70", False),
        ("REFERENCES", "71", True),
        ("APPENDIX A – IMPORTANT SOURCE CODE", "73", True),
        ("APPENDIX B – ADDITIONAL RESULTS", "77", True),
        ("APPENDIX C – QUESTIONNAIRES / DATA COLLECTION FORMS", "80", True),
        ("APPENDIX D – USER MANUAL", "82", True),
        ("APPENDIX E – ADDITIONAL SCREENSHOTS", "84", True),
        ("APPENDIX F – INDIVIDUAL STUDENT CONTRIBUTIONS", "90", True)
    ]
    
    for title, page, is_bold in toc_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2 if is_bold else 0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.02), WD_TAB_ALIGNMENT.RIGHT)
        
        r_t = p.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(10.0 if not is_bold else 10.5)
        r_t.font.bold = is_bold
        
        # Leader dots calculation
        dots_len = max(2, int((6.0 - len(title)*0.08) * 12))
        r_dots = p.add_run(" " + "." * dots_len + "\t" + page)
        r_dots.font.name = 'Times New Roman'
        r_dots.font.size = Pt(10.0 if not is_bold else 10.5)
        r_dots.font.bold = is_bold

    # 6. LIST OF FIGURES (Pages vii - viii)
    doc.add_page_break()
    add_heading_chapter(doc, "LIST OF FIGURES")
    
    fig_list = [
        ("Figure 3.1: Proposed methodology: three processing paths from ingestion to consumer", "19"),
        ("Figure 3.2: High-level system architecture of Engineering Intelligence Hub (EIH)", "20"),
        ("Figure 3.3: Task orchestration, multi-agent state machine, and self-correction loop", "25"),
        ("Figure 3.4: Work breakdown structure (WBS) of the EIH platform project", "26"),
        ("Figure 3.5: Project development timeline (approximate Gantt chart)", "26"),
        ("Figure 4.1: UML Use case diagram of EIH developer platform", "31"),
        ("Figure 4.2: Data flow diagram, level 0 (context diagram)", "32"),
        ("Figure 4.3: Data flow diagram, level 1 (detailed ingestion and task workflow)", "32"),
        ("Figure 4.4: UML Activity diagram of task execution, validation and retry loop", "33"),
        ("Figure 4.5: UML Sequence diagram of a grounded RAG query and citation linking", "34"),
        ("Figure 4.6: State transition diagram of an autonomous engineering task lifecycle", "35"),
        ("Figure 4.7: Entity–relationship diagram of the 17 platform database tables", "36"),
        ("Figure 4.8: Cloud deployment architecture: Render backend and Vercel frontend", "38"),
        ("Figure 4.9: Interaction architecture between the 12 specialized autonomous agents", "39"),
        ("Figure 4.10: Code ingestion, AST parsing, and knowledge preparation pipeline", "40"),
        ("Figure 4.11: Distribution of AST code entities extracted across 5 benchmark repositories", "40"),
        ("Figure 4.12: Grounded RAG architecture with AST-driven line citation verification", "42"),
        ("Figure 4.13: Executive Dashboard showing connected repositories and system health", "45"),
        ("Figure 4.14: Code Explorer with recursive file tree and AST symbol catalog", "45"),
        ("Figure 4.15: Grounded RAG Chat with verified line citations and code snippets", "46"),
        ("Figure 4.16: Interactive Architecture Studio with dynamic Mermaid diagrams", "46"),
        ("Figure 4.17: Engineering Task Execution Console with agent status steps", "47"),
        ("Figure 4.18: Change Staging with side-by-side and unified diff viewer and approval gate", "47"),
        ("Figure 4.19: Documentation Studio with 13 technical document generation options", "48"),
        ("Figure 4.20: Responsive mobile phone layout of the Executive Dashboard (390 x 844 px)", "49"),
        ("Figure 4.21: Responsive mobile phone layout of the Code Explorer (390 x 844 px)", "50"),
        ("Figure 5.1: Number of automated pytest test cases per test file (all 140 passed)", "58"),
        ("Figure 5.2: Empirical performance comparison: Single-Agent vs EIH Multi-Agent", "60"),
        ("Figure 5.3: Multi-agent self-correction resolution across retry iterations", "61"),
        ("Figure 5.4: Mean response time and execution latency by platform subsystem", "57"),
        ("Figure 5.5: RAG vector retrieval quality across 36 repository test queries", "61"),
        ("Figure 5.6: Latency vs code quality trade-off across six engineering benchmark tasks", "62"),
        ("Figure 5.8: Repositories management console showing connected codebases and status", "63"),
        ("Figure 5.9: Staged modification diff inspector showing proposed code additions", "63"),
        ("Figure 5.10: Terminal Sandbox execution showing secret redaction and safe test runner", "64"),
        ("Figure 5.11: Swagger OpenAPI REST interactive documentation (/docs endpoint)", "64"),
        ("Figure E.1: Mobile responsive dashboard view (390 x 844 px)", "85"),
        ("Figure E.2: Mobile conversational repository assistant view (390 x 844 px)", "86"),
        ("Figure E.3: Mobile autonomous task execution and status view (390 x 844 px)", "87"),
        ("Figure E.4: Mobile code and file explorer interface (390 x 844 px)", "88"),
        ("Figure E.5: Interactive full-screen system architecture viewer", "88")
    ]
    for title, page in fig_list:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.02), WD_TAB_ALIGNMENT.RIGHT)
        r_t = p.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(9.5)
        dots_len = max(2, int((6.0 - len(title)*0.065) * 12))
        r_d = p.add_run(" " + "." * dots_len + "\t" + page)
        r_d.font.name = 'Times New Roman'
        r_d.font.size = Pt(9.5)

    # 7. LIST OF TABLES (Page ix)
    doc.add_page_break()
    add_heading_chapter(doc, "LIST OF TABLES")
    
    tbl_list = [
        ("Table 1.1: Objectives of the project", "3"),
        ("Table 1.2: Sustainable Development Goal alignment", "6"),
        ("Table 2.1: Comparison of existing approaches and their use in EIH", "10"),
        ("Table 3.1: Functional requirements (FR-01 to FR-10)", "14"),
        ("Table 3.2: Non-functional requirements (NFR-01 to NFR-08)", "15"),
        ("Table 3.3: Software requirements", "17"),
        ("Table 3.4: Hardware requirements", "17"),
        ("Table 3.5: Input–processing–output model", "19"),
        ("Table 3.6: Complexity routing rules and agent workflow assignment", "21"),
        ("Table 3.7: 12 Autonomous specialized agents and primary responsibilities", "23"),
        ("Table 3.8: Work packages (WP-01 to WP-07)", "26"),
        ("Table 3.9: Requirement traceability matrix (RTM)", "28"),
        ("Table 4.1: Major system components across architecture tiers", "30"),
        ("Table 4.2: Relational database schema (17 tables in SQLite / PostgreSQL)", "36"),
        ("Table 4.3: Summary of analyzed software repositories (5 codebases)", "40"),
        ("Table 4.4: Vector store and retrieval model configuration", "42"),
        ("Table 4.5: Main implementation technologies and framework versions", "43"),
        ("Table 4.6: REST API endpoint catalog by router group", "43"),
        ("Table 5.1: Hardware and software specifications of the test machine", "52"),
        ("Table 5.2: Test cases and verification results (TC-01 to TC-20)", "53"),
        ("Table 5.3: Subsystem latency breakdown and execution times (seconds)", "56"),
        ("Table 5.4: Empirical benchmark: Single-Agent vs EIH Multi-Agent Platform", "56"),
        ("Table 5.5: Evaluation metrics and theoretical definitions", "57"),
        ("Table 5.6: Automated pytest suite results by test file (all 140 passed)", "58"),
        ("Table 5.7: Multi-agent self-correction recovery statistics across retry loops", "60"),
        ("Table 5.12: Observed errors, edge-case limitations, and remediation mechanisms", "64"),
        ("Table 5.13: Requirement–implementation–testing traceability validation", "66"),
        ("Table 5.14: Reproducibility checklist and environment parameters", "67"),
        ("Table 6.1: Achievement of project objectives defined in Chapter 1", "69"),
        ("Table B.1: Entity extraction statistics across the 5 benchmark repositories", "77"),
        ("Table B.2: Query retrieval citations and line verification evaluation", "77"),
        ("Table B.3: Multi-agent execution retry logs for complex engineering tasks", "78"),
        ("Table B.4: Comprehensive latency and token consumption metrics", "79"),
        ("Table C.1: Standardized Developer Usability Evaluation Rubric", "80"),
        ("Table D.1: Engineering Intelligence Hub console views and navigation guide", "82"),
        ("Table D.2: Common troubleshooting problems and remediation steps", "83"),
        ("Table F.1: Project team member contributions and module development matrix", "90")
    ]
    for title, page in tbl_list:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.05
        p.paragraph_format.tab_stops.add_tab_stop(Inches(6.02), WD_TAB_ALIGNMENT.RIGHT)
        r_t = p.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(9.5)
        dots_len = max(2, int((6.0 - len(title)*0.065) * 12))
        r_d = p.add_run(" " + "." * dots_len + "\t" + page)
        r_d.font.name = 'Times New Roman'
        r_d.font.size = Pt(9.5)

    # 8. LIST OF ABBREVIATIONS (Page x)
    doc.add_page_break()
    add_heading_chapter(doc, "LIST OF ABBREVIATIONS")
    
    t_abbr = doc.add_table(rows=24, cols=2)
    t_abbr.rows[0].cells[0].text = "Abbreviation"
    t_abbr.rows[0].cells[1].text = "Full Form"
    t_abbr.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_abbr.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    abbr_data = [
        ("AI", "Artificial Intelligence"),
        ("API", "Application Programming Interface"),
        ("ASGI", "Asynchronous Server Gateway Interface"),
        ("AST", "Abstract Syntax Tree"),
        ("CI/CD", "Continuous Integration / Continuous Deployment"),
        ("CORS", "Cross-Origin Resource Sharing"),
        ("CSV", "Comma-Separated Values"),
        ("DFD", "Data Flow Diagram"),
        ("EIH", "Engineering Intelligence Hub"),
        ("ER", "Entity–Relationship"),
        ("Git", "Distributed Version Control System"),
        ("HTTP", "Hypertext Transfer Protocol"),
        ("JSON", "JavaScript Object Notation"),
        ("LLM", "Large Language Model"),
        ("MRR", "Mean Reciprocal Rank"),
        ("NLP", "Natural Language Processing"),
        ("ORM", "Object-Relational Mapping"),
        ("RAG", "Retrieval-Augmented Generation"),
        ("REST", "Representational State Transfer"),
        ("RTM", "Requirement Traceability Matrix"),
        ("SDD", "Software Design Document"),
        ("SDG", "Sustainable Development Goal"),
        ("UI", "User Interface")
    ]
    for i, (ab, ff) in enumerate(abbr_data):
        t_abbr.rows[i+1].cells[0].text = ab
        t_abbr.rows[i+1].cells[1].text = ff
        t_abbr.rows[i+1].cells[0].paragraphs[0].runs[0].font.bold = True
    style_table(t_abbr, col_widths=[Inches(1.8), Inches(4.2)])
