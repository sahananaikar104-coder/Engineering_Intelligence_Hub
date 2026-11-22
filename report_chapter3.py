import os
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_styles import (
    add_p, add_heading_chapter, add_heading_1, add_heading_2, add_heading_3,
    add_bullet, style_table, add_fig, add_code_listing
)

def build_chapter_3(doc):
    add_heading_chapter(doc, "CHAPTER 3 – REQUIREMENTS AND PROPOSED METHODOLOGY")
    
    lead_text = (
        "This chapter defines the comprehensive engineering requirements, feasibility criteria, operational methodology, "
        "and architectural algorithms of the Engineering Intelligence Hub (EIH). EIH is designed around three tightly integrated "
        "processing streams: (1) an AST Ingestion and Code Intelligence path that indexes repository structures and manifests into "
        "relational and vector databases; (2) a Grounded RAG Query path that provides developers with conversational codebase "
        "insights backed by exact line citations; and (3) an Autonomous Multi-Agent Task Execution path that plans, generates, "
        "validates, and stages complex code modifications with automated self-correction and human-in-the-loop review. Formal "
        "functional requirements, mathematical algorithms, project planning breakdowns, and a traceability matrix are specified."
    )
    add_p(doc, lead_text, space_after=12, line_spacing=1.2)
    
    # 3.1 Introduction
    add_heading_1(doc, "3.1 Introduction")
    p1 = (
        "Developing an autonomous, repository-aware developer platform necessitates a clear demarcation of functional and "
        "non-functional boundaries. Traditional coding assistants fail because they do not validate their outputs against the host "
        "environment or provide deterministic citation proofs. EIH establishes a principled engineering methodology wherein static "
        "program analysis (Python AST, manifest parsers) acts as an infallible ground truth, guiding dense semantic vector search "
        "(ChromaDB) and constraining generative language models. Every requirement defined in this chapter directly maps to "
        "concrete software components and automated test suites."
    )
    add_p(doc, p1, space_after=10, line_spacing=1.2)
    
    # 3.2 Requirement Collection
    add_heading_1(doc, "3.2 Requirement Collection")
    add_p(doc, "Requirements were systematically derived from the approved project synopsis, real-world developer workflows on "
               "complex GitHub repositories, and empirical failure analyses of existing AI coding tools. They are categorized into "
               "user, functional, non-functional, software, hardware, and dataset requirements.", space_after=6)
               
    add_heading_2(doc, "3.2.1 User Requirements")
    add_bullet(doc, "The software developer or architect should be able to connect any public or private GitHub repository or "
                    "local folder via a simple URL or directory path.", bold_prefix="• Repository Ingestion:")
    add_bullet(doc, "The user should be able to navigate the repository file tree, inspect extracted AST symbols (classes, methods, "
                    "signatures, docstrings), and view syntax-highlighted source code.", bold_prefix="• Code Exploration:")
    add_bullet(doc, "The user should be able to ask natural-language questions about the codebase and receive concise answers with "
                    "clickable, line-level citations linking directly to the source code.", bold_prefix="• Grounded Q&A:")
    add_bullet(doc, "The developer should be able to submit high-level engineering tasks (features, bug fixes, refactoring) and "
                    "monitor the autonomous multi-agent execution pipeline in real time.", bold_prefix="• Task Automation:")
    add_bullet(doc, "The user must be able to inspect proposed changes in an isolated staging workspace via side-by-side unified "
                    "diffs and explicitly Approve, Reject, or Request Revision.", bold_prefix="• Human Review Gate:")
    add_bullet(doc, "The user should be able to generate and synchronize 13 technical document types and export interactive Mermaid "
                    "architecture and workflow diagrams.", bold_prefix="• Documentation & Diagrams:")
                    
    add_heading_2(doc, "3.2.2 Functional Requirements")
    add_p(doc, "The ten major functional requirements of the platform are formally cataloged in Table 3.1.", space_after=6)
    
    t_fr = doc.add_table(rows=11, cols=3)
    t_fr.rows[0].cells[0].text = "ID"
    t_fr.rows[0].cells[1].text = "Functional Requirement"
    t_fr.rows[0].cells[2].text = "Technical Specification & Scope"
    for j in range(3):
        t_fr.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    fr_data = [
        ("FR-01", "Repository Ingestion & Management", 
         "Safely clone remote Git repositories, discover file hierarchies, parse package manifests, and exclude credentials and binary assets."),
        ("FR-02", "AST Code Parsing & Extraction", 
         "Parse Python AST to extract functions, classes, decorators, docstrings, imports, call expressions, and exact line bounds."),
        ("FR-03", "Code-Aware Chunking & Embedding", 
         "Split code along semantic entity boundaries (classes/functions) and generate 384-dimensional dense embeddings stored in ChromaDB."),
        ("FR-04", "Grounded RAG & Citation Linking", 
         "Retrieve relevant chunks using cosine similarity boosted by AST symbols; generate answers and enforce exact line-level citation tags."),
        ("FR-05", "Task Orchestration & Complexity Routing", 
         "Classify engineering requirements into SIMPLE, MEDIUM, or COMPLEX and dynamically route execution across 12 specialized agents."),
        ("FR-06", "Automated Code & Test Generation", 
         "Synthesize clean, idiomatic source modifications and automated pytest test suites covering edge cases in an isolated workspace."),
        ("FR-07", "Sandbox Validation & Test Execution", 
         "Execute AST syntax validation, import resolution, and run pytest test suites in an isolated subprocess with strict execution timeouts."),
        ("FR-08", "Self-Correction Diagnostic Loop", 
         "Capture validation tracebacks and test failures; iteratively re-prompt the code generator up to 3 times to achieve resolution."),
        ("FR-09", "Change Staging & Human Approval Gate", 
         "Stage all modifications as unified diffs; prevent direct repository overwrites until developer clicks Approve, Reject, or Revise."),
        ("FR-10", "Documentation & Architecture Studio", 
         "Generate 13 technical document types, render interactive Mermaid diagrams (flowchart, sequence, ER), and detect documentation drift.")
    ]
    for i, row in enumerate(fr_data):
        for j, val in enumerate(row):
            t_fr.rows[i+1].cells[j].text = val
            if j == 0:
                t_fr.rows[i+1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_table(t_fr, col_widths=[Inches(0.8), Inches(2.2), Inches(3.0)])
    
    # 3.2.3 Non-Functional Requirements
    add_heading_2(doc, "3.2.3 Non-Functional Requirements")
    add_p(doc, "Non-functional requirements governing security, performance, explainability, and maintainability are listed in Table 3.2.", space_after=6)
    
    t_nfr = doc.add_table(rows=9, cols=3)
    t_nfr.rows[0].cells[0].text = "ID"
    t_nfr.rows[0].cells[1].text = "Requirement Area"
    t_nfr.rows[0].cells[2].text = "Engineering Standard & Metric"
    for j in range(3):
        t_nfr.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    nfr_data = [
        ("NFR-01", "Security & Secret Redaction", 
         "All tokens, API keys (OpenAI, GitHub, Groq), and DB passwords must be detected and masked (e.g., [REDACTED_SECRET]) before storage or display."),
        ("NFR-02", "Explainability & Grounding", 
         "Every claim made by the chat engine must cite exact file_path, start_line, and end_line; ungrounded statements are strictly flagged."),
        ("NFR-03", "Fault Tolerance & Safe Fallback", 
         "If an external LLM API fails or rate-limits, the system must smoothly fallback to local high-fidelity mock or alternate providers without crashing."),
        ("NFR-04", "Performance & Latency", 
         "File scanning < 50ms, AST parsing < 150ms, vector search < 50ms, and RAG answer generation < 1.0s for typical queries."),
        ("NFR-05", "Isolation & Code Safety", 
         "Zero in-place repository overwrites; all changes must execute in isolated staging workspaces with absolute diff tracking."),
        ("NFR-06", "Modularity & Extensibility", 
         "Backend components must adhere to clean hexagonal/service architecture with pluggable LLMProvider and database ORM interfaces."),
        ("NFR-07", "Responsive Usability", 
         "Frontend must provide a responsive dark-mode UI operable across desktop monitors (1440x900) and mobile viewports (390x844)."),
        ("NFR-08", "Reproducibility & Testability", 
         "All platform features must be covered by automated test suites executable with a single command (100% test pass target).")
    ]
    for i, row in enumerate(nfr_data):
        for j, val in enumerate(row):
            t_nfr.rows[i+1].cells[j].text = val
            if j == 0:
                t_nfr.rows[i+1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_table(t_nfr, col_widths=[Inches(0.8), Inches(1.8), Inches(3.4)])
    
    # 3.2.4 & 3.2.5 Software and Hardware Requirements
    add_heading_2(doc, "3.2.4 Software Requirements")
    add_p(doc, "The platform software stack is specified in Table 3.3.", space_after=6)
    
    t_sw = doc.add_table(rows=11, cols=2)
    t_sw.rows[0].cells[0].text = "Software / Framework"
    t_sw.rows[0].cells[1].text = "Purpose & Role in EIH"
    for j in range(2):
        t_sw.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    sw_data = [
        ("Python 3.11+ / Uvicorn", "Core backend runtime, asynchronous event loop, and ASGI web server."),
        ("FastAPI / Pydantic v2", "High-performance REST API routing, request validation, and OpenAPI documentation."),
        ("SQLAlchemy 2.0 (aiosqlite)", "Asynchronous Object-Relational Mapping for relational metadata persistence."),
        ("ChromaDB 0.4.24+", "Embedded vector database for dense code embeddings and semantic similarity search."),
        ("Sentence-Transformers", "all-MiniLM-L6-v2 embedding model generating 384-dimensional dense vector representations."),
        ("GitPython / Git CLI", "Cloning repositories, extracting commit trees, and managing ephemeral worktrees."),
        ("React 18 / Vite / TypeScript", "Frontend single-page application, type-safe components, and fast HMR development server."),
        ("Tailwind CSS / Lucide Icons", "Modern engineering console dark-mode theme, glassmorphism UI, and icon sets."),
        ("Monaco Editor / Mermaid.js", "In-browser code viewing, syntax highlighting, diff inspection, and live diagram rendering."),
        ("Pytest / Pytest-Asyncio", "Automated testing framework for executing unit, integration, and sandbox test suites.")
    ]
    for i, row in enumerate(sw_data):
        for j, val in enumerate(row):
            t_sw.rows[i+1].cells[j].text = val
    style_table(t_sw, col_widths=[Inches(2.2), Inches(3.8)])
    
    add_heading_2(doc, "3.2.5 Hardware Requirements")
    add_p(doc, "Hardware specifications for development, testing, and deployment are detailed in Table 3.4.", space_after=6)
    
    t_hw = doc.add_table(rows=6, cols=2)
    t_hw.rows[0].cells[0].text = "Hardware Component"
    t_hw.rows[0].cells[1].text = "Minimum / Recommended Specification"
    for j in range(2):
        t_hw.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    hw_data = [
        ("Processor", "Modern multi-core x86_64 CPU (Intel Core i5/i7 or AMD Ryzen 5/7, 4+ cores)."),
        ("System Memory (RAM)", "Minimum 8 GB; 16 GB recommended for concurrent ChromaDB embedding and local LLM execution."),
        ("Storage", "Minimum 20 GB free SSD storage for repository clones, vector embeddings, and SQLite database."),
        ("Graphics (GPU)", "Optional; CPU inference is fully sufficient for all-MiniLM-L6-v2 embeddings and cloud LLM APIs."),
        ("Network Interface", "Standard broadband connection (10+ Mbps) for cloning remote GitHub repositories and cloud LLM calls.")
    ]
    for i, row in enumerate(hw_data):
        for j, val in enumerate(row):
            t_hw.rows[i+1].cells[j].text = val
    style_table(t_hw, col_widths=[Inches(2.0), Inches(4.0)])
    
    add_heading_2(doc, "3.2.6 Dataset Requirements")
    add_p(doc, "Unlike conventional supervised ML classification projects that depend on fixed image or tabular datasets, the "
               "dataset for the Engineering Intelligence Hub is the software repository itself. EIH ingests, indexes, and analyzes "
               "five benchmark multi-tier open-source repositories representing diverse application domains: a FastAPI Microservice, "
               "an Authentication Service, a Task Management API, a React Frontend, and a Python Utility Tools library. Across these "
               "repositories, the platform indexes 117 source files, 110 AST code entities, 100 dependency manifests, and 109 documentation "
               "items. No proprietary or personal customer data is stored or required.", space_after=10)
               
    # 3.3 Requirement Analysis
    add_heading_1(doc, "3.3 Requirement Analysis")
    add_p(doc, "Feasibility analysis confirms the technical, operational, and economic viability of the proposed architecture.", space_after=6)
    
    add_bullet(doc, "The core technologies (Python 3.11, FastAPI, React 18, SQLite, ChromaDB) are mature, open-source frameworks "
                    "with extensive enterprise backing. AST parsing executes deterministically within milliseconds on modern CPUs. "
                    "Vector embeddings and RAG retrieval operate entirely in-memory or on local disk without specialized GPU hardware.", 
               bold_prefix="• Technical Feasibility:")
    add_bullet(doc, "The unified engineering console provides intuitive navigation across repositories, symbols, chat, tasks, diffs, "
                    "and documentation. Clear visual indicators, real-time agent status logs, and side-by-side diff viewers ensure that "
                    "developers can operate the system without complex training.", 
               bold_prefix="• Operational Feasibility:")
    add_bullet(doc, "The entire platform is constructed using permissive open-source technologies (MIT/Apache 2.0). Cloud deployment "
                    "leverages free hosting tiers on Render and Vercel, incurring zero baseline operational infrastructure cost.", 
               bold_prefix="• Economic Feasibility:")
    add_bullet(doc, "All external API keys (OpenAI, Groq, GitHub) are managed strictly via environment variables; network connectivity "
                    "is required only when interacting with remote GitHub endpoints or hosted LLMs.", 
               bold_prefix="• Constraints & Assumptions:")
               
    # 3.4 Proposed Solution
    add_heading_1(doc, "3.4 Proposed Solution")
    add_p(doc, "The Engineering Intelligence Hub implements a three-tier architecture that decouples repository ingestion, conversational "
               "code exploration, and multi-agent task execution. Rather than treating code as plain text, EIH builds a structured "
               "knowledge model combining relational AST metadata with dense vector embeddings. Natural-language questions are answered "
               "with verified line-range citations, while engineering tasks are executed through a self-correcting 12-agent state machine "
               "that isolates all proposed modifications in a staging diff viewer awaiting human sign-off.", space_after=10)
               
    # 3.5 Proposed Methodology
    add_heading_1(doc, "3.5 Proposed Methodology")
    add_p(doc, "The methodology organizes the lifecycle of repository intelligence into three distinct execution paths, as formalized "
               "in the Input-Processing-Output model (Table 3.5) and illustrated in Figure 3.1.", space_after=6)
               
    t_ipo = doc.add_table(rows=5, cols=3)
    t_ipo.rows[0].cells[0].text = "Input Artifact"
    t_ipo.rows[0].cells[1].text = "Core Processing Workflow"
    t_ipo.rows[0].cells[2].text = "Verified Output Deliverable"
    for j in range(3):
        t_ipo.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    ipo_data = [
        ("GitHub URL / Local Directory Path", 
         "Safe Git clone, ignore filtering (.git, secrets), manifest parsing, Python AST entity extraction, code-aware chunking, dense vector indexing.", 
         "Indexed SQLite symbol database, 384-d ChromaDB vector collection, file hierarchy tree, dependency topology."),
        ("Natural-Language Developer Query", 
         "Query embedding, hybrid semantic search, symbol keyword boosting, top-k chunk retrieval, prompt assembly, citation verification.", 
         "Contextual answer with interactive, line-level source code citations (file_path:start_line-end_line)."),
        ("Engineering Task Requirement (Feature/Fix)", 
         "Complexity classification, 12-agent orchestration (Requirement -> Design -> CodeGen -> TestGen -> Validation -> Review), sandbox execution.", 
         "Staged code modifications, automated pytest test suite, Software Design Document, Mermaid diagrams, review scores."),
        ("Test / Validation Failure Diagnostic", 
         "Traceback parsing, error classification, prompt re-engineering with diagnostic context, iterative code re-synthesis (up to 3 retries).", 
         "Self-corrected code modifications that successfully pass AST validation and sandbox pytest execution.")
    ]
    for i, row in enumerate(ipo_data):
        for j, val in enumerate(row):
            t_ipo.rows[i+1].cells[j].text = val
    style_table(t_ipo, col_widths=[Inches(1.8), Inches(2.6), Inches(1.8)])
    
    fig_dir = "extracted_figures"
    add_fig(doc, os.path.join(fig_dir, "fig3_1_methodology.png"), 
            "Figure 3.1: Proposed methodology: three processing paths from ingestion to consumer.", width=Inches(5.8))
            
    # 3.6 Proposed System Architecture
    add_heading_1(doc, "3.6 Proposed System Architecture")
    add_p(doc, "The high-level system architecture of the Engineering Intelligence Hub is depicted in Figure 3.2. The system encompasses "
               "six coordinated horizontal layers: Presentation Layer (React 18 + TypeScript + Vite + Tailwind CSS), API Gateway Layer "
               "(FastAPI async routers), Core Services Layer (AST Parser, File Scanner, Sandbox Runner, Vector Retriever), 12-Agent "
               "Orchestration Layer (State Machine and Self-Correction Controller), Persistence Layer (SQLite/PostgreSQL, ChromaDB, "
               "Git Staging Workspaces), and the Pluggable LLM Provider Layer (OpenAI, Groq, Ollama, Anthropic, and Local Mock).", space_after=8)
               
    add_fig(doc, os.path.join(fig_dir, "fig3_2_architecture.png"), 
            "Figure 3.2: High-level system architecture of Engineering Intelligence Hub (EIH).", width=Inches(5.6))
            
    # 3.7 Algorithms and Techniques
    add_heading_1(doc, "3.7 Algorithms and Techniques")
    add_p(doc, "The deterministic and agentic mechanisms governing EIH are specified below through formal algorithms and rule definitions.", space_after=6)
    
    add_heading_2(doc, "3.7.1 Complexity Routing and Orchestrator Workflow")
    add_p(doc, "The Orchestrator Agent analyzes incoming natural-language requirements and assigns a complexity tier (SIMPLE, MEDIUM, COMPLEX) "
               "to prevent unnecessary compute overhead for trivial lookups, as defined in Table 3.6 and Algorithm 1.", space_after=6)
               
    t_comp_rules = doc.add_table(rows=4, cols=3)
    t_comp_rules.rows[0].cells[0].text = "Complexity Tier"
    t_comp_rules.rows[0].cells[1].text = "Requirement Heuristic / Keywords"
    t_comp_rules.rows[0].cells[2].text = "Activated Agent Workflow Sequence"
    for j in range(3):
        t_comp_rules.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    comp_rules_data = [
        ("SIMPLE", 
         "'explain', 'where is', 'what is', single-symbol inspection, word count < 6 without modification verbs.", 
         "Repository Agent -> Code Understanding Agent -> RAG Generation -> Line Citation Builder."),
        ("MEDIUM", 
         "'test', 'fix', single-function bug fix, isolated method addition, word count <= 10.", 
         "Requirement Agent -> CodeGen Agent -> TestGen Agent -> Validation Agent -> Diff Stager."),
        ("COMPLEX", 
         "Multi-objective statements, 'auth', 'jwt', 'database', 'refactor', 'api', multiple file targets.", 
         "Full 12-Agent Workflow: Requirement -> Repo -> CodeUnder -> Design -> CodeGen -> TestGen -> Diagram -> Doc -> Validation -> Review.")
    ]
    for i, row in enumerate(comp_rules_data):
        for j, val in enumerate(row):
            t_comp_rules.rows[i+1].cells[j].text = val
            if j == 0:
                t_comp_rules.rows[i+1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_table(t_comp_rules, col_widths=[Inches(1.2), Inches(2.4), Inches(2.6)])
    
    alg1_code = (
        "Algorithm 1: Orchestrator Task Execution & Self-Correction Workflow\n"
        "Input: requirement_str, repository_id, staging_workspace\n"
        "Output: AgentState { status: 'staged' | 'failed', diffs, tests, scores }\n"
        "\n"
        "1: state <- InitializeAgentState(requirement_str, repository_id, staging_workspace)\n"
        "2: state.complexity <- ClassifyComplexity(requirement_str)\n"
        "3: if state.complexity == 'SIMPLE' then\n"
        "4:     return ExecuteSimpleLookupWorkflow(state)\n"
        "5: end if\n"
        "6: \n"
        "7: state.task_plan <- RequirementAgent.Decompose(state.user_requirement)\n"
        "8: state.repo_context <- RepositoryAnalysisAgent.Scan(state.workspace_path)\n"
        "9: state.symbols <- CodeUnderstandingAgent.TraceCallGraph(state.workspace_path)\n"
        "10: state.design_doc <- DesignAgent.SynthesizeSDD(state.task_plan, state.symbols)\n"
        "11: \n"
        "12: repeat\n"
        "13:     state.proposed_changes <- CodeGenerationAgent.GenerateDiffs(state)\n"
        "14:     state.generated_tests <- TestGenerationAgent.GeneratePytestSuites(state)\n"
        "15:     state.validation_results <- ValidationAgent.RunSandbox(state.workspace_path)\n"
        "16:     \n"
        "17:     if state.validation_results.syntax_ok and state.validation_results.tests_passed then\n"
        "18:         state.review_results <- ReviewAgent.AuditSecurityAndQuality(state)\n"
        "19:         state.diagrams <- DiagramAgent.GenerateMermaid(state)\n"
        "20:         state.documentation <- DocumentationAgent.UpdateTechnicalDocs(state)\n"
        "21:         state.status <- 'staged'\n"
        "22:         return state   # Ready for Human-in-the-Loop review\n"
        "23:     else\n"
        "24:         state.retry_count <- state.retry_count + 1\n"
        "25:         state.diagnostic_error <- state.validation_results.error_summary\n"
        "26:         LogSelfCorrectionIteration(state.retry_count, state.diagnostic_error)\n"
        "27:     end if\n"
        "28: until state.retry_count >= MAX_RETRIES (3)\n"
        "29: \n"
        "30: state.status <- 'failed'\n"
        "31: return state"
    )
    add_code_listing(doc, alg1_code, "Algorithm 1: Orchestrator Task Execution & Self-Correction Workflow")
    
    add_heading_2(doc, "3.7.2 12 Specialized Autonomous Agents")
    add_p(doc, "The division of responsibilities across the 12 specialized agents is summarized in Table 3.7.", space_after=6)
    
    t_agents = doc.add_table(rows=13, cols=3)
    t_agents.rows[0].cells[0].text = "Agent Name"
    t_agents.rows[0].cells[1].text = "Primary Responsibility"
    t_agents.rows[0].cells[2].text = "Produced Artifact / Deliverable"
    for j in range(3):
        t_agents.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    agents_data = [
        ("1. Orchestrator Agent", "Lifecycle control, complexity routing, and self-correction controller.", "Workflow execution state & status."),
        ("2. Requirement Agent", "Decomposes natural language into discrete, ordered tasks.", "Structured TaskPlan list."),
        ("3. Repository Agent", "Scans file trees, manifests, frameworks, and entry points.", "Tech stack & manifest metadata."),
        ("4. Code Understanding", "Extracts AST symbols, call hierarchies, and execution traces.", "Symbol graph & route catalog."),
        ("5. System Design Agent", "Designs software architecture, component models, and contracts.", "Software Design Document (SDD)."),
        ("6. Code Generation", "Synthesizes clean code modifications and unified diffs in staging.", "Staged file changes & diffs."),
        ("7. Test Generation", "Synthesizes automated pytest test suites covering edge cases.", "tests/test_*.py test files."),
        ("8. Documentation Agent", "Updates technical README, API documentation, and changelogs.", "Markdown documentation items."),
        ("9. Diagram Studio Agent", "Generates Mermaid component, sequence, and flowchart diagrams.", "Mermaid diagram code & SVG."),
        ("10. Validation Agent", "Executes AST syntax checks, imports, and sandbox test runners.", "Validation report & error trace."),
        ("11. Review Agent", "Performs static security audits and code quality scoring (0–100).", "Quality & Security scorecards."),
        ("12. Change Impact Agent", "Identifies documentation drift caused by code modifications.", "Outdated doc flags & gap list.")
    ]
    for i, row in enumerate(agents_data):
        for j, val in enumerate(row):
            t_agents.rows[i+1].cells[j].text = val
            if j == 0:
                t_agents.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_agents, col_widths=[Inches(1.8), Inches(2.4), Inches(2.0)])
    
    add_heading_2(doc, "3.7.3 Code-Aware AST Chunking with Line Preservation")
    add_p(doc, "To prevent function truncation during vector indexing, Algorithm 2 extracts discrete AST classes and functions as "
               "atomic chunks and attaches exact line coordinates.", space_after=6)
               
    alg2_code = (
        "Algorithm 2: Code-Aware AST Chunking with Line Preservation\n"
        "Input: file_path, source_content, language, target_lines = 50, overlap = 10\n"
        "Output: List of Chunks [{ file_path, start_line, end_line, content, entity_name }]\n"
        "\n"
        "1: lines <- SplitLines(source_content)\n"
        "2: total_lines <- len(lines)\n"
        "3: chunks <- []\n"
        "4: \n"
        "5: if language == 'Python' then\n"
        "6:     ast_entities <- PythonASTParser.ExtractClassesAndFunctions(source_content)\n"
        "7:     for entity in ast_entities do\n"
        "8:         s_line <- max(1, entity.start_line)\n"
        "9:         e_line <- min(total_lines, entity.end_line)\n"
        "10:        chunk_text <- Join(lines[s_line - 1 : e_line], '\\n')\n"
        "11:        chunks.append({ file_path: file_path, start_line: s_line, end_line: e_line,\n"
        "12:                        content: chunk_text, entity_name: entity.name, entity_type: entity.type })\n"
        "13:    end for\n"
        "14: end if\n"
        "15: \n"
        "16: if len(chunks) == 0 then  # Fallback sliding window for non-Python or non-entity code\n"
        "17:     curr_line <- 0\n"
        "18:     while curr_line < total_lines do\n"
        "19:         end_line <- min(total_lines, curr_line + target_lines)\n"
        "20:         chunk_text <- Join(lines[curr_line : end_line], '\\n')\n"
        "21:         chunks.append({ file_path: file_path, start_line: curr_line + 1, end_line: end_line,\n"
        "22:                         content: chunk_text, entity_name: None, entity_type: 'file_chunk' })\n"
        "23:         curr_line <- curr_line + (target_lines - overlap)\n"
        "24:     end while\n"
        "25: end if\n"
        "26: return chunks"
    )
    add_code_listing(doc, alg2_code, "Algorithm 2: Code-Aware AST Chunking with Line Preservation")
    
    add_heading_2(doc, "3.7.4 Verifiable Line Citation Resolution")
    add_p(doc, "Algorithm 3 cross-examines generated answers against AST symbol boundaries, extracting verified line-level references.", space_after=6)
    
    alg3_code = (
        "Algorithm 3: Verifiable Line Citation Builder\n"
        "Input: query_str, retrieved_chunks, raw_llm_response, ast_symbol_db\n"
        "Output: VerifiedResponse { answer_text, verified_citations: List[Citation] }\n"
        "\n"
        "1: citations <- []\n"
        "2: cited_references <- ExtractCitationSpans(raw_llm_response) # detects [file:lines] patterns\n"
        "3: \n"
        "4: for ref in cited_references do\n"
        "5:     real_symbol <- ast_symbol_db.QuerySymbol(ref.file_path, ref.line_number)\n"
        "6:     if real_symbol is not None then\n"
        "7:         citations.append({\n"
        "8:             file_path: ref.file_path,\n"
        "9:             start_line: real_symbol.start_line,\n"
        "10:            end_line: real_symbol.end_line,\n"
        "11:            snippet: ExtractSourceLines(ref.file_path, real_symbol.start_line, real_symbol.end_line),\n"
        "12:            entity_name: real_symbol.name,\n"
        "13:            verified: True\n"
        "14:        })\n"
        "15:    else\n"
        "16:        citations.append({ file_path: ref.file_path, start_line: ref.line, end_line: ref.line, verified: False })\n"
        "17:    end if\n"
        "18: end for\n"
        "19: return VerifiedResponse(raw_llm_response, citations)"
    )
    add_code_listing(doc, alg3_code, "Algorithm 3: Verifiable Line Citation Builder")
    
    add_heading_2(doc, "3.7.5 Sandbox Validation and Self-Correction Loop")
    add_p(doc, "Figure 3.3 illustrates the decision flow of the multi-agent state machine, while Algorithm 4 governs automated retry execution.", space_after=6)
    
    alg4_code = (
        "Algorithm 4: Sandbox Validation & Diagnostic Error Recovery\n"
        "Input: staging_workspace, proposed_files, generated_tests, retry_limit = 3\n"
        "Output: ValidationResult { passed: bool, error_trace: str }\n"
        "\n"
        "1: for file in proposed_files do\n"
        "2:     ast_res <- PythonASTParser.Parse(file.staged_content)\n"
        "3:     if ast_res.has_syntax_error then\n"
        "4:         return ValidationResult(passed=False, error_trace=ast_res.error_message)\n"
        "5:     end if\n"
        "6: end for\n"
        "7: \n"
        "8: test_proc <- Subprocess.Run(['pytest', 'tests'], cwd=staging_workspace, timeout=30)\n"
        "9: if test_proc.returncode != 0 then\n"
        "10:    sanitized_output <- SecretRedactor.Mask(test_proc.stderr + test_proc.stdout)\n"
        "11:    return ValidationResult(passed=False, error_trace=sanitized_output)\n"
        "12: end if\n"
        "13: return ValidationResult(passed=True, error_trace=None)"
    )
    add_code_listing(doc, alg4_code, "Algorithm 4: Sandbox Validation & Diagnostic Error Recovery")
    
    add_heading_2(doc, "3.7.6 Change Staging and Human Approval Gating")
    add_p(doc, "Algorithm 5 enforces zero in-place mutations by gating repository modifications behind a developer approval modal.", space_after=6)
    
    alg5_code = (
        "Algorithm 5: Human-in-the-Loop Change Staging & Application Gate\n"
        "Input: task_id, action: 'approve' | 'reject' | 'revise', feedback_str\n"
        "Output: { applied: bool, message: str }\n"
        "\n"
        "1: task <- DB.GetTask(task_id)\n"
        "2: if action == 'approve' then\n"
        "3:     for change in task.proposed_changes do\n"
        "4:         ApplyUnifiedDiff(target_repo=task.repo_path, diff=change.diff_content)\n"
        "5:     end for\n"
        "6:     task.status <- 'approved'\n"
        "7:     return { applied: True, message: 'Changes successfully merged into repository.' }\n"
        "8: else if action == 'reject' then\n"
        "9:     WorkspaceManager.DiscardStagingWorkspace(task.workspace_path)\n"
        "10:    task.status <- 'rejected'\n"
        "11:    return { applied: False, message: 'Proposed modifications discarded.' }\n"
        "12: else if action == 'revise' then\n"
        "13:    task.status <- 'revising'\n"
        "14:    return OrchestratorAgent.ExecuteRevision(task, feedback_str)\n"
        "15: end if"
    )
    add_code_listing(doc, alg5_code, "Algorithm 5: Human-in-the-Loop Change Staging & Application Gate")
    
    # 3.8 Project Planning
    add_heading_1(doc, "3.8 Project Planning")
    add_p(doc, "The project was structured into seven discrete Work Packages (WP-01 to WP-07) as shown in Figure 3.4 and Table 3.8. "
               "The approximate development timeline spanning August to November 2026 is graphed in Figure 3.5.", space_after=8)
               
    add_fig(doc, os.path.join(fig_dir, "fig3_4_wbs.png"), 
            "Figure 3.4: Work breakdown structure (WBS) of the EIH platform project.", width=Inches(5.8))
            
    t_wp = doc.add_table(rows=8, cols=3)
    t_wp.rows[0].cells[0].text = "Work Package"
    t_wp.rows[0].cells[1].text = "Key Technical Activities"
    t_wp.rows[0].cells[2].text = "Milestone Deliverables"
    for j in range(3):
        t_wp.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    wp_data = [
        ("WP-01: Study & Planning", 
         "Literature review, requirement specifications, synopsis approval, architecture design.", 
         "Approved synopsis, architecture blueprint, SRS document."),
        ("WP-02: Code Ingestion & AST", 
         "Git clone manager, directory tree discovery, manifest parsers, Python AST entity extraction.", 
         "Ingestion service, AST parser, relational symbol catalog."),
        ("WP-03: Vector Indexing & RAG", 
         "Code-aware line chunker, ChromaDB integration, all-MiniLM-L6-v2 embeddings, citation builder.", 
         "Vector store pipeline, Grounded RAG Chat API with citations."),
        ("WP-04: Multi-Agent State Machine", 
         "Orchestrator implementation, 12 specialized agents, prompt engineering, complexity router.", 
         "Agent state machine, task planning, SDD generator, code/test synthesis."),
        ("WP-05: Validation & Self-Correction", 
         "Subprocess sandbox runner, pytest executor, diagnostic traceback parser, retry loop.", 
         "Sandbox test suite, self-correction engine (<= 3 retries)."),
        ("WP-06: Staging & User Interface", 
         "Staging workspace manager, unified diff viewer, React 18 console, Mermaid studio.", 
         "Responsive dark-mode UI, diff viewer, terminal, cloud configs."),
        ("WP-07: Testing & Documentation", 
         "Pytest test suite (140 tests), empirical benchmark suite, latency profiling, final report.", 
         "Automated test report, empirical benchmarks, final project report.")
    ]
    for i, row in enumerate(wp_data):
        for j, val in enumerate(row):
            t_wp.rows[i+1].cells[j].text = val
            if j == 0:
                t_wp.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_wp, col_widths=[Inches(1.8), Inches(2.4), Inches(2.0)])
    
    add_fig(doc, os.path.join(fig_dir, "fig3_5_timeline.png"), 
            "Figure 3.5: Project development timeline (approximate Gantt chart).", width=Inches(5.8))
            
    # 3.9 Requirement Traceability Matrix
    add_heading_1(doc, "3.9 Requirement Traceability Matrix")
    add_p(doc, "Table 3.9 maps each functional requirement to its implementation files in the backend and frontend codebases and "
               "to its verifying automated test cases in Chapter 5.", space_after=6)
               
    t_rtm = doc.add_table(rows=11, cols=4)
    t_rtm.rows[0].cells[0].text = "Req ID"
    t_rtm.rows[0].cells[1].text = "Requirement Name"
    t_rtm.rows[0].cells[2].text = "Backend / Frontend Source Implementation"
    t_rtm.rows[0].cells[3].text = "Test Case ID"
    for j in range(4):
        t_rtm.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    rtm_data = [
        ("FR-01", "Repository Ingestion", "app/repository/git_service.py, file_scanner.py", "TC-01, TC-05"),
        ("FR-02", "AST Code Parsing", "app/code_intelligence/ast_parser.py, generic_parser.py", "TC-02, TC-10"),
        ("FR-03", "Code-Aware Chunking", "app/rag/chunker.py, embeddings.py, vector_store.py", "TC-03, TC-18"),
        ("FR-04", "Grounded RAG Chat", "app/rag/citation_builder.py, api/v1/chat.py", "TC-04, TC-14"),
        ("FR-05", "Task Orchestration", "app/agents/orchestrator.py, state.py, requirement_agent.py", "TC-06, TC-12"),
        ("FR-06", "Code & Test Synthesis", "app/agents/code_agent.py, test_agent.py, design_agent.py", "TC-07, TC-15"),
        ("FR-07", "Sandbox Validation", "app/sandbox/sandbox_runner.py, validation_agent.py", "TC-08, TC-13"),
        ("FR-08", "Self-Correction Loop", "app/agents/orchestrator.py (retry budget loop)", "TC-09, TC-17"),
        ("FR-09", "Change Staging & HITL", "app/repository/workspace_manager.py, api/v1/changes.py", "TC-11, TC-16"),
        ("FR-10", "Documentation & Studio", "app/services/documentation_service.py, diagram_agent.py", "TC-19, TC-20")
    ]
    for i, row in enumerate(rtm_data):
        for j, val in enumerate(row):
            t_rtm.rows[i+1].cells[j].text = val
            if j in (0, 3):
                t_rtm.rows[i+1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_table(t_rtm, col_widths=[Inches(0.8), Inches(1.8), Inches(2.5), Inches(1.1)])
    
    # 3.10 Summary
    add_heading_1(doc, "3.10 Summary")
    add_p(doc, "This chapter established the complete requirement specification, architectural feasibility, operational methodology, "
               "formal algorithms, project planning WBS, and requirement traceability matrix for the Engineering Intelligence Hub. "
               "The platform coordinates repository ingestion, line-verified RAG, 12-agent orchestration, sandbox self-correction, and "
               "change staging. Chapter 4 provides the comprehensive system design and implementation details across all system components.", 
          space_after=0)
