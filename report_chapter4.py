import os
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_styles import (
    add_p, add_heading_chapter, add_heading_1, add_heading_2, add_heading_3,
    add_bullet, style_table, add_fig, add_code_listing
)

def build_chapter_4(doc):
    add_heading_chapter(doc, "CHAPTER 4 – SYSTEM DESIGN AND IMPLEMENTATION")
    
    lead_text = (
        "This chapter details the comprehensive architectural design, software modeling, component interactions, database "
        "schemas, data preparation pipelines, implementation frameworks, user interfaces, and deployment configurations of the "
        "Engineering Intelligence Hub (EIH). EIH is architected as an enterprise-grade, three-tier developer platform comprising "
        "a React 18 single-page application, an asynchronous FastAPI backend service, an embedded relational SQLite/PostgreSQL "
        "database, a persistent ChromaDB vector store, and a pluggable LLM provider layer. Complete UML diagrams (use case, "
        "activity, sequence, state), Level 0 and Level 1 Data Flow Diagrams, Entity-Relationship schemas spanning 17 database "
        "tables, module contracts, and genuine application interfaces are presented."
    )
    add_p(doc, lead_text, space_after=12, line_spacing=1.2)
    
    # 4.1 Introduction
    add_heading_1(doc, "4.1 Introduction")
    p1 = (
        "Modern developer tools require modular, loosely coupled architectures capable of handling high-concurrency background "
        "processing, persistent vector storage, and secure sandboxed execution. EIH bridges deterministic static analysis with "
        "generative AI reasoning by organizing capabilities into specialized micro-services. This chapter delineates how each "
        "subsystem—from low-level AST tokenization and Git tree cloning to complex multi-agent state machines and diff visualization—is "
        "designed and realized in production code."
    )
    add_p(doc, p1, space_after=10, line_spacing=1.2)
    
    # 4.2 Detailed System Design
    add_heading_1(doc, "4.2 Detailed System Design")
    add_heading_2(doc, "4.2.1 Overall System Architecture")
    p_arch = (
        "The system follows a three-tier architectural paradigm decoupled into Presentation, Application, and Persistence layers. "
        "The frontend communicates with the backend exclusively via RESTful HTTPS endpoints and Server-Sent Events (SSE) for live "
        "task streaming. The backend encapsulates business logic within isolated service modules, preventing direct coupling "
        "between LLM generation and database persistence. Major platform components and their respective responsibilities are "
        "cataloged in Table 4.1."
    )
    add_p(doc, p_arch, space_after=8, line_spacing=1.2)
    
    t_comp = doc.add_table(rows=11, cols=3)
    t_comp.rows[0].cells[0].text = "Component Name"
    t_comp.rows[0].cells[1].text = "Primary Architectural Responsibility"
    t_comp.rows[0].cells[2].text = "Input Context & Output Deliverable"
    for j in range(3):
        t_comp.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    comp_data = [
        ("React UI Console & Explorer", "Interactive web application; repository management, file tree explorer, diff viewer, RAG chat.", "Input: User actions; Output: Visual UI state, SSE updates."),
        ("FastAPI REST Gateway", "API routing, Pydantic request validation, CORS management, secret redaction, health monitoring.", "Input: HTTP requests; Output: Validated JSON payloads & streams."),
        ("Git Ingestion Engine", "Safe repository cloning, branch tracking, directory tree traversal, commit history extraction.", "Input: GitHub URL / path; Output: Cloned sandbox workspace."),
        ("Python AST Parser", "Deterministic code parsing; function/class definitions, decorators, docstrings, exact line bounds.", "Input: Source code; Output: Structured AST entity catalog."),
        ("Code-Aware Chunker", "Splits source code into atomic chunks respecting class and function bounds with line tags.", "Input: AST entities & source; Output: Line-bounded chunks."),
        ("ChromaDB Vector Store", "Dense semantic indexing using all-MiniLM-L6-v2 embeddings (384 dimensions); cosine search.", "Input: Text chunks; Output: Nearest-neighbor code passages."),
        ("Grounded RAG Chat Engine", "Contextual question answering with AST-backed citation verification (file_path:start_line-end_line).", "Input: Question + chunks; Output: Answer with exact citations."),
        ("12-Agent Orchestrator", "Multi-agent lifecycle management, task complexity routing, and 3-attempt self-correction controller.", "Input: Task requirement; Output: Executed workflow & status."),
        ("Sandbox Test Runner", "Subprocess execution of pytest test suites, AST syntax validation, and traceback capture.", "Input: Workspace files; Output: Pass/fail logs & error traces."),
        ("Relational Persistence", "SQLAlchemy ORM managing 17 relational tables on SQLite or production PostgreSQL.", "Input: Entity models; Output: Persistent ACID relational data.")
    ]
    for i, row in enumerate(comp_data):
        for j, val in enumerate(row):
            t_comp.rows[i+1].cells[j].text = val
            if j == 0:
                t_comp.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_comp, col_widths=[Inches(1.8), Inches(2.4), Inches(2.0)])
    
    # 4.3 System Modelling
    add_heading_1(doc, "4.3 System Modelling")
    add_p(doc, "Formal structural and behavioural models visualize the platform's actors, workflows, data flow, and lifecycles.", space_after=6)
    
    fig_dir = "extracted_figures"
    
    # 4.3.1 Use Case Diagram
    add_heading_2(doc, "4.3.1 Use Case Diagram")
    add_p(doc, "Figure 4.1 models the primary interactions between the human software engineer/architect and the EIH platform boundary. "
               "The developer connects repositories, navigates AST symbols, asks grounded RAG questions, reviews staged diffs, triggers "
               "multi-agent engineering tasks, and synchronizes technical documentation.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_1_usecase.png"), 
            "Figure 4.1: UML Use case diagram of EIH developer platform.", width=Inches(5.5))
            
    # 4.3.2 Data Flow Diagrams
    add_heading_2(doc, "4.3.2 Data Flow Diagrams")
    add_p(doc, "Figure 4.2 presents the Context Diagram (DFD Level 0), illustrating the system as a centralized process exchanging repository "
               "metadata, queries, and diffs between external entities (Developer, GitHub Server, and LLM Providers). Figure 4.3 decomposes the "
               "system into Level 1 subprocesses: Ingestion (1.0), AST Parsing (2.0), Vector Retrieval (3.0), Multi-Agent Orchestration (4.0), "
               "Sandbox Validation (5.0), and Change Staging (6.0).", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_2_dfd0.png"), 
            "Figure 4.2: Data flow diagram, level 0 (context diagram).", width=Inches(5.6))
    add_fig(doc, os.path.join(fig_dir, "fig4_3_dfd1.png"), 
            "Figure 4.3: Data flow diagram, level 1 (detailed ingestion and task workflow).", width=Inches(5.8))
            
    # 4.3.3 Activity Diagram
    add_heading_2(doc, "4.3.3 Activity Diagram")
    add_p(doc, "Figure 4.4 details the dynamic control flow during an engineering task. The Orchestrator receives a natural-language "
               "prompt, decomposes the plan, coordinates CodeGen and TestGen, and invokes the Validation Agent. If tests fail, the diagnostic "
               "branch loops back to CodeGen; if tests pass, the changes advance to the Human Approval Gate.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_4_activity.png"), 
            "Figure 4.4: UML Activity diagram of task execution, validation and retry loop.", width=Inches(5.4))
            
    # 4.3.4 Sequence Diagram
    add_heading_2(doc, "4.3.4 Sequence Diagram")
    add_p(doc, "Figure 4.5 captures the asynchronous message sequence of a Grounded RAG Query. The user's query is embedded, ChromaDB returns "
               "top-k code chunks, AST symbols verify physical line existence, the LLM generates the response, and the Citation Builder tags "
               "verified clickable spans before returning JSON.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_5_sequence.png"), 
            "Figure 4.5: UML Sequence diagram of a grounded RAG query and citation linking.", width=Inches(5.6))
            
    # 4.3.5 State Diagram
    add_heading_2(doc, "4.3.5 State Transition Diagram")
    add_p(doc, "Figure 4.6 formalizes the lifecycle states of an Engineering Task: PENDING -> PLANNING -> GENERATING -> VALIDATING. "
               "From VALIDATING, the task transitions to STAGED (if assertions pass), loops back to GENERATING (if retry count <= 3), "
               "or transitions to FAILED (if max retries are exhausted). Once STAGED, human review leads to APPROVED or REJECTED.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_6_state.png"), 
            "Figure 4.6: State transition diagram of an autonomous engineering task lifecycle.", width=Inches(5.6))
            
    # 4.3.6 ER Diagram and Database Design
    add_heading_2(doc, "4.3.6 Entity–Relationship Diagram and Database Schema")
    add_p(doc, "The relational persistence schema is organized into 17 tables in SQLite (development) and PostgreSQL (production), "
               "as detailed in Table 4.2 and illustrated in the ER diagram (Figure 4.7).", space_after=8)
               
    t_db = doc.add_table(rows=18, cols=3)
    t_db.rows[0].cells[0].text = "Table Name"
    t_db.rows[0].cells[1].text = "Purpose & Domain Entities"
    t_db.rows[0].cells[2].text = "Key Columns & Foreign Keys"
    for j in range(3):
        t_db.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    db_tables_data = [
        ("repositories", "Tracks ingested Git repositories and connection parameters.", "id (PK), name, url, default_branch, local_path, created_at"),
        ("repository_files", "Hierarchical file records with path and language info.", "id (PK), repository_id (FK), file_path, file_size, language"),
        ("dependencies", "Parsed packages from package.json, requirements.txt, etc.", "id (PK), repository_id (FK), package_name, version, manifest_file"),
        ("code_entities", "AST extracted classes, functions, docstrings, and lines.", "id (PK), repository_id (FK), name, entity_type, start_line, end_line"),
        ("conversations", "Chat sessions tied to specific repository contexts.", "id (PK), repository_id (FK), title, created_at, updated_at"),
        ("messages", "Individual Q&A chat messages with citation metadata.", "id (PK), conversation_id (FK), sender, content, citations_json"),
        ("engineering_tasks", "High-level feature/bugfix requests and status states.", "id (PK), repository_id (FK), requirement, status, complexity, retry_count"),
        ("task_steps", "Individual execution steps decomposed by the Orchestrator.", "id (PK), task_id (FK), step_number, agent_name, status, details"),
        ("agent_logs", "Granular execution trace logs emitted by specialized agents.", "id (PK), task_id (FK), agent_name, level, message, timestamp"),
        ("code_changes", "Proposed file additions, modifications, and unified diffs.", "id (PK), task_id (FK), file_path, change_type, diff_content, staged_path"),
        ("validation_results", "Sandbox test execution outcomes and AST syntax checks.", "id (PK), task_id (FK), passed, error_summary, test_output"),
        ("review_results", "Automated code quality and security audit scores (0–100).", "id (PK), task_id (FK), maintainability_score, security_score, issues_json"),
        ("task_diagrams", "Generated Mermaid architecture, sequence, and flow diagrams.", "id (PK), task_id (FK), diagram_type, mermaid_code, svg_export"),
        ("documentation_items", "Tracks 13 technical documentation types per repository.", "id (PK), repository_id (FK), doc_type, title, content_markdown, is_outdated"),
        ("documentation_gaps", "Identifies missing or undocumented modules/APIs.", "id (PK), repository_id (FK), file_path, gap_description, severity"),
        ("documentation_versions", "Revision history for technical documents with diffs.", "id (PK), doc_item_id (FK), version_number, content, changed_at"),
        ("terminal_sessions", "Ephemeral interactive terminal sandbox environments.", "id (PK), repository_id (FK), session_token, is_active, last_heartbeat")
    ]
    for i, row in enumerate(db_tables_data):
        for j, val in enumerate(row):
            t_db.rows[i+1].cells[j].text = val
            if j == 0:
                t_db.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_db, col_widths=[Inches(1.8), Inches(2.2), Inches(2.2)])
    
    add_fig(doc, os.path.join(fig_dir, "fig4_7_er.png"), 
            "Figure 4.7: Entity–relationship diagram of the 17 platform database tables.", width=Inches(5.8))
            
    # 4.3.7 Deployment Diagram
    add_heading_2(doc, "4.3.7 Deployment Diagram")
    add_p(doc, "Figure 4.8 illustrates the cloud deployment architecture. The React frontend is deployed as a static build on Vercel's global CDN, "
               "communicating over HTTPS with the FastAPI ASGI web service hosted on Render. Render manages the Python 3.11 runtime, SQLite/PostgreSQL "
               "databases, and local ChromaDB vector instances. External calls connect to GitHub API and LLM providers.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_8_deployment.png"), 
            "Figure 4.8: Cloud deployment architecture: Render backend and Vercel frontend.", width=Inches(5.6))
            
    # 4.3.8 Module Interaction Diagram
    add_heading_2(doc, "4.3.8 Module Interaction Diagram")
    add_p(doc, "Figure 4.9 maps the inter-agent collaboration topology. The Orchestrator manages state transitions across Planning (Requirement, "
               "Repo, CodeUnderstanding), Build (Design, CodeGen, TestGen), Verify (Validation, Review), and Sync (Doc, Diagram, ChangeImpact) clusters.", space_after=8)
    add_fig(doc, os.path.join(fig_dir, "fig4_9_agents.png"), 
            "Figure 4.9: Interaction architecture between the 12 specialized autonomous agents.", width=Inches(5.6))
            
    # 4.4 Module Design
    add_heading_1(doc, "4.4 Module Design")
    add_p(doc, "Each core module in EIH is engineered with strict encapsulation, defined inputs, processing logic, and outputs:", space_after=6)
    
    add_heading_2(doc, "4.4.1 Ingestion & AST Parsing Module")
    add_bullet(doc, "Clones repositories, filters ignored files (.git, node_modules, .venv, binaries), and parses AST into symbols with line coordinates.", bold_prefix="• Functionality:")
    add_bullet(doc, "Input: GitHub URL or local path. Output: Relational file records, dependency catalog, and AST symbol entities.", bold_prefix="• Contract:")
    add_bullet(doc, "Source: app/repository/git_service.py, app/code_intelligence/ast_parser.py.", bold_prefix="• Implementation:")
    
    add_heading_2(doc, "4.4.2 Grounded RAG Chat Module")
    add_bullet(doc, "Embeds queries, performs hybrid vector retrieval in ChromaDB, prompts LLM with line context, and verifies citation spans.", bold_prefix="• Functionality:")
    add_bullet(doc, "Input: Query string and repository ID. Output: Answer markdown and verified line citations (file_path:start_line-end_line).", bold_prefix="• Contract:")
    add_bullet(doc, "Source: app/rag/retriever.py, app/rag/citation_builder.py, app/api/v1/chat.py.", bold_prefix="• Implementation:")
    
    add_heading_2(doc, "4.4.3 Autonomous Task Orchestrator & Agents")
    add_bullet(doc, "Routes task complexity, decomposes engineering tasks, generates code/tests, executes sandbox runs, and manages self-correction.", bold_prefix="• Functionality:")
    add_bullet(doc, "Input: Task requirement string. Output: Staged unified diffs, pytest suites, SDD documentation, and security audit scores.", bold_prefix="• Contract:")
    add_bullet(doc, "Source: app/agents/orchestrator.py, state.py, code_agent.py, test_agent.py, validation_agent.py.", bold_prefix="• Implementation:")
    
    add_heading_2(doc, "4.4.4 Change Staging & Human Approval Gate")
    add_bullet(doc, "Maintains isolated staging workspaces, renders side-by-side diffs, and gates codebase writes behind developer approval.", bold_prefix="• Functionality:")
    add_bullet(doc, "Input: Task ID and approval action (approve/reject/revise). Output: Applied commit or discarded workspace.", bold_prefix="• Contract:")
    add_bullet(doc, "Source: app/repository/workspace_manager.py, app/api/v1/changes.py.", bold_prefix="• Implementation:")
    
    # 4.5 Dataset and Data Preparation
    add_heading_1(doc, "4.5 Dataset and Data Preparation")
    add_p(doc, "Figure 4.10 depicts the data acquisition and preparation pipeline. The repository source code serves as the primary dataset. "
               "Five diverse open-source codebases were analyzed, yielding 110 AST entities and 100 dependencies as summarized in Table 4.3. "
               "Figure 4.11 graphs the distribution of extracted classes, functions, and endpoints across the five codebases.", space_after=8)
               
    add_fig(doc, os.path.join(fig_dir, "fig4_10_pipeline.png"), 
            "Figure 4.10: Code ingestion, AST parsing, and knowledge preparation pipeline.", width=Inches(5.8))
            
    t_ds = doc.add_table(rows=7, cols=5)
    t_ds.rows[0].cells[0].text = "Repository Name"
    t_ds.rows[0].cells[1].text = "Primary Language"
    t_ds.rows[0].cells[2].text = "Files"
    t_ds.rows[0].cells[3].text = "AST Entities"
    t_ds.rows[0].cells[4].text = "Dependencies"
    for j in range(5):
        t_ds.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    ds_data = [
        ("FastAPI Microservice", "Python 3.11", "34", "46", "24"),
        ("Authentication Service", "Python / JWT", "18", "26", "18"),
        ("Task Management API", "Python / SQLAlchemy", "28", "40", "22"),
        ("React Frontend", "TypeScript / React", "22", "22", "26"),
        ("Utility Tools Repository", "Python", "15", "16", "10"),
        ("Total Aggregated Dataset", "Multi-Language", "117", "110", "100")
    ]
    for i, row in enumerate(ds_data):
        for j, val in enumerate(row):
            t_ds.rows[i+1].cells[j].text = val
            if j in (0, 1):
                pass
            else:
                t_ds.rows[i+1].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            if i == len(ds_data) - 1:
                t_ds.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_ds, col_widths=[Inches(2.0), Inches(1.3), Inches(0.8), Inches(1.0), Inches(1.1)])
    
    add_fig(doc, os.path.join(fig_dir, "fig4_11_entity_dist.png"), 
            "Figure 4.11: Distribution of AST code entities extracted across 5 benchmark repositories.", width=Inches(5.8))
            
    # 4.6 Model Development and Algorithm Design
    add_heading_1(doc, "4.6 Model Development and Algorithm Design")
    add_p(doc, "Table 4.4 details the retrieval model and vector storage configuration. Figure 4.12 illustrates the Grounded RAG pipeline.", space_after=6)
    
    t_mdl = doc.add_table(rows=7, cols=2)
    t_mdl.rows[0].cells[0].text = "Configuration Parameter"
    t_mdl.rows[0].cells[1].text = "Selected Engineering Setting & Specification"
    for j in range(2):
        t_mdl.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    mdl_data = [
        ("Vector Embedding Model", "all-MiniLM-L6-v2 (Sentence-Transformers), 384 dense vector dimensions."),
        ("Vector Store Engine", "ChromaDB 0.4.24 with persistent disk storage in chroma_db/ directory."),
        ("Distance Metric", "Cosine Similarity (threshold >= 0.65 for chunk inclusion in prompt context)."),
        ("Code-Aware Chunk Size", "Target 50 lines with 10-line overlap; atomic preservation of AST functions/classes."),
        ("Pluggable LLM Providers", "OpenAI (gpt-4o), Groq (llama-3.3-70b-versatile), Ollama (local), Anthropic, Mock Provider."),
        ("LLM Generation Hyperparameters", "Temperature = 0.2 (low randomness for deterministic syntax), Top_p = 0.95, Max_tokens = 2048.")
    ]
    for i, row in enumerate(mdl_data):
        for j, val in enumerate(row):
            t_mdl.rows[i+1].cells[j].text = val
    style_table(t_mdl, col_widths=[Inches(2.2), Inches(4.0)])
    
    add_fig(doc, os.path.join(fig_dir, "fig4_12_rag_pipeline.png"), 
            "Figure 4.12: Grounded RAG architecture with AST-driven line citation verification.", width=Inches(5.6))
            
    # 4.7 System Implementation
    add_heading_1(doc, "4.7 System Implementation")
    add_p(doc, "Implementation technologies and verified framework versions are listed in Table 4.5. Table 4.6 catalogs the REST API routers.", space_after=6)
    
    t_tech = doc.add_table(rows=9, cols=3)
    t_tech.rows[0].cells[0].text = "Technology / Layer"
    t_tech.rows[0].cells[1].text = "Installed Version"
    t_tech.rows[0].cells[2].text = "Functional Purpose in System"
    for j in range(3):
        t_tech.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    tech_data = [
        ("Python Runtime", "3.11.9", "Core backend runtime, asyncio event loop, and typing."),
        ("FastAPI / Uvicorn", "0.110.0 / 0.28.0", "Asynchronous ASGI web application framework and HTTP server."),
        ("SQLAlchemy (aiosqlite)", "2.0.28", "Asynchronous Object-Relational Mapping for relational metadata."),
        ("ChromaDB", "0.4.24", "High-performance embedded vector database for dense embeddings."),
        ("Sentence-Transformers", "all-MiniLM-L6-v2", "Local dense vector embedding model generating 384-d vectors."),
        ("React / React DOM", "18.2.0", "Frontend user interface component library and virtual DOM."),
        ("Vite Build Tool", "5.1.6", "Frontend bundle tool and fast development server."),
        ("Tailwind CSS", "3.4.1", "Utility-first responsive CSS styling with dark-mode theme.")
    ]
    for i, row in enumerate(tech_data):
        for j, val in enumerate(row):
            t_tech.rows[i+1].cells[j].text = val
            if j == 0:
                t_tech.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_tech, col_widths=[Inches(1.8), Inches(1.4), Inches(3.0)])
    
    t_api = doc.add_table(rows=8, cols=3)
    t_api.rows[0].cells[0].text = "Router Group"
    t_api.rows[0].cells[1].text = "Prefix & Endpoints"
    t_api.rows[0].cells[2].text = "Example Operations & Methods"
    for j in range(3):
        t_api.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    api_data = [
        ("Repositories", "/api/v1/repositories", "POST /connect, GET /list, GET /{id}/tree, GET /{id}/file"),
        ("Intelligence", "/api/v1/intelligence", "GET /{id}/symbols, GET /{id}/dependencies, GET /{id}/topology"),
        ("Grounded Chat", "/api/v1/chat", "POST /, GET /conversations, GET /{id}/messages"),
        ("Task Workflow", "/api/v1/tasks", "POST /create, POST /{id}/execute, GET /{id}/status, SSE /stream"),
        ("Change Staging", "/api/v1/changes", "GET /{id}/diff, POST /{id}/approve, POST /{id}/reject, POST /{id}/revise"),
        ("Documentation", "/api/v1/documentation", "GET /{id}/docs, POST /{id}/generate, GET /{id}/gaps"),
        ("Terminal Sandbox", "/api/v1/terminal", "POST /session/start, POST /command/run, POST /session/terminate")
    ]
    for i, row in enumerate(api_data):
        for j, val in enumerate(row):
            t_api.rows[i+1].cells[j].text = val
            if j == 0:
                t_api.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_api, col_widths=[Inches(1.8), Inches(2.0), Inches(2.4)])
    
    # 4.8 User Interface / Application
    add_heading_1(doc, "4.8 User Interface / Application")
    add_p(doc, "The EIH frontend is a responsive single-page application engineered in React 18 with TypeScript and Tailwind CSS. "
               "Genuine full-resolution application screenshots illustrate the platform views in Figures 4.13 to 4.21.", space_after=8)
               
    add_fig(doc, os.path.join(fig_dir, "screenshot_dashboard.png"), 
            "Figure 4.13: Executive Dashboard showing connected repositories and system health.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_explorer.png"), 
            "Figure 4.14: Code Explorer with recursive file tree and AST symbol catalog.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_chat.png"), 
            "Figure 4.15: Grounded RAG Chat with verified line citations and code snippets.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_architecture.png"), 
            "Figure 4.16: Interactive Architecture Studio with dynamic Mermaid diagrams.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_tasks.png"), 
            "Figure 4.17: Engineering Task Execution Console with agent status steps.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_repositories.png"), 
            "Figure 4.18: Repositories management console showing connected codebases.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_documentation.png"), 
            "Figure 4.19: Documentation Studio with 13 technical document generation options.", width=Inches(5.8))
    add_fig(doc, os.path.join(fig_dir, "screenshot_mobile_dashboard.png"), 
            "Figure 4.20: Responsive mobile phone layout of the Executive Dashboard (390 x 844 px).", width=Inches(3.2))
    add_fig(doc, os.path.join(fig_dir, "screenshot_mobile_explorer.png"), 
            "Figure 4.21: Responsive mobile phone layout of the Code Explorer (390 x 844 px).", width=Inches(3.2))
            
    # 4.9 Deployment
    add_heading_1(doc, "4.9 Deployment")
    add_p(doc, "The platform is configured for automated cloud deployment: the FastAPI backend runs on Render as a web service "
               "(Python 3.11 with render.yaml), while the React frontend is deployed on Vercel's global CDN. Environment variables "
               "manage CORS origins, database URLs, and API keys securely.", space_after=8)
               
    # 4.10 Summary
    add_heading_1(doc, "4.10 Summary")
    add_p(doc, "This chapter presented the comprehensive design and implementation of the Engineering Intelligence Hub. The three-tier "
               "architecture integrates AST code intelligence, dense ChromaDB vector search, 12 specialized autonomous agents, automated "
               "sandbox validation, and human-in-the-loop staging. Chapter 5 details the experimental evaluation, 140-test automated suite, "
               "performance profiling, and comparative benchmarks.", space_after=0)
