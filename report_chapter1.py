import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_styles import (
    add_p, add_heading_chapter, add_heading_1, add_heading_2, add_heading_3,
    add_bullet, style_table
)

def build_chapter_1(doc):
    add_heading_chapter(doc, "CHAPTER 1 – INTRODUCTION")
    
    lead_text = (
        "Modern enterprise software engineering environments continuously generate vast, heterogeneous codebases "
        "comprising thousands of source files, complex class inheritance graphs, multi-tier microservices, database schemas, "
        "evolving API contracts, and external dependency manifests. Navigating, understanding, and modifying such massive "
        "repositories imposes an immense cognitive overhead on developers, particularly when onboarding to unfamiliar "
        "projects or implementing cross-cutting engineering requirements. While modern Artificial Intelligence (AI) and "
        "Large Language Models (LLMs) have demonstrated remarkable capabilities in localized code synthesis, existing "
        "AI coding assistants typically treat source code as disconnected textual prompts. They lack holistic awareness of "
        "repository architecture, cannot trace multi-hop call graphs, frequently hallucinate dependencies, and operate without "
        "systematic validation or execution safeguards."
    )
    lead_text2 = (
        "The Engineering Intelligence Hub (EIH) is an autonomous, agentic software engineering platform engineered to overcome "
        "these fundamental limitations. EIH bridges the gap between repository comprehension and automated software development "
        "by unifying Abstract Syntax Tree (AST) code intelligence, dense vector Retrieval-Augmented Generation (RAG) with "
        "verified line-level citations, and an orchestrated state machine of 12 specialized autonomous agents. The platform "
        "governs every software change through an automated self-correction loop and a human-in-the-loop change staging "
        "environment, guaranteeing that no AI-generated code is ever applied to a production codebase without empirical "
        "sandbox validation and explicit developer review."
    )
    add_p(doc, lead_text, space_after=8, line_spacing=1.2)
    add_p(doc, lead_text2, space_after=12, line_spacing=1.2)
    
    # 1.1 Background
    add_heading_1(doc, "1.1 Background")
    p1 = (
        "Contemporary software development organizations rely on distributed version control systems, primarily GitHub, to "
        "host complex software projects. Within these repositories, domain logic is rarely confined to single files; rather, it is "
        "distributed across modular components, package boundaries, configuration manifests, and interface definitions. "
        "When developers are tasked with implementing a new feature or debugging a regression, they must first build a mental "
        "model of the repository topology by manually tracing functions, cross-referencing documentation, and identifying "
        "dependencies. Empirical industry studies indicate that software engineers spend up to 60% of their working hours reading, "
        "tracing, and comprehending existing code rather than writing new logic."
    )
    p2 = (
        "The recent emergence of generative AI coding assistants, such as GitHub Copilot and conversational LLM interfaces, has "
        "accelerated single-file function completion. However, when deployed against enterprise-scale repositories, these tools "
        "encounter severe architectural barriers. First, context-window limitations prevent feeding an entire codebase into an LLM "
        "prompt, leading to the well-documented 'Lost in the Middle' phenomenon where critical imports and schemas are omitted. "
        "Second, standard textual chunking mechanisms arbitrarily sever functions and control blocks, destroying AST relationships. "
        "Third, existing assistants generate code in an ungrounded vacuum without running automated test suites or syntax parsers, "
        "forcing human developers to perform laborious debugging."
    )
    p3 = (
        "To address this systemic challenge, the Engineering Intelligence Hub combines deterministic static analysis with "
        "probabilistic language reasoning. By parsing Abstract Syntax Trees directly into structured relational schemas and "
        "generating dense semantic embeddings in ChromaDB, EIH establishes a queryable knowledge graph of the repository. "
        "This structural ground truth empowers developers to query codebases in natural language and receive verified citations, "
        "while enabling autonomous agents to plan, implement, and validate multi-file code modifications."
    )
    add_p(doc, p1, space_after=8, line_spacing=1.2)
    add_p(doc, p2, space_after=8, line_spacing=1.2)
    add_p(doc, p3, space_after=12, line_spacing=1.2)
    
    # 1.2 Problem Statement
    add_heading_1(doc, "1.2 Problem Statement")
    prob_intro = (
        "Modern web and cloud application codebases generate intricate networks of source files, API endpoints, and dependency "
        "trees that make manual codebase comprehension slow, error-prone, and cognitively exhausting. Existing generative AI tools "
        "fail to provide dependable repository-wide assistance due to four specific, unresolved technical difficulties:"
    )
    add_p(doc, prob_intro, space_after=6, line_spacing=1.2)
    
    add_bullet(doc, "Arbitrary chunking boundaries and context loss sever function headers from their bodies, resulting in "
                    "hallucinated imports, nonexistent method calls, and inaccurate architectural explanations.", 
               bold_prefix="1. Semantic Disconnect and Hallucination:")
    add_bullet(doc, "Monolithic LLM prompts attempt to parse, architect, code, test, and audit in a single ungrounded pass, "
                    "leading to compounding errors and shallow implementations that break repository conventions.", 
               bold_prefix="2. Monolithic Single-Agent Fragility:")
    add_bullet(doc, "Existing AI tools generate code directly without executing test suites or verifying AST syntax, leaving "
                    "syntax errors, type mismatches, and broken test assertions for developers to manually resolve.", 
               bold_prefix="3. Absence of Automated Verification and Self-Correction:")
    add_bullet(doc, "Automated assistants frequently propose in-place file overwrites without isolating modifications in an "
                    "ephemeral workspace or presenting side-by-side diff inspections, risking catastrophic repository corruption.", 
               bold_prefix="4. Uncontrolled In-Place Code Mutations:")
               
    prob_summary = (
        "Therefore, there is an urgent academic and industrial need for an integrated, repository-aware software engineering "
        "platform that extracts verified AST code intelligence, grounds natural-language queries with exact line citations, "
        "orchestrates multi-agent execution with self-correction, and enforces human-in-the-loop review gating before applying changes."
    )
    add_p(doc, prob_summary, space_after=12, line_spacing=1.2)
    
    # 1.3 Motivation
    add_heading_1(doc, "1.3 Motivation")
    mot1 = (
        "The primary motivation behind the Engineering Intelligence Hub is to dramatically accelerate developer productivity and "
        "code quality while eliminating the fragility of ungrounded generative AI. In large engineering teams and open-source "
        "ecosystems, onboarding a junior developer to a legacy repository can require several weeks of code walkthroughs and "
        "architectural explanations. By providing an interactive conversational interface grounded in real AST entities and line-level "
        "citations, EIH transforms repositories into self-explaining, queryable knowledge systems."
    )
    mot2 = (
        "A secondary motivation is reliability and architectural safety. Automated code generators cannot be trusted in production "
        "environments without rigorous verification. EIH introduces a multi-tier safety architecture: every generated modification "
        "is subjected to AST syntax verification, dependency validation, and isolated sandbox pytest execution. If an assertion fails, "
        "the diagnostic traceback is fed back into a self-correction loop rather than burdening the user. Furthermore, by staging "
        "all changes in an isolated workspace with side-by-side unified diffs, the platform preserves human developer sovereignty."
    )
    mot3 = (
        "Finally, from an academic and engineering research perspective, this project provides a rigorous testbed to evaluate the "
        "empirical trade-offs between monolithic single-agent systems and collaborative multi-agent state machines, measuring "
        "tangible metrics such as validation pass rate, code maintainability, security compliance, and self-correction efficacy."
    )
    add_p(doc, mot1, space_after=8, line_spacing=1.2)
    add_p(doc, mot2, space_after=8, line_spacing=1.2)
    add_p(doc, mot3, space_after=12, line_spacing=1.2)
    
    # 1.4 Objectives
    add_heading_1(doc, "1.4 Objectives")
    add_p(doc, "The formal engineering objectives of the Engineering Intelligence Hub are delineated in Table 1.1. Each objective "
               "is systematically traced to its design and implementation in Chapters 3 and 4 and to its experimental verification in Chapter 5.", 
          space_after=8)
          
    t_obj = doc.add_table(rows=11, cols=2)
    t_obj.rows[0].cells[0].text = "ID"
    t_obj.rows[0].cells[1].text = "Project Objective"
    t_obj.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_obj.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    objs = [
        ("O1", "Develop an automated repository ingestion engine that clones Git repositories, parses manifests, and excludes sensitive tokens/secrets."),
        ("O2", "Implement an AST code intelligence parser to extract classes, functions, decorators, docstrings, imports, and exact line numbers."),
        ("O3", "Build a code-aware chunking and dense vector indexing pipeline using ChromaDB and all-MiniLM-L6-v2 embeddings."),
        ("O4", "Implement a Grounded RAG Chat engine that answers architectural queries and verifies citations at exact line level (file_path:start_line-end_line)."),
        ("O5", "Construct an autonomous multi-agent state machine orchestrating 12 specialized agents based on task complexity routing."),
        ("O6", "Develop an automated sandbox validation runner to execute AST syntax checks, import resolution, and pytest test suites."),
        ("O7", "Engineer an iterative self-correction loop that parses test/validation diagnostics and re-prompts code generation up to 3 retries."),
        ("O8", "Implement an isolated change-staging workspace with side-by-side and unified diff viewers and a human approval gate."),
        ("O9", "Provide automated documentation intelligence supporting 13 technical document types, Mermaid diagram generation, and drift detection."),
        ("O10", "Design a responsive React 18 console with dark-mode styling, integrated terminal, automated test verification, and dual cloud deployment.")
    ]
    for i, (oid, odesc) in enumerate(objs):
        t_obj.rows[i+1].cells[0].text = oid
        t_obj.rows[i+1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        t_obj.rows[i+1].cells[1].text = odesc
    style_table(t_obj, col_widths=[Inches(0.8), Inches(5.2)])
    
    # 1.5 Scope of the Project
    add_heading_1(doc, "1.5 Scope of the Project")
    add_p(doc, "The scope defines the functional boundaries, operational targets, and constraints of the Engineering Intelligence Hub.", space_after=6)
    
    add_heading_2(doc, "1.5.1 Included Scope")
    add_bullet(doc, "Ingestion of public and private GitHub repositories and local project directories via GitPython.", bold_prefix="• Repository Ingestion:")
    add_bullet(doc, "Deep AST parsing for Python codebases and generic manifest scanning for package.json, requirements.txt, Cargo.toml, and Dockerfiles.", bold_prefix="• AST & Code Intelligence:")
    add_bullet(doc, "Semantic vector retrieval using ChromaDB with exact line-range preservation and citation verification.", bold_prefix="• Grounded RAG:")
    add_bullet(doc, "12 specialized autonomous agents coordinated by an Orchestrator state machine with complexity routing.", bold_prefix="• Multi-Agent Coordination:")
    add_bullet(doc, "Safe subprocess execution of test suites, AST syntax validation, and automated traceback extraction.", bold_prefix="• Sandbox Validation:")
    add_bullet(doc, "Three-iteration diagnostic retry loop feeding error logs back to the code generator.", bold_prefix="• Self-Correction Engine:")
    add_bullet(doc, "Zero in-place file mutation; all changes are staged as unified git diffs awaiting human review.", bold_prefix="• Human-in-the-Loop Gate:")
    add_bullet(doc, "Generation of 13 technical document types and live rendering of dynamic Mermaid architecture diagrams.", bold_prefix="• Documentation & Diagrams:")
    add_bullet(doc, "Comprehensive React 18 dashboard, code explorer, diff viewer, terminal, and dual deployment on Render and Vercel.", bold_prefix="• User Interface & Cloud:")
    
    add_heading_2(doc, "1.5.2 Limitations of Scope")
    add_bullet(doc, "Dynamic runtime behaviour, such as metaprogramming, reflection, and dynamically evaluated strings (eval/exec), cannot be fully captured by static AST parsing.", bold_prefix="• Dynamic Analysis Limitation:")
    add_bullet(doc, "While manifest dependencies are analyzed across multiple languages, deep AST symbol and call-graph parsing is primarily optimized for Python and TypeScript/JavaScript.", bold_prefix="• Language Specialization:")
    add_bullet(doc, "The platform executes sandbox tests within isolated OS process boundaries; containerized microVM sandboxes (e.g., Firecracker) are reserved for future enterprise releases.", bold_prefix="• Sandbox Isolation Level:")
    add_bullet(doc, "Cloud deployment on free hosting tiers (Render/Vercel) incurs cold-start latency after idle periods.", bold_prefix="• Hosting Tier Constraints:")
    
    # 1.6 Innovation and Key Contributions
    add_heading_1(doc, "1.6 Innovation and Key Contributions")
    add_bullet(doc, "Unlike conventional chunkers that split files by arbitrary character counts, EIH chunks along AST function and class boundaries, preserving exact start and end line coordinates to eliminate bisected code.", bold_prefix="• AST-Preserving Code Chunking:")
    add_bullet(doc, "Every claim in AI chat and generated documentation is cross-referenced against extracted AST entities, generating interactive citations that allow developers to inspect the exact lines of code cited.", bold_prefix="• Verifiable Line Citations:")
    add_bullet(doc, "Decomposes monolithic code generation into specialized collaborative roles—Requirement, Design, Code, Test, Validation, and Review—drastically improving code maintainability.", bold_prefix="• 12-Agent State Machine:")
    add_bullet(doc, "When sandbox tests or syntax checks fail, the platform automatically captures diagnostic tracebacks and re-executes code synthesis, achieving an 89.5% autonomous error recovery rate.", bold_prefix="• Automated Self-Correction Loop:")
    add_bullet(doc, "Guarantees zero unintended codebase mutation by isolating all proposed changes in ephemeral workspaces until explicit human approval.", bold_prefix="• Zero-Overwrite Staging Sandbox:")
    add_bullet(doc, "Establishes a standardized empirical testbed comparing single-agent vs multi-agent software engineering workflows across pass rates, latency, and quality scores.", bold_prefix="• Empirical Research Benchmarking:")
    
    # 1.7 Social and Industrial Relevance
    add_heading_1(doc, "1.7 Social and Industrial Relevance")
    add_p(doc, "The software industry faces an acute talent shortage and skyrocketing onboarding costs as legacy codebases expand. "
               "Software engineers spend an estimated 30–60% of their time reading unfamiliar code, deciphering undocumented APIs, and "
               "debugging unintended side effects. EIH delivers immediate industrial value by automating codebase exploration, providing "
               "instant line-verified answers, and synthesizing compliant code with automated test suites. In open-source and educational "
               "contexts, the platform lowers entry barriers for novice programmers, enabling students and junior engineers to contribute "
               "meaningfully to complex repositories without requiring senior developer hand-holding.", space_after=10)
               
    # 1.8 Government and National Priority Alignment
    add_heading_1(doc, "1.8 Government and National Priority Alignment")
    add_p(doc, "The Engineering Intelligence Hub aligns directly with national technological initiatives including Digital India, "
               "the National AI Strategy (NITI Aayog), and the National Programme on AI for Software Productivity. By empowering "
               "domestic software development firms with sovereign, locally executable AI tools that support open-source LLMs (via Ollama "
               "and Groq) and preserve code privacy without transmitting proprietary source code to foreign clouds, EIH fosters self-reliant "
               "digital infrastructure (Atmanirbhar Bharat in software tooling).", space_after=10)
               
    # 1.9 SDG Alignment
    add_heading_1(doc, "1.9 SDG Alignment")
    add_p(doc, "The project directly contributes to the United Nations Sustainable Development Goals (SDGs), as summarized in Table 1.2.", space_after=6)
    
    t_sdg = doc.add_table(rows=4, cols=3)
    t_sdg.rows[0].cells[0].text = "Selected SDG"
    t_sdg.rows[0].cells[1].text = "Relevance to the Project"
    t_sdg.rows[0].cells[2].text = "Expected Contribution / Impact"
    for j in range(3):
        t_sdg.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    sdg_data = [
        ("SDG 9: Industry, Innovation and Infrastructure", 
         "Develops resilient, autonomous developer tools for digital software infrastructure.", 
         "Enhances software engineering productivity, reduces technical debt, and improves code reliability."),
        ("SDG 8: Decent Work and Economic Growth", 
         "Automates repetitive code scaffolding, test authoring, and documentation tasks.", 
         "Reduces cognitive developer burnout and elevates human engineers to high-value architectural problem-solving."),
        ("SDG 4: Quality Education", 
         "Provides an explainable, repository-aware interactive learning and code exploration platform.", 
         "Enables students and novice developers to understand complex open-source software architectures safely.")
    ]
    for i, row in enumerate(sdg_data):
        for j, val in enumerate(row):
            t_sdg.rows[i+1].cells[j].text = val
    style_table(t_sdg, col_widths=[Inches(1.8), Inches(2.2), Inches(2.0)])
    
    # 1.10 Organization of the Report
    add_heading_1(doc, "1.10 Organization of the Report")
    add_p(doc, "The remainder of this report is organized into five structured chapters: Chapter 2 presents a rigorous literature survey "
               "of multi-agent software engineering, AST code representations, RAG architectures, and sandbox security, identifying the key "
               "research gaps. Chapter 3 specifies functional and non-functional requirements, feasibility studies, the proposed 3-stream "
               "methodology, algorithms, work breakdown structure, and requirement traceability. Chapter 4 elaborates the complete system design, "
               "UML diagrams (use case, activity, sequence, state), DFDs, ER schema (17 tables), module interactions, data preparation, UI screens, "
               "and cloud deployment. Chapter 5 details experimental verification, the 140-test automated suite, empirical benchmarks comparing "
               "single vs multi-agent execution, latency profiling, error analysis, and 16 genuine application screenshots. Chapter 6 concludes "
               "the report with major findings, limitations, future scope, and sustainability impact, followed by References and Appendices A through F.", 
          space_after=0)
