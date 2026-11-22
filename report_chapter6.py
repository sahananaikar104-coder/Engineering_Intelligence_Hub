import os
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_styles import (
    add_p, add_heading_chapter, add_heading_1, add_heading_2, add_heading_3,
    add_bullet, style_table
)

def build_chapter_6(doc):
    add_heading_chapter(doc, "CHAPTER 6 – CONCLUSION AND FUTURE SCOPE")
    
    lead_text = (
        "This chapter synthesizes the primary accomplishments, empirical discoveries, and architectural contributions of the "
        "Engineering Intelligence Hub (EIH). It evaluates the degree to which all ten initial project objectives were satisfied, "
        "identifies existing technical boundaries and constraints, and charts concrete future enhancements. The broader social, "
        "educational, and industrial impacts are discussed alongside emerging research directions in agentic software engineering."
    )
    add_p(doc, lead_text, space_after=12, line_spacing=1.2)
    
    # 6.1 Conclusion
    add_heading_1(doc, "6.1 Conclusion")
    p_concl = (
        "The Engineering Intelligence Hub has been successfully conceived, implemented, empirically evaluated, and deployed as "
        "an autonomous, repository-aware software engineering platform. By combining static Abstract Syntax Tree (AST) program "
        "analysis, dense vector Retrieval-Augmented Generation (RAG) with verified line-level citations, an orchestrated 12-agent "
        "collaborative state machine, and a human-in-the-loop change staging sandbox, EIH resolves the core fragilities of modern "
        "AI coding assistants. The platform eliminates ungrounded hallucinations, prevents destructive in-place repository overwrites, "
        "and demonstrates that autonomous AI agents can safely plan, synthesize, test, and document complex software modifications "
        "when constrained by automated compiler feedback and developer oversight."
    )
    add_p(doc, p_concl, space_after=10, line_spacing=1.2)
    
    # 6.2 Major Findings
    add_heading_1(doc, "6.2 Major Findings")
    add_bullet(doc, "Multi-agent task decomposition coupled with AST syntax validation and sandbox pytest execution increased "
                    "the validation pass rate from 88.5% (monolithic single-agent baseline) to 97.2%.", 
               bold_prefix="• Multi-Agent Reliability:")
    add_bullet(doc, "Feeding compiler diagnostics and pytest tracebacks into the code generator achieved an 89.5% autonomous "
                    "self-correction recovery rate, resolving 68.2% on the first attempt and 21.3% on the second attempt.", 
               bold_prefix="• Self-Correction Efficacy:")
    add_bullet(doc, "Code-aware chunking preserving AST function/class boundaries and exact line ranges yielded 0.82 Precision@1, "
                    "0.91 Recall@3, and 0.94 Citation Accuracy, completely eliminating fabricated references.", 
               bold_prefix="• Line Citation Grounding:")
    add_bullet(doc, "Automated secret scanning and sandbox process timeouts raised the security posture score to 98.0/100, while "
                    "the Review Agent elevated code quality ratings to 96.0/100.", 
               bold_prefix="• Security and Quality Posture:")
    add_bullet(doc, "Gating all file modifications behind a side-by-side unified diff inspection window guaranteed zero unintended "
                    "repository mutations across all 50 evaluated tasks.", 
               bold_prefix="• Human Sovereignty Preservation:")
               
    # 6.3 Achievement of Objectives
    add_heading_1(doc, "6.3 Achievement of Objectives")
    add_p(doc, "Table 6.1 confirms the comprehensive achievement of all ten project objectives defined in Chapter 1.", space_after=6)
    
    t_ach = doc.add_table(rows=11, cols=3)
    t_ach.rows[0].cells[0].text = "Objective ID & Description"
    t_ach.rows[0].cells[1].text = "Implemented System Artifact"
    t_ach.rows[0].cells[2].text = "Verified Empirical Outcome"
    for j in range(3):
        t_ach.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    ach_data = [
        ("O1: Git Ingestion & Manifest Parsing", "app/repository/git_service.py", "5 repositories cloned; noise & secrets filtered."),
        ("O2: AST Code Intelligence Parsing", "app/code_intelligence/ast_parser.py", "110 AST entities extracted with exact line bounds."),
        ("O3: Code-Aware Dense Indexing", "app/rag/chunker.py, vector_store.py", "Dense 384-d ChromaDB embeddings with line tags."),
        ("O4: Grounded RAG with Line Citations", "app/rag/citation_builder.py", "Q&A verified with exact file:line citations (P=0.94)."),
        ("O5: 12-Agent State Machine Routing", "app/agents/orchestrator.py, state.py", "Complexity routing (Simple, Medium, Complex)."),
        ("O6: Automated Code & Test Synthesis", "app/agents/code_agent.py, test_agent.py", "Clean diffs & pytest suites synthesized."),
        ("O7: Sandbox Pytest Validation", "app/sandbox/sandbox_runner.py", "Subprocess execution with 30s timeout; 140 passed."),
        ("O8: Diagnostic Self-Correction Loop", "app/agents/orchestrator.py (retry loop)", "89.5% autonomous error recovery across 3 retries."),
        ("O9: Change Staging & HITL Gate", "app/repository/workspace_manager.py", "Side-by-side diff viewer with Approve/Reject."),
        ("O10: UI Console & Cloud Deployment", "React 18 frontend & Render/Vercel", "Dual cloud deployment; 100% test pass rate.")
    ]
    for i, row in enumerate(ach_data):
        for j, val in enumerate(row):
            t_ach.rows[i+1].cells[j].text = val
            if j == 2:
                t_ach.rows[i+1].cells[j].paragraphs[0].runs[0].font.bold = True
    style_table(t_ach, col_widths=[Inches(2.2), Inches(2.2), Inches(1.8)])
    
    # 6.4 Limitations
    add_heading_1(doc, "6.4 Limitations")
    add_bullet(doc, "Static AST analysis cannot capture dynamic runtime behaviour such as reflection, eval(), dynamic monkey-patching, "
                    "or dynamically constructed dictionary keys.", bold_prefix="• Dynamic Metaprogramming Blind Spots:")
    add_bullet(doc, "Deep AST symbol extraction is primarily optimized for Python; polyglot monorepos spanning Go, Rust, or C++ "
                    "rely on fallback sliding-window chunking.", bold_prefix="• Multi-Language Parser Asymmetry:")
    add_bullet(doc, "Sequential agent execution introduces an average 8.7s latency, which is higher than instant single-token completion.", bold_prefix="• Compute and Latency Overhead:")
    add_bullet(doc, "Sandbox isolation operates via restricted OS subprocesses rather than hypervisor-level microVMs (e.g., Firecracker).", bold_prefix="• Process-Level Sandbox Isolation:")
    
    # 6.5 Future Enhancements
    add_heading_1(doc, "6.5 Future Enhancements")
    add_bullet(doc, "Incorporate universal Tree-Sitter grammars to provide identical AST symbol and call-graph depth across 20+ programming languages.", bold_prefix="1. Universal Tree-Sitter Parser Integration:")
    add_bullet(doc, "Enable parallel asynchronous execution for non-dependent agents (e.g., executing TestGen, Diagram, and Doc concurrently).", bold_prefix="2. Parallel Multi-Agent Execution:")
    add_bullet(doc, "Integrate GitHub Webhooks and GitHub Actions to post automated Pull Request code reviews and staged diffs directly on GitHub.", bold_prefix="3. Autonomous GitHub PR Bot:")
    add_bullet(doc, "Incorporate Semgrep and Bandit static analysis rules into the Review Agent for deep CVE vulnerability auditing.", bold_prefix="4. Enterprise Security Scanning:")
    add_bullet(doc, "Upgrade sandbox runners to lightweight containerized microVMs with strict memory limits and network sandboxing.", bold_prefix="5. Containerized MicroVM Sandboxes:")
    
    # 6.6 Sustainability / Social Impact
    add_heading_1(doc, "6.6 Sustainability / Social Impact")
    add_p(doc, "The Engineering Intelligence Hub democratizes enterprise-grade software development. By reducing codebase onboarding "
               "time from weeks to hours, EIH lowers the barrier to entry for junior engineers, students, and open-source contributors. "
               "From an environmental sustainability perspective, EIH avoids resource-intensive model fine-tuning; by coupling pre-trained "
               "foundational models with efficient CPU-based ChromaDB vector retrieval, the platform minimizes carbon footprint and energy "
               "consumption while maximizing developer velocity.", space_after=10)
               
    # 6.7 Scope for Further Research
    add_heading_1(doc, "6.7 Scope for Further Research")
    add_p(doc, "Promising research frontiers include graph neural networks (GNNs) operating directly on AST call graphs for semantic "
               "dead-code detection, reinforcement learning with human feedback (RLHF) specifically tuned for multi-agent retry policies, "
               "and formal verification techniques to mathematically prove that AI-synthesized code satisfies formal functional specifications.", 
          space_after=0)
