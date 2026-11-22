import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from report_styles import (
    add_p, add_heading_chapter, add_heading_1, add_heading_2, add_heading_3,
    add_bullet, style_table
)

def build_chapter_2(doc):
    add_heading_chapter(doc, "CHAPTER 2 – LITERATURE SURVEY")
    
    lead_text = (
        "The literature survey investigates state-of-the-art foundations relevant to the core technical pillars of the "
        "Engineering Intelligence Hub (EIH): repository-level code understanding, Abstract Syntax Tree (AST) static analysis, "
        "multi-agent orchestration for software engineering, Retrieval-Augmented Generation (RAG) for source code, context-window "
        "constraints, automated test generation, and human-in-the-loop change management. This survey analyzes recent academic "
        "benchmarks, commercial developer tools, and theoretical breakthroughs to establish the technical foundation and "
        "identify the explicit research gap that EIH resolves."
    )
    add_p(doc, lead_text, space_after=12, line_spacing=1.2)
    
    # 2.1 Introduction
    add_heading_1(doc, "2.1 Introduction")
    p1 = (
        "Software development is inherently a collaborative, multi-artifact discipline. As codebases scale, software artifacts "
        "proliferate across source files, interfaces, database schemas, test cases, build scripts, and documentation. While early "
        "neural code synthesis focused on isolated function generation (such as HumanEval or MBPP benchmarks), real-world software "
        "engineering demands repository-aware reasoning. A single feature implementation typically modifies multiple interdependent "
        "files, adheres to project-specific conventions, updates documentation, and requires corresponding unit and integration test "
        "suites. The review examines how recent research has tackled these challenges through multi-agent collaboration, semantic "
        "retrieval, and automated execution verification."
    )
    add_p(doc, p1, space_after=10, line_spacing=1.2)
    
    # 2.2 Review of Existing Systems and Methods
    add_heading_1(doc, "2.2 Review of Existing Systems and Methods")
    
    # 2.2.1 Code Broker
    add_heading_2(doc, "2.2.1 Code Broker – A Multi-Agent System for Automated Code Quality Assessment")
    p_cb = (
        "S. Attrah (2026) introduced Code Broker [1], a multi-agent framework specifically engineered to evaluate repository-level "
        "code quality. Code Broker distributes static code analysis, style auditing, and complexity scoring across specialized AI "
        "agents coordinated by a central message broker. Rather than relying on a single monolithic prompt, Code Broker assigns "
        "distinct agents to calculate cyclomatic complexity, flag code smells, and assess architectural cohesion. The primary "
        "relevance to EIH is the validation of the multi-agent paradigm: specialized agents outperform generalist models by focusing "
        "on narrow, well-defined analytical tasks. However, Code Broker is purely an audit and assessment system; it does not "
        "generate new code, conduct repository question answering, synthesize test suites, or execute self-correction loops."
    )
    add_p(doc, p_cb, space_after=8, line_spacing=1.2)
    
    # 2.2.2 RepoAI
    add_heading_2(doc, "2.2.2 RepoAI – Automated Code Refactoring through Multi-Agent LLM Orchestration")
    p_repo = (
        "N. Chondamrongkul et al. (2026) developed RepoAI [2], an autonomous refactoring system that combines multi-agent LLM "
        "orchestration with Retrieval-Augmented Generation to perform repository-wide code refactoring. RepoAI implements agents "
        "for dependency mapping, code transformation, and post-refactor verification. The authors demonstrated that RAG is essential "
        "for identifying cross-file symbol references before applying refactoring transformations. While RepoAI represents a major "
        "step forward in repository-scale reasoning, its scope is restricted solely to refactoring existing syntax. It lacks a general "
        "engineering task execution engine, cannot author end-to-end features or test suites from natural-language requirements, and "
        "does not feature an interactive user console or change-staging diff viewer."
    )
    add_p(doc, p_repo, space_after=8, line_spacing=1.2)
    
    # 2.2.3 Agentic AI for Code Quality
    add_heading_2(doc, "2.2.3 Agentic AI for Code Quality and Self-Correction")
    p_ang = (
        "A. P. T. Angadala (2026) proposed a four-agent machine learning architecture for repository refactoring and automated code "
        "quality improvement [3]. The study explored the integration of fast LLM inference engines (Groq) with reinforcement learning "
        "and public RAG to iteratively refine generated code. The central insight of this work is the efficacy of automated error "
        "feedback: providing compiler diagnostics and linter outputs back to the generative agent significantly increases the rate of "
        "syntactically valid code. EIH builds upon this principle by implementing a formalized 3-attempt self-correction loop, but "
        "expands the architecture to 12 specialized agents encompassing design, diagramming, testing, and documentation."
    )
    add_p(doc, p_ang, space_after=8, line_spacing=1.2)
    
    # 2.2.4 RAG for Code Repositories
    add_heading_2(doc, "2.2.4 Retrieval-Augmented Generation for Code Repositories")
    p_rag = (
        "Lewis et al. (2020) pioneered Retrieval-Augmented Generation (RAG) [4], demonstrating that coupling pre-trained language "
        "models with external non-parametric dense vector indices allows models to retrieve domain-specific factual knowledge at "
        "inference time. In software engineering, RAG has become the cornerstone for repository understanding. However, applying "
        "standard NLP RAG to codebases introduces critical failure modes: standard text splitters divide text based on character counts "
        "or whitespace, arbitrarily severing function declarations, decorators, and control structures. EIH overcomes this limitation "
        "by introducing a code-aware AST chunker that respects semantic entity boundaries and attaches exact line metadata."
    )
    add_p(doc, p_rag, space_after=8, line_spacing=1.2)
    
    # 2.2.5 Lost in the Middle
    add_heading_2(doc, "2.2.5 Lost in the Middle: Context-Window Limitations in Large Codebases")
    p_lost = (
        "N. F. Liu et al. (2024) published a seminal investigation into the retrieval performance of large language models operating "
        "over extensive context windows, identifying the 'Lost in the Middle' phenomenon [5]. Their empirical experiments demonstrated "
        "that LLM recall and reasoning degrade significantly when relevant factual evidence is positioned in the middle of a large "
        "prompt, as opposed to the extreme beginning or end. For software repositories spanning tens of thousands of lines, this proves "
        "that simply stuffing an entire codebase into an ultra-long context window (e.g., 1M tokens) is computationally wasteful, "
        "costly, and prone to severe hallucination. This finding strongly reinforces EIH's design decision to use targeted semantic "
        "retrieval and hybrid symbol boosting rather than naive repository dumping."
    )
    add_p(doc, p_lost, space_after=8, line_spacing=1.2)
    
    # 2.2.6 AST and Code-Aware Program Representation
    add_heading_2(doc, "2.2.6 Abstract Syntax Trees (AST) for Deterministic Code Intelligence")
    p_ast = (
        "Static code analysis literature has long established that Abstract Syntax Trees (ASTs) provide an unambiguous, hierarchical "
        "representation of programming language syntax. Unlike regular expressions or token-level n-grams, AST parsers deterministically "
        "extract class inheritance hierarchies, function argument signatures, async coroutine modifiers, docstring blocks, and call "
        "expressions. In EIH, Python's native AST parser and tree-sitter grammars are utilized to establish deterministic ground truth. "
        "By querying AST nodes, the platform verifies whether a cited method or class physically exists in the repository, providing a "
        "rigorous mathematical verification layer that prevents generative hallucinations."
    )
    add_p(doc, p_ast, space_after=8, line_spacing=1.2)
    
    # 2.2.7 Human-in-the-Loop Software Engineering
    add_heading_2(doc, "2.2.7 Human-in-the-Loop Gating and Change Management")
    p_hitl = (
        "Recent human-computer interaction (HCI) and software reliability research emphasizes the critical necessity of "
        "Human-in-the-Loop (HITL) safeguards when deploying autonomous generative agents. Uncontrolled autonomous write operations "
        "can silently corrupt production repositories, introduce subtle logic regressions, or overwrite uncommitted local work. "
        "Industry best practices in continuous deployment dictate that automated changes must be presented as discrete, inspectable "
        "unified diffs. EIH adopts this philosophy by enforcing zero in-place mutations: all agent outputs are isolated in an ephemeral "
        "staging sandbox where developers can review side-by-side unified diffs and explicitly Approve, Reject, or Request Revision."
    )
    add_p(doc, p_hitl, space_after=8, line_spacing=1.2)
    
    # 2.2.8 Sandbox Isolation
    add_heading_2(doc, "2.2.8 Sandbox Isolation and Execution Security")
    p_sbx = (
        "Executing AI-generated code introduces significant security hazards, including arbitrary code execution, unintended network "
        "calls, and secret exfiltration. Prior literature highlights the necessity of sandboxed test environments that enforce process "
        "isolation, strict execution timeouts (e.g., 30–60 seconds), and automated secret scanning. EIH incorporates a dedicated "
        "Sandbox Runner that executes pytest suites within restricted process boundaries and uses an automated Secret Redactor to scan "
        "and mask GitHub tokens, database connection strings, and API keys before any log is persisted."
    )
    add_p(doc, p_sbx, space_after=12, line_spacing=1.2)
    
    # 2.3 Comparison of Existing Approaches
    add_heading_1(doc, "2.3 Comparison of Existing Approaches")
    add_p(doc, "Table 2.1 presents a comprehensive comparative analysis of existing academic and commercial approaches in AI-assisted "
               "software engineering, contrasting their capabilities with the unified architecture of the Engineering Intelligence Hub.", 
          space_after=8)
          
    t_comp = doc.add_table(rows=8, cols=4)
    t_comp.rows[0].cells[0].text = "Approach / Platform"
    t_comp.rows[0].cells[1].text = "Key Advantages"
    t_comp.rows[0].cells[2].text = "Core Limitations"
    t_comp.rows[0].cells[3].text = "Implementation in EIH"
    for j in range(4):
        t_comp.rows[0].cells[j].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
    comp_data = [
        ("Code Broker [1]\n(Attrah, 2026)", 
         "Multi-agent architecture; static quality metrics; distributed analysis.", 
         "Read-only assessment; no code generation, testing, or Q&A.", 
         "Specialized Review Agent evaluates code quality (0–100) and security."),
        ("RepoAI [2]\n(Chondamrongkul, 2026)", 
         "RAG-guided refactoring; multi-file symbol dependency tracking.", 
         "Restricted to syntax refactoring; no full feature synthesis or UI.", 
         "Extended to full 12-agent feature generation, testing, and docs."),
        ("Agentic AI [3]\n(Angadala, 2026)", 
         "Fast Groq reasoning; error feedback loop for syntax correction.", 
         "Limited to 4 agents; lacks interactive staging diffs and test suites.", 
         "Formalized 3-retry self-correction loop with sandbox pytest execution."),
        ("Standard RAG [4]\n(Lewis et al., 2020)", 
         "Retrieves external non-parametric knowledge; reduces hallucination.", 
         "Arbitrary text splitters bisect functions; lacks line preservation.", 
         "CodeAwareChunker preserves class/function bounds and line ranges."),
        ("Long-Context LLMs [5]\n(Liu et al., 2024)", 
         "Accepts entire repositories (up to 1M tokens) in a single prompt.", 
         "Severe 'Lost in the Middle' degradation; high latency and token cost.", 
         "Replaced by hybrid semantic vector search with symbol boosting."),
        ("Conventional IDE Tools\n(Copilot, Cursor)", 
         "Low latency inline autocomplete; tight editor integration.", 
         "Single-file view; ungrounded generation; no automated test validation.", 
         "Repository-wide understanding, multi-file diff staging, and HITL gate."),
        ("Static AST Analyzers\n(SonarQube, Bandit)", 
         "Deterministic syntax validation; zero hallucination; fast execution.", 
         "Cannot generate code, plan requirements, or understand intent.", 
         "Coupled with LLMs to provide mathematical ground truth for citations.")
    ]
    for i, row in enumerate(comp_data):
        for j, val in enumerate(row):
            t_comp.rows[i+1].cells[j].text = val
    style_table(t_comp, col_widths=[Inches(1.5), Inches(1.6), Inches(1.6), Inches(1.5)])
    
    # 2.4 Research and Technology Gap
    add_heading_1(doc, "2.4 Research and Technology Gap")
    p_gap = (
        "The literature demonstrates that existing software engineering tools remain severely fragmented. Current solutions operate "
        "in silos: static analyzers compute metrics but cannot write code; generative assistants suggest code but cannot verify "
        "correctness across repository dependencies; and refactoring tools manipulate syntax without providing developer interaction or "
        "test generation. Furthermore, four explicit technological gaps remain unaddressed in the literature:\n\n"
        "1. Semantic Boundary Mismatch: Standard vector retrieval splits code at arbitrary byte boundaries, bisecting critical control "
        "loops and class hierarchies.\n"
        "2. Absence of Line-Level Citation Verification: Language models answer codebase queries with narrative descriptions but provide "
        "no verifiable proofs linking assertions to exact file lines.\n"
        "3. Lack of Autonomous Diagnostic Recovery: When generated code fails, developers must manually intervene, debug tracebacks, and "
        "re-prompt the AI.\n"
        "4. Risk of In-Place Repository Corruption: Most autonomous agents write directly to working directories without isolating changes "
        "in a diff review sandbox.\n\n"
        "The Engineering Intelligence Hub resolves this entire gap by unifying AST code intelligence, line-verified RAG, 12-agent collaborative "
        "state machines, automated self-correction, and human-in-the-loop staging into a single integrated platform."
    )
    add_p(doc, p_gap, space_after=10, line_spacing=1.2)
    
    # 2.5 Limitations of Existing Methods
    add_heading_1(doc, "2.5 Limitations of Existing Methods")
    add_bullet(doc, "Treating code as raw text ignores the AST syntax hierarchy, leading to syntax errors and broken imports.", bold_prefix="• Lack of Structural Code Awareness:")
    add_bullet(doc, "Passing massive repository dumps into long contexts results in attention decay and hallucinated details.", bold_prefix="• Vulnerability to 'Lost in the Middle':")
    add_bullet(doc, "Monolithic prompt architectures compound errors across planning, coding, and testing phases.", bold_prefix="• Monolithic Single-Pass Fragility:")
    add_bullet(doc, "Generative models produce plausible-looking but broken code without automated compiler or test verification.", bold_prefix="• Absence of Execution Feedback:")
    add_bullet(doc, "Direct codebase mutations risk silent corruption, loss of uncommitted work, and broken production builds.", bold_prefix="• Inadequate Change Staging:")
    
    # 2.6 Summary
    add_heading_1(doc, "2.6 Summary")
    add_p(doc, "The literature confirms that an effective AI software engineering platform must combine structural AST parsing, "
               "selective dense vector retrieval, multi-agent specialization, automated verification, and human oversight. EIH realizes "
               "this vision through a coordinated three-path architecture. Chapter 3 translates these architectural insights into formal "
               "engineering requirements, mathematical algorithms, and a proposed methodology.", space_after=0)
