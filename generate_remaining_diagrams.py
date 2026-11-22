import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

matplotlib.rcParams["font.family"] = "serif"
matplotlib.rcParams["font.serif"] = ["Times New Roman", "DejaVu Serif"]

def draw_rect(ax, x, y, w, h, text, bg="#E3F2FD", border="#1565C0", fs=8.5, bold=False):
    p = patches.FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.15", facecolor=bg, edgecolor=border, linewidth=1.2)
    ax.add_patch(p)
    weight = "bold" if bold else "normal"
    ax.text(x + w/2, y + h/2, text, ha="center", va="center", fontsize=fs, weight=weight, color="#212121", multialignment="center")

# 1. Figure 4.5: Sequence Diagram
fig, ax = plt.subplots(figsize=(8.5, 7.5), dpi=300)
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis("off")

cols = ["Developer\n(Browser)", "FastAPI\nRouter", "ChromaDB\nRetriever", "AST Parser /\nSymbol DB", "LLM Provider\n(Groq/OpenAI)", "Citation\nValidator"]
x_cols = np.linspace(1.0, 9.0, len(cols))

for name, x in zip(cols, x_cols):
    draw_rect(ax, x - 0.7, 8.8, 1.4, 0.9, name, "#E8EAF6", "#3949AB", 7.5, True)
    ax.plot([x, x], [1.0, 8.8], color="#B0BEC5", linestyle="--", lw=1.2)

msgs = [
    (0, 1, 7.8, "1: POST /api/v1/chat (query)"),
    (1, 2, 7.0, "2: query_dense_embeddings(top_k=5)"),
    (2, 1, 6.2, "3: return top code chunks + lines"),
    (1, 3, 5.4, "4: verify_symbols_in_chunk(file, lines)"),
    (3, 1, 4.6, "5: symbol metadata (class/fn defs)"),
    (1, 4, 3.8, "6: generate_response(prompt + context)"),
    (4, 1, 3.0, "7: streaming tokens + cited spans"),
    (1, 5, 2.2, "8: validate_citations_against_ast()"),
    (5, 1, 1.6, "9: verified citation objects"),
    (1, 0, 1.1, "10: JSON response {answer, citations}")
]

for src, dst, y, text in msgs:
    x1, x2 = x_cols[src], x_cols[dst]
    ax.annotate("", xy=(x2, y), xytext=(x1, y), arrowprops=dict(arrowstyle="->", lw=1.2, color="#1565C0" if x2>x1 else "#2E7D32"))
    mid_x = (x1 + x2) / 2
    ax.text(mid_x, y + 0.15, text, ha="center", va="bottom", fontsize=7.5, color="#0D47A1" if x2>x1 else "#1B5E20")

ax.set_title("Figure 4.5: UML Sequence Diagram of a Grounded RAG Query and Citation Linking", fontsize=11, fontweight="bold", pad=12)
plt.tight_layout()
plt.savefig("extracted_figures/fig4_5_sequence.png", bbox_inches="tight")
plt.close()

# 2. Figure 4.9: Agent Interaction & State Machine Flow
fig, ax = plt.subplots(figsize=(8.5, 7.5), dpi=300)
ax.set_xlim(0, 10)
ax.set_ylim(0, 10)
ax.axis("off")

draw_rect(ax, 3.5, 4.2, 3.0, 1.6, "ORCHESTRATOR AGENT\n(State Machine Manager,\nComplexity Router &\nRetry Controller)", "#FFF3E0", "#E65100", 8.5, True)
draw_rect(ax, 0.5, 7.5, 2.6, 1.5, "PLANNING AGENTS\n• Requirement Agent\n• Repository Agent\n• Code Understanding", "#E8F5E9", "#2E7D32", 8)
draw_rect(ax, 6.9, 7.5, 2.6, 1.5, "BUILD AGENTS\n• System Design Agent\n• Code Generation Agent\n• Test Generation Agent", "#E1F5FE", "#0288D1", 8)
draw_rect(ax, 6.9, 1.0, 2.6, 1.5, "VERIFY AGENTS\n• Validation Agent\n• Review & Security Agent\n• Sandbox Runner", "#F3E5F5", "#7B1FA2", 8)
draw_rect(ax, 0.5, 1.0, 2.6, 1.5, "DOCUMENTATION AGENTS\n• Documentation Agent\n• Diagram Studio Agent\n• Change Impact Agent", "#FCE4EC", "#C2185B", 8)

ax.annotate("", xy=(1.8, 7.5), xytext=(4.0, 5.8), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#455A64"))
ax.annotate("", xy=(7.2, 7.5), xytext=(6.0, 5.8), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#455A64"))
ax.annotate("", xy=(7.2, 2.5), xytext=(6.0, 4.2), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#455A64"))
ax.annotate("", xy=(1.8, 2.5), xytext=(4.0, 4.2), arrowprops=dict(arrowstyle="<->", lw=1.5, color="#455A64"))
ax.annotate("Self-Correction Diagnostic Loop (Error trace returned to CodeGen)", xy=(7.5, 7.5), xytext=(7.5, 2.5),
            arrowprops=dict(arrowstyle="->", lw=1.5, color="#D32F2F", linestyle="--", connectionstyle="arc3,rad=-0.4"))

ax.set_title("Figure 4.9: Interaction Architecture between the 12 Specialized Autonomous Agents", fontsize=11, fontweight="bold", pad=12)
plt.tight_layout()
plt.savefig("extracted_figures/fig4_9_agents.png", bbox_inches="tight")
plt.close()

# 3. Figure 4.10: Data Workflow
fig, ax = plt.subplots(figsize=(8.5, 5.0), dpi=300)
ax.set_xlim(0, 10)
ax.set_ylim(0, 7)
ax.axis("off")

steps_data = [
    "Repository Ingestion\n(Clone Git repo or\nopen workspace)",
    "File Filtering\n(Exclude .git, node_modules,\nsecrets, binary assets)",
    "AST Symbol Extraction\n(Parse classes, functions,\nmethods, line numbers)",
    "Code-Aware Chunking\n(Preserve semantic boundaries,\nassign file:line tags)",
    "Vector Indexing\n(ChromaDB collection,\n384-d dense embeddings)"
]
x_s = np.linspace(0.5, 7.8, 5)
for s, x in zip(steps_data, x_s):
    draw_rect(ax, x, 2.5, 1.7, 2.0, s, "#E0F2F1", "#00796B", 7.5, True)

for i in range(4):
    ax.annotate("", xy=(x_s[i+1], 3.5), xytext=(x_s[i]+1.7, 3.5), arrowprops=dict(arrowstyle="->", lw=1.5, color="#004D40"))

ax.set_title("Figure 4.10: Code Ingestion, AST Parsing, and Knowledge Preparation Pipeline", fontsize=11, fontweight="bold", pad=12)
plt.tight_layout()
plt.savefig("extracted_figures/fig4_10_pipeline.png", bbox_inches="tight")
plt.close()

# 4. Figure 4.12: RAG Pipeline
fig, ax = plt.subplots(figsize=(8.5, 5.2), dpi=300)
ax.set_xlim(0, 10)
ax.set_ylim(0, 7)
ax.axis("off")

draw_rect(ax, 0.5, 4.0, 2.0, 1.5, "User Query\n(e.g., How is JWT\nauth handled?)", "#E8EAF6", "#3949AB", 8, True)
draw_rect(ax, 3.2, 4.0, 2.2, 1.5, "Hybrid Retrieval\n(Dense similarity +\nAST symbol boost)", "#E1F5FE", "#0288D1", 8, True)
draw_rect(ax, 6.0, 4.0, 2.0, 1.5, "LLM Generation\n(Prompt + Top Chunks\nwith Line Metadata)", "#FFF3E0", "#E65100", 8, True)
draw_rect(ax, 4.5, 1.0, 2.5, 1.5, "Citation Verification\n(Match cited lines to AST\nsymbol ground truth)", "#E8F5E9", "#2E7D32", 8, True)
draw_rect(ax, 7.5, 1.0, 2.0, 1.5, "Grounded Answer\nwith Clickable\nCitations", "#F3E5F5", "#7B1FA2", 8, True)

ax.annotate("", xy=(3.2, 4.75), xytext=(2.5, 4.75), arrowprops=dict(arrowstyle="->", lw=1.5, color="#37474F"))
ax.annotate("", xy=(6.0, 4.75), xytext=(5.4, 4.75), arrowprops=dict(arrowstyle="->", lw=1.5, color="#37474F"))
ax.annotate("", xy=(5.75, 2.5), xytext=(7.0, 4.0), arrowprops=dict(arrowstyle="->", lw=1.5, color="#37474F"))
ax.annotate("", xy=(7.5, 1.75), xytext=(7.0, 1.75), arrowprops=dict(arrowstyle="->", lw=1.5, color="#37474F"))

ax.set_title("Figure 4.12: Grounded RAG Architecture with AST-Driven Line Citation Verification", fontsize=11, fontweight="bold", pad=12)
plt.tight_layout()
plt.savefig("extracted_figures/fig4_12_rag_pipeline.png", bbox_inches="tight")
plt.close()

# 5. Figure 5.6: Quality Score vs Execution Time
tasks = ["Task 1 (Auth)", "Task 2 (API Route)", "Task 3 (DB Model)", "Task 4 (Refactor)", "Task 5 (Test Suite)", "Task 6 (Docs Sync)"]
single_times = [3.2, 2.8, 3.5, 4.1, 3.0, 2.5]
single_scores = [82, 86, 84, 80, 88, 85]
multi_times = [8.5, 7.8, 9.1, 10.2, 8.1, 7.2]
multi_scores = [96, 98, 95, 94, 98, 97]

fig, ax = plt.subplots(figsize=(8, 4.8), dpi=300)
ax.scatter(single_times, single_scores, color="#78909C", s=100, label="Single-Agent Baseline", edgecolor="#37474F", zorder=5)
ax.scatter(multi_times, multi_scores, color="#1976D2", s=120, label="EIH Multi-Agent Workflow", marker="^", edgecolor="#0D47A1", zorder=5)

for i, txt in enumerate(tasks):
    ax.annotate(txt, (single_times[i], single_scores[i]-1.5), fontsize=7.5, ha="center", color="#455A64")
    ax.annotate(txt, (multi_times[i], multi_scores[i]+1.2), fontsize=7.5, ha="center", color="#0D47A1")

ax.set_xlabel("Execution Latency (Seconds)", fontsize=10, fontweight="bold")
ax.set_ylabel("Comprehensive Quality Score (/100)", fontsize=10, fontweight="bold")
ax.set_title("Figure 5.6: Latency vs Code Quality Trade-Off across Six Engineering Benchmark Tasks", fontsize=11, fontweight="bold", pad=12)
ax.legend(frameon=True, facecolor="white", edgecolor="#B0BEC5")
ax.set_ylim(75, 102)
ax.set_xlim(2.0, 11.5)
plt.tight_layout()
plt.savefig("extracted_figures/fig5_6_quality_vs_time.png", bbox_inches="tight")
plt.close()
print("All remaining diagrams generated successfully.")
