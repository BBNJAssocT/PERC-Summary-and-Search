#!/usr/bin/env python3
"""
Case Summarizer — SUPABASE VERSION
====================================
Same extraction, Claude summarization, law-enforcement detection, citation
matching, and Word-doc formatting as the local version -- but PDFs and
results now live in Supabase (Storage + Postgres) instead of a local
folder and SQLite file, so multiple invited people can search a shared,
always-current archive from a website (built separately, Phase 6).

This script still runs from YOUR machine, administratively, whenever you
choose to refresh -- it is not automatically triggered by anyone else.

-----------------------------------------------------------------
SETUP (one time)
-----------------------------------------------------------------
1. pip install pypdf python-docx anthropic supabase
2. Set THREE environment variables (never hardcoded in this file):
       export ANTHROPIC_API_KEY="sk-ant-..."
       export SUPABASE_URL="https://YOUR-PROJECT.supabase.co"
       export SUPABASE_SERVICE_KEY="..."   # the SECRET service_role key,
                                            # from Supabase dashboard ->
                                            # Settings -> API. This bypasses
                                            # Row Level Security, so treat
                                            # it exactly like your Anthropic
                                            # key: environment variable only,
                                            # never pasted anywhere, never
                                            # committed to a shared script.
3. Set CASE_SUMMARIZER_FOLDER to your local PDF folder (same as before):
       export CASE_SUMMARIZER_FOLDER="/path/to/your/pdf/folder"

-----------------------------------------------------------------
USAGE
-----------------------------------------------------------------
    python case_summarizer_supabase.py                  # refresh + build outputs
    python case_summarizer_supabase.py --force           # re-summarize everything
    python case_summarizer_supabase.py --search "sick leave"   # search Supabase directly
"""

import argparse
import hashlib
import json
import os
import re as _re
import sys
import uuid
from datetime import datetime, timezone
from pathlib import Path

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
BASE_FOLDER = os.environ.get("CASE_SUMMARIZER_FOLDER", ".")
OUTPUT_DIR = BASE_FOLDER
DOCX_OUTPUT = os.path.join(OUTPUT_DIR, "Case_Summaries.docx")
PDF_OUTPUT = os.path.join(OUTPUT_DIR, "Case_Summaries.pdf")
MAX_CHARS_PER_PDF = 20000        # sent to Claude for summarization (cost control)
MAX_CHARS_FOR_SEARCH = 500000    # stored/indexed for full-text search (not sent to Claude)
CLAUDE_MODEL = "claude-sonnet-4-6"

PDF_BUCKET = "case-pdfs"
SUMMARY_SNAPSHOT_KEY = "_Case_Summaries.pdf"   # leading underscore keeps it sorted first,
                                                 # away from real case filenames in the bucket

# Filenames this script itself generates -- excluded when scanning for case
# PDFs, so the script never tries to re-summarize its own output.
GENERATED_FILENAMES = {"Case_Summaries.pdf"}


def get_supabase_client():
    from supabase import create_client

    url = os.environ.get("SUPABASE_URL")
    key = os.environ.get("SUPABASE_SERVICE_KEY")
    if not url or not key:
        print("ERROR: Set both SUPABASE_URL and SUPABASE_SERVICE_KEY environment variables first.")
        sys.exit(1)
    return create_client(url, key)


# --------------------------------------------------------------------------
# PDF text extraction
# --------------------------------------------------------------------------
def extract_pdf_text(pdf_path: str, max_chars: int) -> str:
    from pypdf import PdfReader
    try:
        reader = PdfReader(pdf_path)
        text_parts, total = [], 0
        for page in reader.pages:
            t = page.extract_text() or ""
            text_parts.append(t)
            total += len(t)
            if total >= max_chars:
                break
        return "\n".join(text_parts)[:max_chars]
    except Exception as e:
        return f"[Could not extract text: {e}]"


def extract_decision_date(full_text: str):
    """Pulls the decision's actual issue date out of the document text, e.g.
    'ISSUED: August 27, 2026'. Returns an ISO date string (YYYY-MM-DD), or
    None if no recognizable date is found."""
    patterns = [
        r"ISSUED:?\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"Decided:?\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})",
        r"Dated:?\s*([A-Za-z]+\s+\d{1,2},\s*\d{4})",
    ]
    for pattern in patterns:
        match = _re.search(pattern, full_text, _re.IGNORECASE)
        if match:
            try:
                parsed = datetime.strptime(match.group(1).strip(), "%B %d, %Y")
                return parsed.strftime("%Y-%m-%d")
            except ValueError:
                continue
    return None


def file_fingerprint(path: str) -> str:
    stat = os.stat(path)
    return hashlib.sha256(f"{stat.st_size}-{stat.st_mtime}".encode()).hexdigest()


# --------------------------------------------------------------------------
# Claude summarization
# --------------------------------------------------------------------------
SYSTEM_PROMPT = """You are a legal document analyst reviewing labor/employment
relations board decisions (e.g. PERC-style rulings). You will be given the
raw text of one decision.

Respond ONLY with a single JSON object, no other text, no markdown fences,
with exactly these keys:

{
  "parties": "Names of the parties involved, e.g. 'PBA Local 11A v. City of Trenton'",
  "summary": "A concise 3-5 sentence plain-English summary of the case.",
  "factors": ["A short bullet phrase for each key factor or issue considered", "..."],
  "decision": "A concise summary of the decision, ruling, or determination reached.",
  "category": "A short substantive employment/labor-law SUBJECT MATTER label
    describing the underlying workplace issue this case is actually ABOUT,
    based on the core issue in dispute -- e.g. 'Management Prerogative',
    'Sick Leave', 'Payroll/Compensation', 'Discipline/Just Cause',
    'Duty to Negotiate', 'Grievance Arbitrability', 'Past Practice',
    'Layoffs/Reduction in Force', 'Health Benefits', 'Overtime',
    'Union Security/Dues', 'Retaliation'. Do NOT use procedural postures
    (e.g. 'Interim Relief', 'Timeliness') as the category -- look past the
    procedural posture to the actual employment topic underneath."
}

If a field cannot be determined, use "Not specified" (or an empty list for
factors). Do not invent facts not present in the text.
"""


def summarize_with_claude(client, text: str) -> dict:
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=1500,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": text}],
    )
    raw = "".join(b.text for b in message.content if b.type == "text").strip().strip("`")
    if raw.lower().startswith("json"):
        raw = raw[4:].strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return {
            "parties": "Not specified", "summary": "Could not parse a summary.",
            "factors": [], "decision": "Not specified", "category": "Uncategorized",
        }


# --------------------------------------------------------------------------
# Citation extraction (pure functions, no DB dependency)
# --------------------------------------------------------------------------
_CITATION_PATTERN = _re.compile(
    r"(P\.E\.R\.C\.|D\.U\.P\.)\s*No\.?\s*(\d{2,4}-\d+)", _re.IGNORECASE
)


def _normalize_citation(label: str, number: str) -> str:
    clean_label = label.upper().replace(".", "").replace(" ", "")
    if clean_label == "PERC":
        return f"P.E.R.C. No. {number}"
    if clean_label == "DUP":
        return f"D.U.P. No. {number}"
    return f"{label} No. {number}"


def _all_normalized_citations(full_text: str) -> list:
    if not full_text:
        return []
    matches = _CITATION_PATTERN.findall(full_text)
    return [_normalize_citation(label, number) for label, number in matches]


def extract_self_citation(full_text: str):
    normalized = _all_normalized_citations(full_text)
    if not normalized:
        return None
    from collections import Counter
    return Counter(normalized).most_common(1)[0][0]


def extract_related_decisions(full_text: str, max_items: int = 8) -> list:
    normalized = _all_normalized_citations(full_text)
    if not normalized:
        return []
    self_citation = extract_self_citation(full_text)
    seen = set()
    related = []
    for citation in normalized:
        if citation == self_citation or citation in seen:
            continue
        seen.add(citation)
        related.append(citation)
        if len(related) >= max_items:
            break
    return related


def resolve_related_decisions(supabase, related_citations: list) -> list:
    """For each related citation, checks whether a document with that exact
    self-citation already exists in Supabase. Returns a list of dicts:
    {"citation": str, "storage_path": str|None}. Re-running this against the
    live database means a citation automatically starts resolving to a real
    link as soon as that document gets added in a future run."""
    results = []
    for citation in related_citations:
        resp = (
            supabase.table("documents")
            .select("storage_path")
            .eq("citation", citation)
            .limit(1)
            .execute()
        )
        if resp.data:
            results.append({"citation": citation, "storage_path": resp.data[0]["storage_path"]})
        else:
            results.append({"citation": citation, "storage_path": None})
    return results


# --------------------------------------------------------------------------
# Law enforcement detection (pure function, no DB dependency)
# --------------------------------------------------------------------------
LAW_ENFORCEMENT_KEYWORDS = [
    "law enforcement", "police", "detective", "detectives", "corrections",
    "investigator", "investigators", "sheriff", "trooper", "troopers",
    "patrolman", "patrolmen", "constable", "officer", "officers",
    "prosecutor", "prosecutors", "pba", "fop", "soa",
]
_LAW_ENFORCEMENT_PATTERN = _re.compile(
    r"\b(" + "|".join(_re.escape(k) for k in LAW_ENFORCEMENT_KEYWORDS) + r")\b",
    _re.IGNORECASE,
)


def is_law_enforcement_case(parties, category, summary, decision, factors) -> bool:
    """factors here is a plain Python list (Supabase returns jsonb natively
    as a list, unlike the local SQLite version which stored it as a JSON
    string needing json.loads)."""
    factors = factors or []
    haystack = " ".join([parties or "", category or "", summary or "", decision or "", " ".join(factors)])
    return bool(_LAW_ENFORCEMENT_PATTERN.search(haystack))


# --------------------------------------------------------------------------
# Scanning
# --------------------------------------------------------------------------
def find_pdfs(base_folder: str):
    base = Path(base_folder)
    return [
        p for p in base.rglob("*.pdf")
        if p.name not in GENERATED_FILENAMES
    ]


def scan_and_summarize(base_folder: str, supabase, force: bool = False):
    import anthropic

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        print("ERROR: Set the ANTHROPIC_API_KEY environment variable first.")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key)
    base = Path(base_folder)
    pdfs = find_pdfs(base_folder)

    if not pdfs:
        print(f"No PDFs found under {base_folder}")
        return

    new_count, skipped_count = 0, 0
    for local_path in pdfs:
        # storage_path is relative to the folder, NOT the local absolute
        # path -- this is what's stored in the bucket and the database, so
        # nothing machine-specific ever leaves your computer.
        storage_path = str(local_path.relative_to(base)).replace("\\", "/")
        fp = file_fingerprint(str(local_path))

        existing = (
            supabase.table("documents")
            .select("file_hash")
            .eq("storage_path", storage_path)
            .limit(1)
            .execute()
        )
        if existing.data and existing.data[0]["file_hash"] == fp and not force:
            skipped_count += 1
            continue

        print(f"Summarizing: {storage_path}")
        text_for_ai = extract_pdf_text(str(local_path), max_chars=MAX_CHARS_PER_PDF)
        full_text = extract_pdf_text(str(local_path), max_chars=MAX_CHARS_FOR_SEARCH)
        data = summarize_with_claude(client, text_for_ai)
        category = data.get("category", "Uncategorized")
        decision_date = extract_decision_date(full_text)
        self_citation = extract_self_citation(full_text)

        # Upload the actual PDF bytes to Supabase Storage
        with open(local_path, "rb") as f:
            file_bytes = f.read()
        try:
            supabase.storage.from_(PDF_BUCKET).upload(
                storage_path, file_bytes,
                {"content-type": "application/pdf", "upsert": "true"},
            )
        except Exception as e:
            print(f"  WARNING: PDF upload failed for {storage_path}: {e}")

        # Upsert the row (storage_path is the unique key -- matches the
        # unique index already created on the documents table)
        supabase.table("documents").upsert(
            {
                "filename": local_path.name,
                "storage_path": storage_path,
                "file_hash": fp,
                "parties": data.get("parties", "Not specified"),
                "summary": data.get("summary", "Not specified"),
                "factors": data.get("factors", []),
                "decision": data.get("decision", "Not specified"),
                "category": category,
                "full_text": full_text,
                "decision_date": decision_date,
                "citation": self_citation,
                "updated_at": datetime.now(timezone.utc).isoformat(),
            },
            on_conflict="storage_path",
        ).execute()
        new_count += 1

    # Remove rows for files that no longer exist locally (mirrors local
    # version's cleanup behavior)
    existing_paths = {str(p.relative_to(base)).replace("\\", "/") for p in pdfs}
    all_rows = supabase.table("documents").select("id, storage_path").execute()
    for row in all_rows.data:
        if row["storage_path"] not in existing_paths:
            supabase.table("documents").delete().eq("id", row["id"]).execute()

    print(f"\nDone. {new_count} document(s) processed, {skipped_count} unchanged (skipped).")


# --------------------------------------------------------------------------
# Word document generation
# --------------------------------------------------------------------------
def add_hyperlink(paragraph, url, text, size_pt=None, italic=False):
    from docx.oxml.shared import OxmlElement, qn

    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    color = OxmlElement("w:color"); color.set(qn("w:val"), "0563C1")
    rPr.append(color)
    u = OxmlElement("w:u"); u.set(qn("w:val"), "single")
    rPr.append(u)
    if italic:
        rPr.append(OxmlElement("w:i"))
    if size_pt is not None:
        sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(int(size_pt * 2)))
        rPr.append(sz)
    new_run.append(rPr)
    t = OxmlElement("w:t"); t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink


def add_case_entry(doc, row, supabase):
    from docx.shared import Pt as _Pt2, RGBColor as _RGBColor

    is_le = is_law_enforcement_case(row["parties"], row["category"], row["summary"], row["decision"], row["factors"])
    tag_p = doc.add_paragraph()
    tag_run = tag_p.add_run("LAW ENFORCEMENT DECISION" if is_le else "OTHER — NON-LAW ENFORCEMENT")
    tag_run.bold = True
    tag_run.italic = True
    tag_run.font.size = _Pt2(10)
    tag_run.font.color.rgb = _RGBColor(0x1A, 0x4D, 0x8F) if is_le else _RGBColor(0x88, 0x88, 0x88)

    doc.add_heading(row["parties"] or "Unknown Parties", level=2)

    if row.get("decision_date"):
        date_p = doc.add_paragraph()
        date_run = date_p.add_run(f"Issued: {row['decision_date']}")
        date_run.italic = True
        date_run.font.size = _Pt2(9)

    p = doc.add_paragraph()
    p.add_run("Summary: ").bold = True
    p.add_run(row["summary"] or "Not specified")

    factors = row["factors"] or []
    if factors:
        fp = doc.add_paragraph()
        fp.add_run("Factors Involved:").bold = True
        for factor in factors:
            doc.add_paragraph(factor, style="List Bullet")

    dp = doc.add_paragraph()
    dp.add_run("Decision / Determination: ").bold = True
    dp.add_run(row["decision"] or "Not specified")

    link_p = doc.add_paragraph()
    link_p.add_run("Source document: ").bold = True
    try:
        signed = supabase.storage.from_(PDF_BUCKET).create_signed_url(row["storage_path"], 31536000)  # 1 year
        pdf_url = signed.get("signedURL") or signed.get("signedUrl")
        add_hyperlink(link_p, pdf_url, row["filename"])
    except Exception:
        link_p.add_run(row["filename"])  # fall back to plain text if signing fails

    related = extract_related_decisions(row["full_text"] or "")
    if related:
        resolved = resolve_related_decisions(supabase, related)
        rel_p = doc.add_paragraph()
        rel_label = rel_p.add_run("Related / Referenced Decisions: ")
        rel_label.bold = True
        rel_label.font.size = _Pt2(8)
        for i, item in enumerate(resolved):
            if i > 0:
                sep = rel_p.add_run(", ")
                sep.font.size = _Pt2(8)
                sep.font.color.rgb = _RGBColor(0x66, 0x66, 0x66)
            if item["storage_path"]:
                try:
                    signed = supabase.storage.from_(PDF_BUCKET).create_signed_url(item["storage_path"], 31536000)
                    citation_url = signed.get("signedURL") or signed.get("signedUrl")
                    add_hyperlink(rel_p, citation_url, item["citation"], size_pt=8, italic=True)
                    continue
                except Exception:
                    pass
            plain_run = rel_p.add_run(item["citation"])
            plain_run.italic = True
            plain_run.font.size = _Pt2(8)
            plain_run.font.color.rgb = _RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()


def build_docx(supabase, output_path: str, recent_count: int = 3):
    from docx import Document
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    doc.add_heading("Case Summaries", level=0)

    footer = doc.sections[0].footer
    footer_p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    footer_run = footer_p.add_run(f"Last updated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}")
    footer_run.font.size = Pt(8)
    footer_run.italic = True

    disclosure = doc.add_paragraph()
    disclosure.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = disclosure.add_run(
        "This document was generated with AI assistance (Anthropic Claude) for "
        "reference and internal case-tracking purposes only. Summaries, factors, "
        "and categorizations are AI-generated and may contain errors or omissions. "
        "Always review the linked source PDF and verify accuracy before relying on "
        "any summary, citing it, or taking action based on it."
    )
    run.italic = True
    run.font.size = Pt(9)
    doc.add_paragraph()

    all_rows = supabase.table("documents").select("*").execute().data

    all_rows_by_recency = sorted(
        all_rows, key=lambda r: r.get("decision_date") or r.get("created_at") or "", reverse=True
    )
    recent_rows = all_rows_by_recency[:recent_count]
    recent_ids = {r["id"] for r in recent_rows}
    remaining_rows = [r for r in all_rows_by_recency if r["id"] not in recent_ids]

    if recent_rows:
        doc.add_heading("Recent Decisions", level=1)
        for row in recent_rows:
            add_case_entry(doc, row, supabase)

    categories = {}
    for row in remaining_rows:
        categories.setdefault(row["category"] or "Uncategorized", []).append(row)

    for category in sorted(categories.keys()):
        doc.add_heading(category, level=1)
        sorted_rows = sorted(
            categories[category],
            key=lambda e: (
                not is_law_enforcement_case(e["parties"], e["category"], e["summary"], e["decision"], e["factors"]),
                e["parties"] or "",
            ),
        )
        for row in sorted_rows:
            add_case_entry(doc, row, supabase)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    print(f"Word document written to: {output_path}")


def convert_docx_to_pdf(docx_path: str, pdf_path: str) -> bool:
    output_dir = os.path.dirname(os.path.abspath(pdf_path))
    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path):
            print(f"PDF written to: {pdf_path}")
            return True
    except Exception:
        pass

    import shutil, subprocess
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        try:
            subprocess.run(
                [soffice, "--headless", "--convert-to", "pdf", "--outdir", output_dir, docx_path],
                check=True, capture_output=True,
            )
            generated = os.path.join(output_dir, Path(docx_path).stem + ".pdf")
            if generated != pdf_path and os.path.exists(generated):
                os.replace(generated, pdf_path)
            if os.path.exists(pdf_path):
                print(f"PDF written to: {pdf_path}")
                return True
        except Exception as e:
            print(f"LibreOffice conversion failed: {e}")

    print(
        "Could not create a PDF automatically. Install LibreOffice (free, "
        "https://www.libreoffice.org) or Microsoft Word, then re-run.\n"
        f"In the meantime, you can open {docx_path} and use File > Export/Save As PDF."
    )
    return False


def upload_summary_pdf(supabase, pdf_path: str):
    """Uploads the just-generated Case_Summaries.pdf to Supabase Storage as
    an administrative snapshot -- reflects the archive as of THIS run, not
    live-regenerated on demand by the website."""
    if not os.path.exists(pdf_path):
        print("No summary PDF found to upload (conversion may have failed above).")
        return
    with open(pdf_path, "rb") as f:
        file_bytes = f.read()
    try:
        supabase.storage.from_(PDF_BUCKET).upload(
            SUMMARY_SNAPSHOT_KEY, file_bytes,
            {"content-type": "application/pdf", "upsert": "true"},
        )
        print(f"Summary snapshot uploaded to Supabase (as of this run).")
    except Exception as e:
        print(f"WARNING: Summary snapshot upload failed: {e}")


# --------------------------------------------------------------------------
# Search directly against Supabase (Postgres full-text search)
# --------------------------------------------------------------------------
def terminal_search(term: str, supabase):
    resp = (
        supabase.table("documents")
        .select("parties, category, storage_path, summary")
        .text_search("search_vector", term, options={"config": "english", "type": "websearch"})
        .execute()
    )
    rows = resp.data
    print(f"\n{len(rows)} match(es) for '{term}':\n")
    for r in rows:
        print(f"- {r['parties']}  [{r['category']}]")
        print(f"    {r['storage_path']}")
    print()


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Scan local PDFs, summarize with Claude, sync to Supabase.")
    parser.add_argument("--folder", default=BASE_FOLDER, help="Base folder to scan")
    parser.add_argument("--force", action="store_true", help="Re-summarize every file, ignoring cache")
    parser.add_argument("--search", help="Search Supabase directly and exit")
    args = parser.parse_args()

    if not os.path.isdir(args.folder):
        print(f"ERROR: Folder not found: {args.folder}\nSet CASE_SUMMARIZER_FOLDER or pass --folder \"/path/to/folder\".")
        sys.exit(1)

    supabase = get_supabase_client()

    if args.search:
        terminal_search(args.search, supabase)
        return

    scan_and_summarize(args.folder, supabase, force=args.force)
    build_docx(supabase, DOCX_OUTPUT)
    convert_docx_to_pdf(DOCX_OUTPUT, PDF_OUTPUT)
    upload_summary_pdf(supabase, PDF_OUTPUT)


if __name__ == "__main__":
    main()
