#!/usr/bin/env python3
"""
Online version of the PDF refresh -- runs in GitHub Actions, not on a laptop.

What it does, in order:
  1. Looks in the private Supabase Storage bucket "incoming" for new PDFs.
  2. For each one: extracts the text, asks Claude for the summary and category,
     copies the PDF into the permanent "case-pdfs" bucket (in a year folder),
     saves the row in the documents table, and removes it from "incoming".
  3. If anything new was added, rebuilds the Word summary (with a clickable
     link to every PDF) and the PDF summary, and uploads both.

It reuses the same extraction, summary, law-enforcement and Word-building code
as case_summarizer_supabase.py (the copy sitting next to this file).

Safety notes
  * It NEVER deletes existing archive rows. The laptop version deleted rows for
    files missing from the scanned folder; here the folder only holds new
    files, so that cleanup is deliberately left out.
  * A PDF that is already in the archive is moved to incoming/_skipped/ instead of
    being summarized (and billed) again. "Already in the archive" means the same
    storage path, the same file contents, or the same citation with identical text
    (a re-saved copy under another name). Use --force to redo it.
  * A PDF that fails (unreadable, no text, summary not parseable) stays in
    "incoming" and the run ends with an error so GitHub emails the owner.

Environment variables (stored as GitHub Actions secrets, never in code):
    ANTHROPIC_API_KEY, SUPABASE_URL, SUPABASE_SERVICE_KEY

Usage
    python cloud/cloud_process.py              # process whatever is waiting
    python cloud/cloud_process.py --count      # only report how many are waiting
    python cloud/cloud_process.py --rebuild    # also rebuild the summary files
    python cloud/cloud_process.py --force      # re-summarize files already in the archive
"""

import argparse
import hashlib
import os
import re
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

import case_summarizer_supabase as cs

INCOMING_BUCKET = "incoming"
SKIPPED_PREFIX = "_skipped/"
DOCX_KEY = "_Case_Summaries.docx"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
UNPARSEABLE_SUMMARY = "Could not parse a summary."

# Human-readable notes collected during a run (what each duplicate matched, warnings).
NOTES: list = []


def normalize_text(text: str) -> str:
    """Whitespace- and case-insensitive form of a decision's text, for comparing copies."""
    return re.sub(r"\s+", " ", text or "").strip().lower()


def find_duplicate(supabase, file_hash: str, citation, full_text: str):
    """Returns the storage_path of an archived decision that is the same document, else None.

    Same document means: (1) identical file bytes, or (2) the same citation AND identical
    text once spacing and capitalization are ignored (catches a re-saved copy of the same
    decision under a different file name). Same citation with different text is NOT a
    duplicate -- it is noted as a warning, because the citation read from a PDF can be wrong.
    """
    hit = supabase.table("documents").select("storage_path").eq("file_hash", file_hash).limit(1).execute()
    if hit.data:
        return hit.data[0]["storage_path"]
    if citation:
        same = (
            supabase.table("documents")
            .select("storage_path,full_text")
            .eq("citation", citation)
            .execute()
        )
        mine = normalize_text(full_text)
        for row in same.data or []:
            if normalize_text(row.get("full_text")) == mine:
                return row["storage_path"]
        for row in same.data or []:
            NOTES.append(
                f"Warning: {citation} is also the citation of {row['storage_path']}, "
                "but the text differs, so it was treated as a different decision."
            )
    return None


def list_waiting(supabase) -> list:
    """Top-level PDFs in the incoming bucket. Skips folders (_skipped) and hidden files."""
    items = supabase.storage.from_(INCOMING_BUCKET).list("", {"limit": 1000, "offset": 0})
    names = []
    for item in items or []:
        name = item.get("name") or ""
        if name.startswith((".", "_")) or not name.lower().endswith(".pdf"):
            continue
        names.append(name)
    return sorted(names)


def choose_storage_path(name: str, decision_date) -> str:
    """Matches the existing archive layout ("2020/PERC 2020 33.pdf"): the year comes from
    the file name, then from the decision date, otherwise the file sits at the top level."""
    m = re.search(r"(?<!\d)((?:19|20)\d{2})(?!\d)", name)
    year = m.group(1) if m else (decision_date[:4] if decision_date else None)
    return f"{year}/{name}" if year else name


def process_one(supabase, claude, name: str, force: bool) -> str:
    """Returns 'added' or 'duplicate'. Raises on any problem so the file stays queued."""
    incoming = supabase.storage.from_(INCOMING_BUCKET)
    data = incoming.download(name)

    with tempfile.TemporaryDirectory() as tmp:
        local = Path(tmp) / "document.pdf"
        local.write_bytes(data)
        text_for_ai = cs.extract_pdf_text(str(local), max_chars=cs.MAX_CHARS_PER_PDF)
        full_text = cs.extract_pdf_text(str(local), max_chars=cs.MAX_CHARS_FOR_SEARCH)

    if full_text.startswith("[Could not extract text"):
        raise ValueError(full_text)
    if not full_text.strip():
        raise ValueError("no readable text in this PDF (it may be a scan without OCR)")

    decision_date = cs.extract_decision_date(full_text)
    storage_path = choose_storage_path(name, decision_date)

    file_hash = hashlib.sha256(data).hexdigest()
    citation = cs.extract_self_citation(full_text)

    existing = (
        supabase.table("documents").select("id").eq("storage_path", storage_path).limit(1).execute()
    )
    dup_of = storage_path if existing.data else find_duplicate(supabase, file_hash, citation, full_text)
    if dup_of and not force:
        NOTES.append(f"{name} is a duplicate of {dup_of} (moved to incoming/_skipped).")
        try:
            incoming.move(name, SKIPPED_PREFIX + name)
        except Exception:
            incoming.remove([name])
        return "duplicate"

    summary = cs.summarize_with_claude(claude, text_for_ai)
    if summary.get("summary") == UNPARSEABLE_SUMMARY:
        raise ValueError("Claude's reply could not be read as a summary")

    # Permanent copy first; if this fails the row is never written, so no dead links.
    supabase.storage.from_(cs.PDF_BUCKET).upload(
        storage_path, data, {"content-type": "application/pdf", "upsert": "true"}
    )

    supabase.table("documents").upsert(
        {
            "filename": name,
            "storage_path": storage_path,
            "file_hash": file_hash,
            "parties": summary.get("parties", "Not specified"),
            "summary": summary.get("summary", "Not specified"),
            "factors": summary.get("factors", []),
            "decision": summary.get("decision", "Not specified"),
            "category": summary.get("category", "Uncategorized"),
            "full_text": full_text,
            "decision_date": decision_date,
            "citation": citation,
            "related_citations": cs.extract_related_decisions(full_text),
            "updated_at": datetime.now(timezone.utc).isoformat(),
        },
        on_conflict="storage_path",
    ).execute()

    incoming.remove([name])
    return "added"


def rebuild_summaries(supabase) -> bool:
    """Rebuilds the Word and PDF summaries and uploads both. Returns False if the PDF step failed."""
    with tempfile.TemporaryDirectory() as tmp:
        docx_path = os.path.join(tmp, "Case_Summaries.docx")
        pdf_path = os.path.join(tmp, "Case_Summaries.pdf")

        cs.build_docx(supabase, docx_path)

        with open(docx_path, "rb") as f:
            supabase.storage.from_(cs.PDF_BUCKET).upload(
                DOCX_KEY, f.read(), {"content-type": DOCX_MIME, "upsert": "true"}
            )
        print("Word summary uploaded.")

        if not cs.convert_docx_to_pdf(docx_path, pdf_path):
            return False
        cs.upload_summary_pdf(supabase, pdf_path)
    return True


def write_job_summary(added, duplicates, failures, rebuilt):
    path = os.environ.get("GITHUB_STEP_SUMMARY")
    if not path:
        return
    lines = ["## PDF processing", ""]
    lines.append(f"- Added to the archive: **{len(added)}**")
    lines.append(f"- Already in the archive (moved to incoming/_skipped): **{len(duplicates)}**")
    lines.append(f"- Failed and left in incoming: **{len(failures)}**")
    lines.append(f"- Word and PDF summaries rebuilt: **{'yes' if rebuilt else 'no'}**")
    for name, why in failures:
        lines.append(f"  - {name}: {why}")
    for note in NOTES:
        lines.append(f"- {note}")
    with open(path, "a", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


def main():
    parser = argparse.ArgumentParser(description="Process new PDFs from the incoming bucket.")
    parser.add_argument("--count", action="store_true", help="Only report how many PDFs are waiting")
    parser.add_argument("--force", action="store_true", help="Re-summarize files already in the archive")
    parser.add_argument("--rebuild", action="store_true", help="Rebuild the Word/PDF summaries even if nothing is new")
    args = parser.parse_args()

    supabase = cs.get_supabase_client()
    waiting = list_waiting(supabase)

    if args.count:
        print(f"{len(waiting)} PDF(s) waiting in '{INCOMING_BUCKET}'.")
        out = os.environ.get("GITHUB_OUTPUT")
        if out:
            with open(out, "a", encoding="utf-8") as f:
                f.write(f"waiting={len(waiting)}\n")
        return

    added, duplicates, failures = [], [], []
    if waiting:
        import anthropic

        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            print("ERROR: Set the ANTHROPIC_API_KEY environment variable first.")
            sys.exit(1)
        claude = anthropic.Anthropic(api_key=api_key)

        for name in waiting:
            try:
                result = process_one(supabase, claude, name, args.force)
                (added if result == "added" else duplicates).append(name)
                print(f"{result}: {name}")
            except Exception as e:  # keep going; report at the end
                failures.append((name, str(e)))
                print(f"FAILED: {name}: {e}")
    else:
        print("Nothing waiting in 'incoming'.")

    rebuilt = False
    pdf_ok = True
    if added or args.rebuild or args.force:
        pdf_ok = rebuild_summaries(supabase)
        rebuilt = True

    write_job_summary(added, duplicates, failures, rebuilt)
    print(f"\nDone. Added {len(added)}, already in archive {len(duplicates)}, failed {len(failures)}.")
    if failures or not pdf_ok:
        sys.exit(1)


if __name__ == "__main__":
    main()
