"""
Stage 1: Convert every PDF in gnc_june_2026/ to plain text (cached on disk).

Deterministic, no LLM calls. Run once; re-run is free since text files cache.
Output goes to cv_text_cache/<original_pdf_name>.txt
"""
import os, sys, pathlib
import pdfplumber

ROOT = pathlib.Path(__file__).resolve().parent.parent
PDF_DIR = ROOT / "gnc_june_2026"
OUT_DIR = ROOT / "cv_text_cache"
OUT_DIR.mkdir(exist_ok=True)


def extract(pdf_path: pathlib.Path) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            try:
                txt = page.extract_text() or ""
            except Exception as e:
                txt = f"[PAGE {i} EXTRACTION ERROR: {e}]"
            parts.append(f"\n=== PAGE {i+1} ===\n{txt}")
    return "\n".join(parts)


def main():
    pdfs = sorted(PDF_DIR.glob("*.pdf"))
    print(f"Found {len(pdfs)} PDFs")
    for p in pdfs:
        out = OUT_DIR / (p.stem + ".txt")
        if out.exists() and out.stat().st_size > 0 and "--force" not in sys.argv:
            continue
        try:
            txt = extract(p)
            out.write_text(txt, encoding="utf-8")
            print(f"  ok  {p.name}  ({len(txt)} chars)")
        except Exception as e:
            print(f"  ERR {p.name}: {e}")
            out.write_text(f"[EXTRACTION FAILED: {e}]", encoding="utf-8")


if __name__ == "__main__":
    main()
