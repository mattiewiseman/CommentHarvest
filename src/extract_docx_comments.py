# src/extract_docx_comments.py
import argparse
import zipfile
from pathlib import Path
from typing import Dict, List, Tuple
from lxml import etree as ET
import pandas as pd

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

def _qn(tag: str) -> str:
    # Qualify a WordprocessingML tag name, e.g. w:id
    if ":" in tag:
        pfx, local = tag.split(":", 1)
        return "{%s}%s" % (NS[pfx], local)
    return tag

def read_comments_xml(z: zipfile.ZipFile) -> Dict[int, dict]:
    """
    Parse word/comments.xml and return {id: {"text": str, "author": str|None, "date": str|None}}
    """
    try:
        with z.open("word/comments.xml") as f:
            root = ET.parse(f).getroot()
    except KeyError:
        # No comments part
        return {}

    comments: Dict[int, dict] = {}
    for c in root.findall("w:comment", namespaces=NS):
        cid = int(c.get(_qn("w:id")))
        author = c.get(_qn("w:author"))
        date = c.get(_qn("w:date"))
        # Concat all text nodes inside paragraphs/runs of this comment
        texts: List[str] = []
        for t in c.findall(".//w:t", namespaces=NS):
            if t.text:
                texts.append(t.text)
        comments[cid] = {
            "text": "".join(texts).strip(),
            "author": author,
            "date": date,
        }
    return comments

def read_commented_ranges(z: zipfile.ZipFile) -> Dict[int, str]:
    """
    Walk word/document.xml in document order; collect text inside each comment range id.
    Returns {id: commented_text}
    """
    with z.open("word/document.xml") as f:
        root = ET.parse(f).getroot()

    open_ranges: Dict[int, List[str]] = {}
    collected: Dict[int, str] = {}

    # Iterate in document order
    for elem in root.iter():
        tag = elem.tag
        # comment range start
        if tag == _qn("w:commentRangeStart"):
            cid = int(elem.get(_qn("w:id")))
            if cid not in open_ranges:
                open_ranges[cid] = []
        # text nodes
        elif tag == _qn("w:t"):
            text = elem.text or ""
            if text and open_ranges:
                for cid in list(open_ranges.keys()):
                    open_ranges[cid].append(text)
        # comment range end
        elif tag == _qn("w:commentRangeEnd"):
            cid = int(elem.get(_qn("w:id")))
            parts = open_ranges.pop(cid, [])
            collected[cid] = "".join(parts).strip()

    return collected

def extract_pairs(docx_path: Path) -> List[dict]:
    with zipfile.ZipFile(docx_path) as z:
        comments_map = read_comments_xml(z)  # id -> {text, author, date}
        ranges_map = read_commented_ranges(z)  # id -> commented_text

    rows: List[dict] = []
    for cid, cmeta in comments_map.items():
        commented_text = ranges_map.get(cid, "")
        rows.append({
            "id": cid,
            "Commented Text": commented_text,
            "Comment": cmeta.get("text", ""),
            "Author": cmeta.get("author"),
            "Date": cmeta.get("date"),
        })
    return rows

def main():
    parser = argparse.ArgumentParser(
        description="Export Word (.docx) comments to Excel: (Commented Text, Comment)."
    )
    parser.add_argument("input", type=str, help="Path to the .docx file")
    parser.add_argument("-o", "--output", type=str, help="Path to output .xlsx")
    parser.add_argument("--author", action="store_true", help="Include Author column")
    parser.add_argument("--date", action="store_true", help="Include Date column")
    parser.add_argument("--keep-empty", action="store_true", help="Keep rows with empty commented text")

    args = parser.parse_args()
    docx_path = Path(args.input)
    if not docx_path.exists():
        raise SystemExit(f"Input not found: {docx_path}")

    rows = extract_pairs(docx_path)

    # Default output path
    out_path = Path(args.output) if args.output else docx_path.with_suffix(".xlsx")

    # Build DataFrame with requested columns
    base_cols = ["Commented Text", "Comment"]
    if args.author:
        base_cols.append("Author")
    if args.date:
        base_cols.append("Date")

    df = pd.DataFrame(rows)
    if not args.keep_empty:
        df = df[df["Commented Text"].astype(str).str.strip() != ""]

    df = df[base_cols]

    # Write to Excel
    df.to_excel(out_path, index=False)
    print(f"Wrote: {out_path}")

if __name__ == "__main__":
    main()
