#!/usr/bin/env python3
"""
PDF to Markdown converter using pdfplumber
Converts PDF to MD format following mrconvert conventions
"""

import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List

try:
    import pdfplumber
except ImportError:
    print("Error: pdfplumber is not installed. Install it with: pip install pdfplumber")
    sys.exit(1)


def extract_pdf_to_markdown(pdf_path: Path, keep_layout: bool = False) -> Dict[str, Any]:
    """Extract text from PDF and convert to markdown format"""
    
    meta: Dict[str, Any] = {
        "source": str(pdf_path),
        "type": "pdf",
        "pages": 0,
        "parsed_at": datetime.now().isoformat() + "Z",
        "ocr": {"used": False, "engine": "none", "lang": None},
    }
    
    text_parts: List[str] = []
    tables: List[Dict[str, Any]] = []
    
    with pdfplumber.open(pdf_path) as pdf:
        meta["pages"] = len(pdf.pages)
        
        for i, page in enumerate(pdf.pages, start=1):
            # Extract text
            text = page.extract_text(
                x_tolerance=1.5 if keep_layout else 3.0,
                y_tolerance=1.5 if keep_layout else 3.0
            )
            
            if text:
                # Add page header
                text_parts.append(f"## Page {i}\n\n{text}")
            else:
                text_parts.append(f"## Page {i}\n\n*[No text content]*")
            
            # Extract tables
            try:
                page_tables = page.extract_tables()
                for ti, tbl in enumerate(page_tables or []):
                    tables.append({"page": i, "index": ti, "rows": tbl})
                    
                    # Convert table to markdown format
                    if tbl and len(tbl) > 0:
                        md_table = "\n\n### Table {}-{}\n\n".format(i, ti)
                        # Header row
                        if len(tbl) > 0:
                            md_table += "| " + " | ".join(str(cell or "") for cell in tbl[0]) + " |\n"
                            md_table += "| " + " | ".join("---" for _ in tbl[0]) + " |\n"
                        # Data rows
                        for row in tbl[1:]:
                            md_table += "| " + " | ".join(str(cell or "") for cell in row) + " |\n"
                        text_parts.append(md_table)
            except Exception:
                pass
    
    # Combine all text parts
    markdown_content = "\n\n".join(text_parts).strip()
    
    # Add metadata header
    metadata_header = f"""# {pdf_path.stem}

**Source:** `{pdf_path.name}`  
**Type:** PDF  
**Pages:** {meta['pages']}  
**Parsed At:** {meta['parsed_at']}  

---

"""
    
    full_markdown = metadata_header + markdown_content
    
    return {
        "meta": meta,
        "text": "\n\n".join(text_parts).replace("## ", "").replace("### ", "").strip(),
        "markdown": full_markdown,
        "tables": tables,
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: python convert_pdf_to_md.py <pdf_file> [output_file.md]")
        sys.exit(1)
    
    pdf_path = Path(sys.argv[1])
    if not pdf_path.exists():
        print(f"Error: File not found: {pdf_path}")
        sys.exit(1)
    
    if not pdf_path.suffix.lower() == ".pdf":
        print(f"Error: Not a PDF file: {pdf_path}")
        sys.exit(1)
    
    # Determine output path
    if len(sys.argv) >= 3:
        output_path = Path(sys.argv[2])
    else:
        output_path = pdf_path.parent / f"{pdf_path.stem}.md"
    
    print(f"Converting {pdf_path.name} to Markdown...")
    
    try:
        result = extract_pdf_to_markdown(pdf_path, keep_layout=False)
        
        # Write markdown file
        output_path.write_text(result["markdown"], encoding="utf-8")
        
        print(f"[OK] Successfully converted to: {output_path}")
        print(f"  Pages: {result['meta']['pages']}")
        print(f"  Tables found: {len(result['tables'])}")
        
        # Also save JSON if requested
        if "--json" in sys.argv:
            import json
            json_path = output_path.with_suffix(".json")
            json_path.write_text(
                json.dumps(result, ensure_ascii=False, indent=2),
                encoding="utf-8"
            )
            print(f"[OK] JSON saved to: {json_path}")
        
    except Exception as e:
        print(f"Error during conversion: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

