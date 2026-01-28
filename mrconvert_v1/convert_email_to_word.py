#!/usr/bin/env python3
"""
Convert email markdown file to Word document and/or Outlook MSG format.

This script converts the AGI Transformers Transportation_email.md file
to a Word (.docx) document and optionally to an Outlook (.msg) file.
"""

from pathlib import Path
import sys
import argparse

from src.mrconvert import markdown_to_docx, docx_to_msg, markdown_to_xlsx


def main():
    """Convert email markdown file to Word document and/or Outlook MSG"""
    parser = argparse.ArgumentParser(
        description="Convert email markdown file to Word (.docx) and/or Outlook (.msg) format"
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        default="AGI Transformers Transportation_email.md",
        help="Input markdown file (default: AGI Transformers Transportation_email.md)"
    )
    parser.add_argument(
        "--to-docx",
        action="store_true",
        default=True,
        help="Convert to DOCX format (default: True)"
    )
    parser.add_argument(
        "--to-msg",
        action="store_true",
        help="Also convert to MSG format (requires Windows and Outlook)"
    )
    parser.add_argument(
        "--docx-only",
        action="store_true",
        help="Convert to DOCX only (no MSG)"
    )
    parser.add_argument(
        "--to-xlsx",
        action="store_true",
        help="Also convert to Excel XLSX format"
    )
    
    args = parser.parse_args()
    
    # Default input file
    script_dir = Path(__file__).parent
    input_file = Path(args.input_file)
    
    if not input_file.is_absolute():
        input_file = script_dir / input_file
    
    if not input_file.exists():
        print(f"Error: Input file not found: {input_file}")
        print(f"Usage: python convert_email_to_word.py [input_file.md] [--to-msg]")
        sys.exit(1)
    
    # Determine output formats
    to_docx = args.to_docx and not args.docx_only
    to_msg = args.to_msg
    to_xlsx = args.to_xlsx
    
    results = []
    
    # Convert to DOCX
    if to_docx:
        output_file = input_file.with_suffix('.docx')
        print(f"Converting to DOCX: {input_file.name} → {output_file.name}")
        
        try:
            result = markdown_to_docx.markdown_to_docx(
                src=input_file,
                dst=output_file,
                preserve_metadata=True
            )
            results.append(("DOCX", result))
            print(f"✓ Successfully converted to: {result.output}")
            print(f"  Engine: {result.engine}")
        except Exception as e:
            print(f"✗ Error converting to DOCX: {e}")
            import traceback
            traceback.print_exc()
    
    # Convert to MSG
    if to_msg:
        msg_file = input_file.with_suffix('.msg')
        print(f"\nConverting to MSG: {input_file.name} → {msg_file.name}")
        
        try:
            result = docx_to_msg.markdown_to_msg(
                src=input_file,
                dst=msg_file
            )
            results.append(("MSG", result))
            print(f"✓ Successfully converted to: {result.output}")
            print(f"  Engine: {result.engine}")
        except Exception as e:
            print(f"✗ Error converting to MSG: {e}")
            print(f"  Note: MSG conversion requires Windows and Microsoft Outlook installed.")
            print(f"  Install pywin32 with: pip install pywin32")
            import traceback
            traceback.print_exc()
    
    # Convert to XLSX
    if to_xlsx:
        xlsx_file = input_file.with_suffix('.xlsx')
        print(f"\nConverting to XLSX: {input_file.name} → {xlsx_file.name}")
        
        try:
            result = markdown_to_xlsx.markdown_to_xlsx(
                src=input_file,
                dst=xlsx_file,
                preserve_metadata=True
            )
            results.append(("XLSX", result))
            print(f"✓ Successfully converted to: {result.output}")
            print(f"  Engine: {result.engine}")
        except Exception as e:
            print(f"✗ Error converting to XLSX: {e}")
            print(f"  Note: XLSX conversion requires openpyxl library.")
            print(f"  Install with: pip install openpyxl")
            import traceback
            traceback.print_exc()
    
    # Summary
    if results:
        print(f"\n{'='*60}")
        print("Conversion Summary:")
        for format_name, result in results:
            print(f"  {format_name}: {result.output}")
        print(f"\nYou can now open these files in Microsoft Word or Outlook.")


if __name__ == "__main__":
    main()

