"""Quick debug script: extract raw text from a sick list PDF to inspect jour row format."""
import sys
import pdfplumber

if len(sys.argv) < 2:
    print("Usage: python debug_pdf_lines.py <sicklist.pdf>")
    sys.exit(1)

pdf_path = sys.argv[1]
with pdfplumber.open(pdf_path) as pdf:
    for i, page in enumerate(pdf.pages):
        text = page.extract_text() or ""
        print(f"\n=== PAGE {i+1} ===")
        for j, line in enumerate(text.splitlines()):
            # Show lines that start with a day number (1-31)
            stripped = line.strip()
            if stripped and stripped[0].isdigit():
                print(f"  LINE {j:3d}: {line!r}")
