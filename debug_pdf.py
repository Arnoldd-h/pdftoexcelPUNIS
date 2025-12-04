
import pdfplumber
import sys
import os
import re
# Import the class from the main script
# We need to add the current directory to sys.path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))
from pdf_to_excel_apu import APUConverter

def debug_extraction(pdf_path):
    print(f"Debugging extraction for: {pdf_path}")
    converter = APUConverter(pdf_path)
    rubros = converter.extract_all_rubros()
    
    if not rubros:
        print("No rubros found.")
        return

    # Inspect the second rubro
    if len(rubros) > 1:
        r = rubros[1]
        print(f"\nRubro 2: {r['numero_rubro']} - {r['detalle']}")
        
        print("\n--- EQUIPOS ---")
        if r['equipos']:
            for i, eq in enumerate(r['equipos']):
                print(f"Item {i}:")
                for k, v in eq.items():
                    print(f"  {k}: {v} (Type: {type(v)})")
        else:
            print("No equipos found.")
    else:
        print("Rubro 2 not found.")

if __name__ == "__main__":
    from pathlib import Path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    pdf_files = list(Path(current_dir).glob("*.pdf"))
    if pdf_files:
        debug_extraction(str(pdf_files[0]))
    else:
        print("No PDF found")
