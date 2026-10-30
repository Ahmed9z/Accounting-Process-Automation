"""
Invoice Merger
- Reads the basic info template Excel
- Reads the logistaas_results.xlsx
- Merges them by Reference into one Odoo-ready upload file
- Format: template data + first product on same row,
          subsequent products on next rows (template columns empty),
          empty row separator between invoices
"""

import os
import pandas as pd


# ============================================================
# Pick file
# ============================================================
def pick_excel_file(title: str) -> str:
    try:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        filepath = filedialog.askopenfilename(
            title=title,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        root.destroy()
        return filepath
    except:
        print(f"\n[INFO] {title}")
        print("Please type or paste the full path to your Excel file:")
        print("Example: C:\\Users\\akhai\\Downloads\\invoices.xlsx")
        path = input("Path: ").strip().strip('"')
        return path


# ============================================================
# Main
# ============================================================
def main():
    print("=" * 55)
    print("   Invoice Merger")
    print("=" * 55)

    # Step 1: Pick template file
    print("\n[INFO] Select your basic info template Excel file")
    print("       (the one with Invoice/Bill Date, Reference, Customer, etc.)")
    template_path = pick_excel_file("Select basic info template")
    if not template_path:
        print("[ERROR] No file selected.")
        input("Press Enter to exit...")
        return

    # Step 2: Pick logistaas_results file
    print("\n[INFO] Select your logistaas_results.xlsx file")
    results_path = pick_excel_file("Select logistaas_results.xlsx")
    if not results_path:
        print("[ERROR] No file selected.")
        input("Press Enter to exit...")
        return

    # Read files
    print("\n[INFO] Reading files...")
    template_df = pd.read_excel(template_path)
    results_df  = pd.read_excel(results_path)

    print(f"[INFO] Template columns : {list(template_df.columns)}")
    print(f"[INFO] Results columns  : {list(results_df.columns)}")

    # Find Reference column in template
    template_ref_col = next(
        (c for c in template_df.columns if "reference" in c.lower()), None
    )
    # Find Reference column in results
    results_ref_col = next(
        (c for c in results_df.columns if "reference" in c.lower()), None
    )

    if not template_ref_col:
        print("[ERROR] Could not find Reference column in template file.")
        input("Press Enter to exit...")
        return
    if not results_ref_col:
        print("[ERROR] Could not find Reference column in logistaas_results file.")
        input("Press Enter to exit...")
        return

    print(f"[INFO] Matched on: '{template_ref_col}' <-> '{results_ref_col}'")

    # Product line columns from results (everything except Reference)
    product_cols = [c for c in results_df.columns if c.lower() != results_ref_col.lower()]

    # Template columns
    template_cols = list(template_df.columns)

    # Final output columns = template cols + product cols
    output_cols = template_cols + product_cols

    # --------------------------------------------------------
    # Build merged rows
    # --------------------------------------------------------
    merged_rows = []
    empty_row   = {col: None for col in output_cols}

    # Preserve template order
    template_df[template_ref_col] = template_df[template_ref_col].astype(str).str.strip()
    results_df[results_ref_col]   = results_df[results_ref_col].astype(str).str.strip()

    refs_seen = set()
    for _, t_row in template_df.iterrows():
        ref = str(t_row[template_ref_col]).strip()
        if not ref or ref.lower() in ["nan", ""]:
            continue

        # Get all product lines for this reference
        products = results_df[results_df[results_ref_col] == ref]

        # Build template data dict
        t_data = {col: t_row[col] for col in template_cols}

        if products.empty:
            # No product lines found — write template row only
            row = {**t_data, **{col: None for col in product_cols}}
            merged_rows.append(row)
            print(f"  [WARN] No product lines found for: {ref}")
        else:
            first = True
            for _, p_row in products.iterrows():
                p_data = {col: p_row[col] for col in product_cols}

                if first:
                    # First product: template data + product on same row
                    row = {**t_data, **p_data}
                    first = False
                else:
                    # Subsequent products: empty template cols, only product data
                    row = {**{col: None for col in template_cols}, **p_data}

                merged_rows.append(row)

        # Add empty separator row after each invoice
        merged_rows.append(dict(empty_row))

    # Remove last empty row (not needed after final invoice)
    if merged_rows and all(v is None for v in merged_rows[-1].values()):
        merged_rows.pop()

    # --------------------------------------------------------
    # Save output
    # --------------------------------------------------------
    output_df   = pd.DataFrame(merged_rows, columns=output_cols)
    output_path = os.path.join(
        os.path.dirname(template_path), "odoo_upload.xlsx"
    )
    output_df.to_excel(output_path, index=False)

    print(f"\n[DONE] Merged file saved to: {output_path}")
    print(f"[DONE] Total rows (including separators): {len(merged_rows)}")
    print(f"[DONE] Ready to import into Odoo!")

    # Summary of unmatched references
    template_refs = set(template_df[template_ref_col].astype(str).str.strip())
    results_refs  = set(results_df[results_ref_col].astype(str).str.strip())
    unmatched     = template_refs - results_refs - {"nan", ""}
    if unmatched:
        print(f"\n[WARN] These references had no product lines in logistaas_results:")
        for r in sorted(unmatched):
            print(f"  - {r}")

    input("\nPress Enter to exit...")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n[CRASH] {e}")
        import traceback
        traceback.print_exc()
        input("\nPress Enter to exit...")
