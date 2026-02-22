"""Debug subtotal calculation for section 01.2."""
import sys
sys.path.insert(0, r"d:\CTO")

from core.parser_dpgf import parse_dpgf

LOG = r"d:\CTO\debug_dpgf.txt"

dpgf_df, _ = parse_dpgf(r"d:\CTO\MAB_SUD_OUEST.xlsx")
f = open(LOG, "w", encoding="utf-8")

f.write("=== Section 01.2 children ===\n")
for _, r in dpgf_df.iterrows():
    code = r["Code"]
    if code and code.startswith("01.2"):
        qu = r["Qu."] or 0
        pu = r["Px_U_HT"] or 0
        tot = r["Px_Tot_HT"] or 0
        f.write(f"{r['row_type']:15} Code={code:15} Qu={qu:>10} PU={pu:>10.2f} Tot={tot:>10.2f} {r['Désignation'][:40]}\n")

f.write("\n=== Sums for 01.2 ===\n")
total_art = 0
total_sub = 0
for _, r in dpgf_df.iterrows():
    code = r["Code"]
    if code and code.startswith("01.2.") and r["row_type"] in ("article", "sub_section"):
        val = r["Px_Tot_HT"] or 0
        try:
            val = float(val)
        except (ValueError, TypeError):
            val = 0
        if r["row_type"] == "article":
            total_art += val
            f.write(f"  ART  {code:15} Tot={val:>10.2f}\n")
        else:
            total_sub += val
            f.write(f"  SUB  {code:15} Tot={val:>10.2f}\n")

f.write(f"\nTotal articles only: {total_art:.2f}\n")
f.write(f"Total sub_sections only: {total_sub:.2f}\n")
f.write(f"Total combined: {total_art + total_sub:.2f}\n")
f.write(f"Expected (TCO_FINAL ref): 8788.26\n")
f.close()
print("Done")
