"""
Patch the Rank formula in the Scoring sheet so it survives sorting.

The original formula in V<n> contains a literal row number in the tie-breaker:
    ...(ROW(U$3:U$32)<n)...
When Excel sorts rows, the formula cell moves but the literal n stays, so
the tie-breaker references the wrong row → duplicate ranks / top rank ≠ 1.

Fix: replace the literal with ROW(), which always evaluates to the current
cell's row, regardless of where the row ended up after a sort.
"""
import pathlib, warnings, openpyxl

warnings.filterwarnings("ignore")
ROOT = pathlib.Path(__file__).resolve().parent.parent
XLSX = ROOT / "gnc_june_2026" / "gnc_intern_evaluation_v1.xlsx"

wb = openpyxl.load_workbook(XLSX)
ws = wb["Scoring"]

START, END = 3, 32
patched = 0
for r in range(START, END + 1):
    cell = ws[f"V{r}"]
    if not isinstance(cell.value, str) or not cell.value.startswith("="):
        continue
    new_formula = (
        f'=IF(OR(B{r}="",T{r}="Y"),"",'
        f'1+SUMPRODUCT(((U$3:U$32>U{r})+((U$3:U$32=U{r})*(ROW(U$3:U$32)<ROW())))'
        f'*(T$3:T$32<>"Y")*(B$3:B$32<>"")))'
    )
    if cell.value != new_formula:
        cell.value = new_formula
        patched += 1

wb.save(XLSX)
print(f"Patched {patched} rank formulas (rows {START}-{END}).")
