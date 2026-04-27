"""看 output_full_test.xlsx 第 252, 684, 805, 868, 902, 903 列是哪些類別"""
import openpyxl, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r"C:\RenStudio\case\testcase\shopee\yahoo_converter\output_full_test.xlsx", read_only=True)
ws = wb["可上架商品"]
rows = list(ws.iter_rows(values_only=True))
headers = list(rows[0])
data = rows[1:]

ni = headers.index("商品名稱")
ci = headers.index("類別")
attr1_i = headers.index("屬性欄位1")

target_rows = [252, 684, 805, 868, 902, 903]
print("=== 錯誤列詳情 ===")
for r in target_rows:
    idx = r - 2
    if 0 <= idx < len(data):
        row = data[idx]
        print(f"\n[列 {r}] 類別: {row[ci]}")
        print(f"  商品: {row[ni]}")
        print(f"  屬性欄位1: {row[attr1_i]!r}")
