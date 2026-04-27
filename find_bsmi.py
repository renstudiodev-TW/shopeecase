import openpyxl, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r"C:\RenStudio\case\testcase\shopee\拍賣蝦皮拍賣商品_匯出商品資料_20260421_024920.xlsx", read_only=True, data_only=True)
ws = wb.active
headers = [c.value for c in next(ws.iter_rows(values_only=False))]

# 找所有跟檢驗、BSMI、字號相關的欄位
keywords = ["檢驗", "BSMI", "bsmi", "字號", "標識", "商檢"]
print("=== 含關鍵字的欄位 ===")
for i, h in enumerate(headers):
    if h and any(k in str(h) for k in keywords):
        print(f"  [{i}] {h}")

# 看完整欄位列表結尾段（屬性、規格之後可能有特殊欄位）
print(f"\n=== 最後 30 個欄位 ===")
for i, h in enumerate(headers[-30:], start=len(headers)-30):
    print(f"  [{i}] {h}")
