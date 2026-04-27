import openpyxl, sys, io
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r"C:\RenStudio\case\testcase\shopee\拍賣蝦皮拍賣商品_匯出商品資料_20260421_024920.xlsx", read_only=True, data_only=True)
ws = wb.active

rows = ws.iter_rows(values_only=True)
headers = list(next(rows))
data = [r for r in rows]
print(f"=== 商品總數: {len(data)} 筆 ===\n")

def col(name):
    try: return headers.index(name)
    except ValueError: return -1

required_yahoo = ["類別", "商品名稱", "商品描述", "定價", "商品數量", "刊登天數", "商品新舊", "所在地區", "圖片1"]
print("=== 必填欄位填寫率 ===")
for f in required_yahoo:
    i = col(f)
    if i < 0:
        print(f"  ❌ {f}: 欄位不存在！")
        continue
    filled = sum(1 for r in data if r[i] not in (None, "", " "))
    pct = filled*100//len(data) if data else 0
    flag = "✅" if pct == 100 else ("⚠️ " if pct > 50 else "❌")
    print(f"  {flag} {f}: {filled}/{len(data)} ({pct}%)")

print("\n=== 類別欄位範例 (前 10 筆不同值) ===")
cat_i = col("類別")
cats = Counter(str(r[cat_i]) for r in data if r[cat_i] is not None)
for c, n in list(cats.most_common(10)):
    print(f"  {c!r}  x {n}")

print("\n=== 屬性欄位 1~10 填寫率 (Yahoo 必填規格 / BSMI 通常放這) ===")
for k in range(1, 11):
    i = col(f"屬性欄位{k}")
    if i < 0: continue
    filled = sum(1 for r in data if r[i] not in (None, ""))
    print(f"  屬性欄位{k}: {filled}/{len(data)} 筆有填")

print("\n=== 第一筆商品完整內容 ===")
r = data[0]
for i, h in enumerate(headers):
    v = r[i]
    if v is None or v == "": continue
    s = str(v).replace("\n", " | ")[:120]
    print(f"  [{i:3d}] {h}: {s}")

print("\n=== 商品名稱前 5 筆 ===")
ni = col("商品名稱")
for r in data[:5]:
    print(f"  - {r[ni]}")

print("\n=== 圖片1 前 3 筆 (檢查 URL/檔名格式) ===")
ii = col("圖片1")
for r in data[:3]:
    print(f"  - {r[ii]}")
