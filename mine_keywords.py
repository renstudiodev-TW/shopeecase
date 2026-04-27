import openpyxl, sys, io, re
from collections import Counter
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r"C:\RenStudio\case\testcase\shopee\拍賣蝦皮拍賣商品_匯出商品資料_20260421_024920.xlsx", read_only=True, data_only=True)
ws = wb.active
rows = ws.iter_rows(values_only=True)
headers = list(next(rows))
data = [r for r in rows]

ni = headers.index("商品名稱")
ci = headers.index("類別")
di = headers.index("商品描述")

# 1. 看類別代碼分布
print("=== 類別代碼分布（前 15 個）===")
cats = Counter(r[ci] for r in data)
for cat, n in cats.most_common(15):
    sample_names = [r[ni] for r in data if r[ci] == cat][:2]
    print(f"  {cat}  ×{n}  範例: {sample_names[0][:30] if sample_names else '?'}")

# 2. 商品名稱關鍵字頻率
print("\n=== 商品名稱高頻關鍵字（2-4 字）===")
all_names = " ".join(str(r[ni]) for r in data if r[ni])
# 中文片段
chinese_chunks = re.findall(r'[一-鿿]{2,4}', all_names)
chunks_freq = Counter(chinese_chunks)
for word, n in chunks_freq.most_common(40):
    print(f"  {word}: {n}")

# 3. 拆出英文/型號
print("\n=== 商品名稱常見英文/型號（前 30）===")
eng = re.findall(r'[A-Z][A-Za-z0-9\-]{2,15}', all_names)
eng_freq = Counter(eng)
for word, n in eng_freq.most_common(30):
    print(f"  {word}: {n}")
