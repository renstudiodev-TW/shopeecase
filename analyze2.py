import openpyxl, sys, io, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

wb = openpyxl.load_workbook(r"C:\RenStudio\case\testcase\shopee\拍賣蝦皮拍賣商品_匯出商品資料_20260421_024920.xlsx", read_only=True, data_only=True)
ws = wb.active
rows = ws.iter_rows(values_only=True)
headers = list(next(rows))
data = [r for r in rows]

def col(name):
    try: return headers.index(name)
    except ValueError: return -1

# 找出商品數量空值的那一筆
qi = col("商品數量")
ni = col("商品名稱")
print("=== 商品數量空值的商品 ===")
for idx, r in enumerate(data):
    if r[qi] in (None, ""):
        print(f"  第 {idx+2} 列 ({r[ni]})")

# 商品類型 / 商品類型資訊
ti = col("商品類型")
ti2 = col("商品類型資訊")
print(f"\n=== 商品類型 ===")
from collections import Counter
print("  商品類型 分布:", dict(Counter(r[ti] for r in data).most_common(5)))
print("  商品類型資訊 分布:", dict(Counter(r[ti2] for r in data).most_common(5)))

# 物流方式統計（非輕鬆付那組）
print("\n=== 物流方式 yes 統計 ===")
for f in ["套用全店運費","郵寄掛號","全家取貨付款","7-11取貨付款","萊爾富取貨付款","取貨付款","郵局貨到付款","宅配","低溫寄送","面交/自取/不寄送"]:
    i = col(f)
    if i < 0: continue
    yes = sum(1 for r in data if str(r[i]).lower() == "yes")
    print(f"  {f}: yes = {yes}/{len(data)}")

# 規格資訊
print("\n=== 多規格商品統計 ===")
si = col("第一層規格:名稱")
multi = sum(1 for r in data if r[si] not in (None, ""))
print(f"  有第一層規格的商品: {multi}/{len(data)}")

# 圖片實體檔案是否存在
print("\n=== 檢查圖片實體檔案 ===")
ii = col("圖片1")
sample = [r[ii] for r in data[:5] if r[ii]]
folder = r"C:\RenStudio\case\testcase\shopee"
for s in sample:
    found = False
    # 嘗試在當前目錄、photos子目錄、images子目錄找
    for sub in ["", "photos", "images", "圖片", "_extract"]:
        path = os.path.join(folder, sub, s)
        if os.path.exists(path):
            print(f"  ✅ {s} 存在於 {sub or '(根目錄)'}")
            found = True
            break
    if not found:
        print(f"  ❌ {s} 找不到實體檔")

# 列一下根目錄結構
print("\n=== shopee 資料夾內容 ===")
for f in os.listdir(folder):
    p = os.path.join(folder, f)
    if os.path.isdir(p):
        print(f"  📁 {f}/")
    else:
        print(f"  📄 {f} ({os.path.getsize(p)} bytes)")
