"""GUI smoke test：開啟視窗 1.5 秒後自動關閉，驗證沒有啟動錯誤"""
import sys, os, threading, time, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import App

app = App()

def auto_close():
    time.sleep(1.5)
    app.after(0, app.destroy)

threading.Thread(target=auto_close, daemon=True).start()
print("GUI 啟動中...")
try:
    app.mainloop()
    print("✅ GUI 啟動成功，無例外")
except Exception as e:
    print(f"❌ 啟動失敗: {e}")
    raise
