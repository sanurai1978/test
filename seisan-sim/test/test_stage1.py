# 段階1の自動テスト：Excel 版（LibreOffice で実行）と HTML 版の数字を突き合わせる
#   使い方：python3 test_stage1.py <HTML 版のパス>
#   比べるもの：在庫一覧（在庫・引当・有効在庫・手配済・安全在庫割れ）、標準原価、メニューの件数
import sys, os, re, subprocess
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo_vba import Book
html = os.path.abspath(sys.argv[1])
out = subprocess.run(["node", os.path.join(os.path.dirname(__file__), "html_stage1.js"), html],
                     capture_output=True, text=True, check=True).stdout
num = lambda s: float(s.replace(",", "").replace("¥", ""))
exp = {}   # (事業部, 品番) → dict
cnt = {}
div = None
for line in out.splitlines():
    m = re.match(r"== 在庫 (\w+)", line)
    if m: div = {"kako": "加工", "kumi": "組立"}[m.group(1)]; continue
    m = re.match(r"== 標準原価 (\w+) (.*)", line)
    if m:
        d = {"kako": "加工", "kumi": "組立"}[m.group(1)]
        for part in m.group(2).split(" | "):
            c, v = part.split(" "); exp.setdefault((d, c), {})["標準原価"] = num(v)
        continue
    m = re.match(r"== 件数 (\d+)受注残 (\d+)要手配 (\d+)製造中 (\d+)出荷待ち", line)
    if m: cnt[div] = (int(m.group(1)), int(m.group(3)), int(m.group(4))); continue
    m = re.match(r"(\S+) 在庫([\d.,]+) 引当 ([\d.,]+) ／ 有効 (-?[\d.,]+)(.*)", line)
    if m:
        e = exp.setdefault((div, m.group(1)), {})
        e.update(在庫=num(m.group(2)), 引当計=num(m.group(3)), 有効在庫=num(m.group(4)))
        t = re.search(r"手配済 ([\d.,]+)", m.group(5)); e["手配済"] = num(t.group(1)) if t else 0.0
        e["注意"] = "安全在庫割れ" if "安全在庫割れ" in m.group(5) else ("欠品" if "欠品" in m.group(5) else "")
b = Book()
ng = 0
msg = b.run("Pr_初期化")
assert "初期データを入れました" in msg, msg
for d, macro in (("加工", "Pr_加工事業部に切替"), ("組立", "Pr_組立事業部に切替")):
    b.run(macro)
    b.run("Pr_在庫一覧更新")
    rows = b.table("在庫一覧")
    head = rows[3]
    col = {h: i for i, h in enumerate(head)}
    for r in rows[4:]:
        e = exp.get((d, r[0]))
        if e is None: print("HTML に無い品目", d, r[0]); ng += 1; continue
        for k in ("在庫", "引当計", "有効在庫", "手配済", "注意"):
            if r[col[k]] != e[k]: print(f"不一致 {d} {r[0]} {k}: Excel={r[col[k]]} HTML={e[k]}"); ng += 1
    got = (b.cell("メニュー", "C6"), b.cell("メニュー", "E6"), b.cell("メニュー", "F6"))
    if tuple(int(x) for x in got) != cnt[d]: print("件数 不一致", d, got, cnt[d]); ng += 1
for r in b.table("品目")[1:]:
    e = exp.get((r[0], r[1]))
    if e and abs(r[14] - e["標準原価"]) > 0.5: print("標準原価 不一致", r[1], r[14], e["標準原価"]); ng += 1
b.close()
print(f"段階1：比較 {len(exp)} 品目、不一致 {ng} 件")
sys.exit(1 if ng else 0)
