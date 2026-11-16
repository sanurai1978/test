# 段階2の自動テスト：HTML 版と同じ操作を Excel 版（LibreOffice で実行）で行い、
# 各時点の MRP 手配案・製造指図・発注・メニューの件数を突き合わせる。
#   使い方：python3 test_stage2.py <HTML 版のパス>
import sys, os, json, subprocess, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo_vba import Book
html = os.path.abspath(sys.argv[1])
exp = json.loads(subprocess.run(["node", os.path.join(os.path.dirname(os.path.abspath(__file__)), "html_stage2.js"), html],
                                capture_output=True, text=True, check=True).stdout)
md = lambda v: (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).strftime("%-m/%-d") if isinstance(v, float) else v
b = Book()
ng = 0
def check(key, div):
    global ng
    e = exp[key]
    # MRP シート（計算し直して読む）
    b.run("Pr_MRP計算")
    rows = b.table("MRP")
    got = []
    for r in rows[8 - 1 + 1:]:          # 9行目から（table は1行目から）
        if r[3] == "" or r[2] not in ("製造", "発注"): break
        got.append({"kind": r[2], "code": r[3], "qty": r[5], "st": "今日" if r[2] == "発注" else md(r[7]), "du": md(r[8]),
                    "nd": "補充" if r[9] == "安全在庫の補充" else md(r[9]), "so": r[10], "late": "間に合わない" in r[11]})
    want = [{k: x[k] for k in ("kind", "code", "qty", "st", "du", "nd", "so", "late")} for x in e["mrp"]]
    if got != want:
        ng += 1; print(f"[{key}] MRP 不一致\n  Excel={got}\n  HTML ={want}")
    # 製造指図（HTML は新しい順）
    mo = [r for r in b.table("製造指図")[1:] if r[0] == div][::-1]
    got = [(r[1], r[3], r[4], md(r[5]), md(r[6]), r[7]) for r in mo]
    want = [(m["no"], m["qty"], m["st"], m["start"], m["due"], m["so"]) for m in e["mo"]]
    if got != want: ng += 1; print(f"[{key}] 製造指図 不一致\n  Excel={got}\n  HTML ={want}")
    po = [r for r in b.table("発注")[1:] if r[0] == div][::-1]
    got = [(r[1], r[3], md(r[4])) for r in po]
    want = [(x["no"], x["qty"], x["eta"]) for x in e["po"]]
    if got != want: ng += 1; print(f"[{key}] 発注 不一致\n  Excel={got}\n  HTML ={want}")
    st = [b.cell("メニュー", a) for a in ("C6", "D6", "E6", "F6")]
    if [int(x) for x in st] != e["stats"]: ng += 1; print(f"[{key}] 件数 不一致 Excel={st} HTML={e['stats']}")
    b.run("Pr_負荷一覧更新")
    lt = b.table("負荷一覧")
    got = []
    for r in lt[5:]:
        if r[1] == "": break
        got.append([r[2]] + [int(x) for x in r[3:17]])
    if got != e["load"]: ng += 1; print(f"[{key}] 負荷 不一致\n  Excel={got}\n  HTML ={e['load']}")
    print(f"[{key}] 確認済み")
def put(sheet, cells):
    for a, v in cells.items(): b.set_cell(sheet, a, v)
def run(macro, ok_text=None, answers=None):
    m = b.run(macro, answers)
    if ok_text is not None and "[MSG]" in m: print("  想定外のメッセージ:", macro, m)
    return m

assert "初期データを入れました" in b.run("Pr_初期化")
b.run("Pr_加工事業部に切替"); check("A_加工_初期", "加工")
run("Pr_MRP一括手配", ""); check("B_加工_一括手配後", "加工")
put("入力_受注", {"C6": "P-103", "C7": 5, "C8": "北斗精工", "C9": 4}); run("Pr_受注登録", ""); check("C_加工_受注P103", "加工")
put("入力_受注", {"C12": "SO-0005"}); run("Pr_受注向け製造指図作成", ""); check("D_加工_受注向け指図", "加工")
put("入力_見込", {"C6": "P-101", "C7": 50, "C8": 20}); run("Pr_見込登録", ""); check("E_加工_見込追加", "加工")
put("入力_製造", {"C11": "MO-0001", "C12": 5}); run("Pr_着手日変更", ""); check("F_加工_着手日変更", "加工")
b.run("Pr_組立事業部に切替"); check("G_組立_初期", "組立")
put("入力_製造", {"C11": "MO-0002"}); run("Pr_製造指図取消", "", answers=[6]); check("H_組立_指図取消", "組立")
pk = [r for r in b.table("ピッキング")[1:] if r[0] == "組立"]
if pk: ng += 1; print("ピッキングが消えていない", pk)
put("入力_受注", {"C6": "A-301", "C7": 2, "C8": "東和機械", "C9": 30}); run("Pr_受注登録", ""); check("I_組立_受注A301", "組立")
run("Pr_MRP一括手配", ""); check("J_組立_一括手配後", "組立")
put("入力_発注", {"C6": "B-004", "C7": 30}); run("Pr_発注登録", ""); check("K_組立_手動発注", "組立")
# エラーになるべき操作（メッセージが出て、データが変わらない）
for sheet, cells, macro, word in [
    ("入力_受注", {"C12": "SO-0001"}, "Pr_受注引当", "有効在庫が足りません"),
    ("入力_受注", {"C6": "S-101", "C7": 1, "C9": 3}, "Pr_受注登録", "加工品・完成品だけ"),
    ("入力_受注", {"C6": "A-201", "C7": 1.5, "C9": 3}, "Pr_受注登録", "整数"),
    ("入力_見込", {"C6": "A-301", "C7": 5, "C8": 3}, "Pr_見込登録", "見込生産"),
    ("入力_製造", {"C11": "MO-0001"}, "Pr_製造指図取消", "「計画」"),
    ("入力_発注", {"C6": "A-201", "C7": 5}, "Pr_発注登録", "購入品"),
]:
    put(sheet, cells); m = b.run(macro)
    if word not in m: ng += 1; print("エラーにならない／文言が違う:", macro, cells, m)
    else: print("  期待どおりエラー:", macro, word)
# 選んだ行だけ手配：受注を足して MRP を計算 → 1行目にだけ○ → その1件だけ作られる
put("入力_受注", {"C6": "A-202", "C7": 20, "C8": "北斗精工", "C9": 9}); run("Pr_受注登録", "")
b.run("Pr_MRP計算"); n_before = len([r for r in b.table("製造指図")[1:] if r[0] == "組立"]) + len([r for r in b.table("発注")[1:] if r[0] == "組立"])
b.set_cell("MRP", "B9", "○"); m = b.run("Pr_MRP選択手配")
n_after = len([r for r in b.table("製造指図")[1:] if r[0] == "組立"]) + len([r for r in b.table("発注")[1:] if r[0] == "組立"])
if n_after != n_before + 1: ng += 1; print("選んだ行の手配 不一致", n_before, n_after, m)
else: print("  選んだ1行だけ手配された:", b.cell("MRP", "B3"))
# 表示が古いと手配しない（受注を足してから、計算し直さずに押す）
put("入力_受注", {"C6": "A-201", "C7": 100, "C8": "北斗精工", "C9": 9}); run("Pr_受注登録", "")
b.set_cell("MRP", "B9", "○"); m = b.run("Pr_MRP選択手配")
if "違っていました" not in m: ng += 1; print("古い手配案を止めない", m)
else: print("  古い手配案は止めた")
b.run("Pr_負荷一覧更新"); print("  負荷一覧:", b.cell("負荷一覧", "B3"))
b.close()
print(f"段階2：時点 {len([k for k in exp if not k.startswith('L_')])} 件を比較、不一致 {ng} 件")
sys.exit(1 if ng else 0)
