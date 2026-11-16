# 段階3の自動テスト：HTML 版と同じ操作を Excel 版（LibreOffice で実行）で行い、
# 各時点の在庫・受注・製造指図・工程・外注・ピッキング・出荷・見込・発注・履歴を突き合わせる。
#   使い方：python3 test_stage3.py <HTML 版のパス>
import sys, os, json, subprocess, datetime
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo_vba import Book
html = os.path.abspath(sys.argv[1])
res = json.loads(subprocess.run(["node", os.path.join(os.path.dirname(os.path.abspath(__file__)), "html_stage3.js"), html],
                                capture_output=True, text=True, check=True).stdout)
exp, errs = res["out"], res["errs"]
iso = lambda v: (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))).isoformat() if isinstance(v, float) else v
num = lambda v: None if v in ("", None) else float(v)
DIV = {"加工": "kako", "組立": "kumi"}
b = Book()
ng = 0

def rows(sheet, div):
    return [r for r in b.table(sheet)[1:] if r[0] == div]

def excel_state(div):
    st = {}
    st["stock"] = {r[1]: r[13] for r in rows("品目", div)}
    st["so"] = [(r[1], r[2], r[3], r[4], r[5], iso(r[6]), r[7]) for r in rows("受注", div)][::-1]
    st["mo"] = [(r[1], r[2], r[3], r[4], r[8], r[7], num(r[9]) if r[4] == "完成" else None, r[10]) for r in rows("製造指図", div)][::-1]
    ops = {}
    for r in rows("指図工程", div):
        done = r[12] == "完了"
        ops.setdefault(r[1], []).append((int(r[2]), str(int(r[3])), r[12], num(r[13]) if done else None, num(r[14]) if done else None, num(r[15]) if done else None))
    st["ops"] = ops
    st["sc"] = [(r[1], r[2], int(r[3]), r[4], r[5], r[6], r[7], r[8], iso(r[9]), r[10], num(r[11]), num(r[12])) for r in rows("外注注文", div)][::-1]
    st["pk"] = [(r[1], r[2], r[3], r[4], r[5], r[6], r[7]) for r in rows("ピッキング", div)][::-1]
    st["pkl"] = {}
    for r in rows("ピッキング明細", div): st["pkl"].setdefault(r[1], []).append((r[3], r[4], r[5], r[6] == "○"))
    st["sh"] = [(r[1], r[2], r[3], r[4], r[5], iso(r[6]), r[7], r[8]) for r in rows("出荷実績", div)][::-1]
    st["fc"] = sorted((int(r[1]), r[2], r[3], iso(r[4])) for r in rows("見込", div))
    st["po"] = [(r[1], r[2], r[3], iso(r[4]), r[5]) for r in rows("発注", div)][::-1]
    st["log"] = [(r[2], r[3], r[4], r[5], r[6]) for r in rows("履歴", div)][::-1]
    return st

def html_state(d):
    st = {}
    st["stock"] = {i["code"]: float(i["stock"]) for i in d["items"]}
    st["so"] = [(o["no"], o["item"], float(o["qty"]), float(o.get("shipped") or 0), o["cust"], o["due"], o["st"]) for o in d["so"]]
    st["mo"] = [(m["no"], m["item"], float(m["qty"]), m["st"], m.get("pk") or "", m.get("so") or "", float(m["out"]) if m["st"] == "完成" else None, m.get("note") or "") for m in d["mo"]]
    st["ops"] = {m["no"]: [(o["seq"], o["tno"], o["st"], *( (float(o["good"]), float(o["ng"]), float(o["min"])) if o["st"] == "完了" else (None, None, None))) for o in m["ops"]] for m in d["mo"]}
    st["sc"] = [(s["no"], s["mo"], s["seq"], s["opname"], s["item"], s["vendor"], float(s["qty"]), float(s["price"]), s["eta"], s["st"],
                 float(s["good"]) if "good" in s else None, float(s["ng"]) if "ng" in s else None) for s in d["sc"]]
    st["pk"] = [(k["no"], k["kind"], k["ref"], k["title"], k["st"], k.get("carrier") or "", k.get("sh") or "") for k in d["pk"]]
    st["pkl"] = {k["no"]: [(l["item"], float(l["qty"]), l["loc"], bool(l["ok"])) for l in k["lines"]] for k in d["pk"]}
    st["sh"] = [(s["no"], s["so"], s["item"], float(s["qty"]), s["cust"], s["date"], s["carrier"], float(s["price"])) for s in d["sh"]]
    st["fc"] = sorted((f["id"], f["item"], float(f["qty"]), f["due"]) for f in d["fc"])
    st["po"] = [(x["no"], x["item"], float(x["qty"]), x["eta"], x["st"]) for x in d["po"]]
    st["log"] = [(x["kind"], x["code"], float(x["delta"]), float(x["bal"]), x["ref"]) for x in d["log"]]
    return st

def check(key, div):
    global ng
    e = html_state(exp[key][DIV[div]]); g = excel_state(div)
    bad = [k for k in e if e[k] != g[k]]
    for k in bad:
        ng += 1; print(f"[{key}] {k} 不一致\n  Excel={g[k]}\n  HTML ={e[k]}")
    print(f"[{key}] {'一致' if not bad else '不一致あり'}")

def put(sheet, cells):
    for a, v in cells.items(): b.set_cell(sheet, a, v)
def run(macro, answers=None, expect_ok=True):
    m = b.run(macro, answers)
    if expect_ok and "[MSG]" in m: print("  想定外のメッセージ:", macro, m.strip()[:200])
    return m
def tno(div, mo, seq):
    for r in rows("指図工程", div):
        if r[1] == mo and int(r[2]) == seq: return str(int(r[3]))
def op(div, mo, seq, vals=None):
    put("入力_製造", {"C17": tno(div, mo, seq)}); run("Pr_手配No読込")
    if vals: put("入力_製造", {"C18": vals[0], "C19": vals[1], "C20": vals[2]})
    run("Pr_工程実績登録")
def newest_pk(div): return rows("ピッキング", div)[-1][1]
def pick_all(no):
    put("入力_ピッキング", {"C6": no}); run("Pr_ピッキング読込"); run("Pr_ピッキング全部済"); run("Pr_ピッキング確定")

assert "初期データを入れました" in b.run("Pr_初期化")
b.run("Pr_加工事業部に切替")
put("入力_外注", {"C6": "SC-0001", "C7": 38, "C8": 2}); run("Pr_外注受入"); check("S1_外注受入", "加工")
op("加工", "MO-0002", 40, (37, 1, 15)); check("S2_実績で完成", "加工")
put("入力_製造", {"C11": "MO-0001"}); run("Pr_製造ピッキング指示"); pick_all(newest_pk("加工"))
op("加工", "MO-0001", 10, (20, 0, 95))
put("入力_製造", {"C17": tno("加工", "MO-0001", 20)}); run("Pr_外注に出す")
put("入力_外注", {"C6": rows("外注注文", "加工")[-1][1], "C7": 19, "C8": 1}); run("Pr_外注受入")
op("加工", "MO-0001", 30, (19, 0, 15)); check("S3_払出から完成まで", "加工")
pick_all("PK-0001"); check("S4_出荷ピッキング確定", "加工")
put("入力_受注", {"C6": "P-101", "C7": 30, "C8": "北斗精工", "C9": 5}); run("Pr_受注登録")
put("入力_受注", {"C12": "SO-0005"}); run("Pr_受注引当")
put("入力_出荷", {"C6": "SO-0005", "C7": 20, "C8": "宅配便"}); run("Pr_出荷ピッキング指示"); pick_all(newest_pk("加工")); check("S5_分納出荷", "加工")
put("入力_出荷", {"C11": "SH-0002"}); run("Pr_出荷取消", answers=[6]); check("S6_出荷取消", "加工")
put("入力_在庫調整", {"C6": "M-001", "C7": 390}); run("Pr_在庫調整"); check("S7_在庫調整", "加工")
put("入力_受注", {"C12": "SO-0004"}); run("Pr_受注向け製造指図作成")
put("入力_製造", {"C11": "MO-0003"}); run("Pr_製造ピッキング指示"); pick_all(newest_pk("加工"))
for s in (10, 20, 30): op("加工", "MO-0003", s)
check("S8_受注生産の完成と自動引当", "加工")
b.run("Pr_組立事業部に切替")
pick_all("PK-0001")
for s in (10, 20, 30): op("組立", "MO-0002", s)
check("S9_組立_払出から完成", "組立")
op("組立", "MO-0001", 20, (9, 1, 30)); check("S10_組立_実績", "組立")
put("入力_発注", {"C10": "PO-0001"}); run("Pr_入荷計上"); check("S11_入荷", "組立")
put("入力_出荷", {"C6": "SO-0003", "C7": 4, "C8": "路線便"}); run("Pr_出荷ピッキング指示"); pick_all(newest_pk("組立")); check("S12_組立_出荷", "組立")

# ---- エラーになる操作 ----
def expect_err(macro, word, answers=None):
    global ng
    m = b.run(macro, answers)
    if word not in m: ng += 1; print("  エラーにならない／文言が違う:", macro, repr(m[:200]))
    else: print("  期待どおりエラー:", macro, "…" + word)
put("入力_製造", {"C6": "A-201", "C7": 100, "C8": 0}); run("Pr_製造指図作成")
put("入力_製造", {"C11": "MO-0003"})
m = b.run("Pr_製造ピッキング指示")
want = errs["E1_部材不足"].replace("部材が足りません：", "").strip().split("\n")
if all(w in m for w in want): print("  期待どおりエラー（HTML と同じ不足の内訳）:", " / ".join(want))
else: ng += 1; print("  部材不足の文言が違う:", m, want)
put("入力_製造", {"C17": tno("組立", "MO-0003", 10)}); expect_err("Pr_工程実績登録", "着手していません")
put("入力_製造", {"C6": "A-202", "C7": 1, "C8": 0}); run("Pr_製造指図作成")
mo = rows("製造指図", "組立")[-1][1]
put("入力_製造", {"C11": mo}); run("Pr_製造ピッキング指示"); pk = newest_pk("組立")
put("入力_ピッキング", {"C6": pk}); run("Pr_ピッキング読込"); expect_err("Pr_ピッキング確定", "まだ集めていない行が 4 行")
b.set_cell("入力_ピッキング", "B11", "○"); expect_err("Pr_ピッキング確定", "まだ集めていない行が 3 行")
saved = sum(1 for r in rows("ピッキング明細", "組立") if r[1] == pk and r[6] == "○")
if saved != 1: ng += 1; print("  済の印が保存されていない", saved)
else: print("  確定できなくても済の印は保存された")
put("入力_ピッキング", {"C6": "PK-0001"}); expect_err("Pr_ピッキング確定", "完了しています")
put("入力_ピッキング", {"C6": pk}); run("Pr_ピッキング読込"); run("Pr_ピッキング全部済"); run("Pr_ピッキング確定")
put("入力_製造", {"C17": tno("組立", mo, 10)}); run("Pr_手配No読込"); put("入力_製造", {"C18": 1, "C19": 1}); expect_err("Pr_工程実績登録", "投入数（1）と一致")
put("入力_製造", {"C17": tno("組立", mo, 20)}); expect_err("Pr_工程実績登録", "いま作業する工程ではありません")
put("入力_製造", {"C17": tno("加工", "MO-0001", 10)}); expect_err("Pr_手配No読込", "加工事業部 の工程")
put("入力_受注", {"C6": "A-202", "C7": 3, "C8": "北斗精工", "C9": 5}); run("Pr_受注登録")
so = rows("受注", "組立")[-1][1]; put("入力_受注", {"C12": so}); run("Pr_受注引当")
put("入力_出荷", {"C6": so, "C7": 5, "C8": "路線便"}); expect_err("Pr_出荷ピッキング指示", "3 までです")
put("入力_在庫調整", {"C6": "A-202", "C7": 1}); expect_err("Pr_在庫調整", "引当済（3）より少なくはできません")
put("入力_出荷", {"C6": so, "C7": ""}); run("Pr_出荷ピッキング指示")
q = [r for r in rows("ピッキング明細", "組立") if r[1] == newest_pk("組立")][0][4]
if q != 3: ng += 1; print("  出荷数の既定（残り全部）が違う", q)
else: print("  出荷数を空にすると残り全部（3）でピッキング指示が出た")
put("入力_受注", {"C12": so}); expect_err("Pr_受注引当解除", "ピッキング中のため解除できません")
b.run("Pr_在庫一覧更新")
b.close()
print(f"段階3：時点 {len(exp)} 件を比較、不一致 {ng} 件")
sys.exit(1 if ng else 0)
