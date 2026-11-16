# VBA の書き方の機械点検（依頼の「VBA の書き方」）
#  ① 変数・定数・引数の接頭辞が型と合っているか（英字2文字＋_）
#  ② Sub/Function 名が Pr_/Fn_ で始まるか（イベントプロシージャを除く）
#  ③ Sub の呼び出しに Call が付いているか
#  ④ メイン（On Error GoTo 終了処理 を持つ Public Sub）に書いてよいもの以外が無いか
#  ⑤ メインの呼び出し行すべてに行末コメントがあるか、Pr_変数格納 を呼んでいるか
import re, sys, pathlib
PFX = {"String": "St_", "Long": "Lg_", "Double": "Db_", "Boolean": "Bl_", "Date": "Dt_", "Variant": "Va_",
       "Object": "Ob_", "Collection": "Co_", "Worksheet": "Ws_", "Workbook": "Wb_", "Range": "Rg_"}
EVENTS = {"Class_Initialize", "Class_Terminate", "Workbook_Open"}
root = pathlib.Path(__file__).resolve().parent.parent / "src"
ng = []
def strip(line):
    q = False
    for j, c in enumerate(line):
        if c == '"': q = not q
        elif c == "'" and not q: return line[:j], line[j:]
    return line, ""
for p in sorted(root.glob("*.bas")):
    raw = p.read_text(encoding="utf-8").split("\n")
    # 行継続をつないだ論理行（元の行番号つき）
    lines, buf, start = [], "", 0
    for i, l in enumerate(raw, 1):
        body, cmt = strip(l)
        if not buf: start = i
        if body.rstrip().endswith(" _"):
            buf += body.rstrip()[:-1] + " "; continue
        lines.append((start, buf + body, cmt)); buf = ""
    proc, in_main, main_calls_store = None, False, False
    for no, body, cmt in lines:
        b = body.strip()
        m = re.match(r'(Public |Private )?(Sub|Function) (\S+?)\(', b)
        if m:
            proc = m.group(3)
            if proc not in EVENTS and not re.match(r'(Pr|Fn)_', proc):
                ng.append(f"{p.name}:{no}: ② 名前に Pr_/Fn_ がない {proc}")
            if m.group(2) == "Sub" and proc.startswith("Fn_") or m.group(2) == "Function" and proc.startswith("Pr_"):
                ng.append(f"{p.name}:{no}: ② Sub/Function と接頭辞が合わない {proc}")
            in_main = False; main_calls_store = False
            # 引数
            args = b[b.index('(') + 1: b.rindex(')')]
            for a in filter(None, [x.strip() for x in args.split(',')]):
                am = re.match(r'(Optional )?(ByVal |ByRef )?(\S+?)(\(\))? As (\w+)', a)
                if am and am.group(5) in PFX and not am.group(3).startswith(PFX[am.group(5)]):
                    ng.append(f"{p.name}:{no}: ① 引数の接頭辞 {am.group(3)} As {am.group(5)}")
            continue
        if re.match(r'End (Sub|Function)', b):
            if in_main and not main_calls_store: ng.append(f"{p.name}:{no}: ⑤ メイン {proc} が Pr_変数格納 を呼んでいない")
            proc = None; in_main = False; continue
        # 宣言
        dm = re.match(r'(Dim|Private|Public|Const|Private Const|Public Const) (?!Sub|Function|Enum|Type)(.*)', b)
        if dm and " As " in b:
            for d in dm.group(2).split(','):
                d = d.strip().replace("Const ", "")
                vm = re.match(r'(\S+?)(\(.*\))? As (\w+)', d)
                if vm:
                    n, t = vm.group(1), vm.group(3)
                    if not re.match(r'[A-Z][a-z]_', n): ng.append(f"{p.name}:{no}: ① 接頭辞が英字2文字＋_でない {n}")
                    elif t in PFX and not n.startswith(PFX[t]): ng.append(f"{p.name}:{no}: ① 型と接頭辞が合わない {n} As {t}")
        # Call の付け忘れ
        for stmt in b.split(": "):
            s = stmt.strip()
            s = re.sub(r'^(If .* Then |Else )', '', s)
            if re.match(r'Pr_\w+', s) and not re.match(r'Pr_\w+\s*=', s):
                ng.append(f"{p.name}:{no}: ③ Call がない {s[:40]}")
        if proc and b.startswith("On Error GoTo 終了処理"):
            in_main = True
        if in_main and b:
            ok = (re.match(r'Dim ', b) or b.startswith("On Error GoTo 終了処理") or re.match(r'Call (Pr|Fn)_', b)
                  or re.match(r'(Set )?\w+ = Fn_\w+\(.*\)$', b) or b == "終了処理:")
            if not ok: ng.append(f"{p.name}:{no}: ④ メイン {proc} に書けない行 {b[:40]}")
            if re.match(r'(Call |(Set )?\w+ = )(Pr|Fn)_', b) and not cmt.strip():
                ng.append(f"{p.name}:{no}: ⑤ メインの呼び出し行にコメントがない")
            if "Pr_変数格納" in b: main_calls_store = True
        if proc and not in_main and re.match(r'Call Pr_シート存在確認', b):
            pass
print("\n".join(ng) if ng else "点検OK（指摘なし）")
sys.exit(1 if ng else 0)
