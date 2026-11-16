# LibreOffice（VBA 互換モード）で src/*.bas を動かす試験ハーネス。
#   ・Excel にしか無い部品を、試験のときだけ置き換える：
#       Scripting.Dictionary → Cs_辞書（.Item(k) の読み書きは Fn_取得／Pr_設定 に書き換え）
#       MsgBox → 記録するだけの Pr_試験表示／Fn_試験応答
#       フォームのボタン（Buttons.Add）→ 作らない
#   ・Excel 用のコード（src/、vba/）は書き換えない。置き換えはメモリ上だけ。
import uno, os, re, time, subprocess, json, pathlib
from com.sun.star.beans import PropertyValue
HERE = pathlib.Path(__file__).resolve().parent
SRC = HERE.parent / "src"
def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p

def _match_paren(s, i):
    # s[i] == '(' の対応する ')' の位置
    depth, q = 0, False
    for j in range(i, len(s)):
        c = s[j]
        if c == '"': q = not q
        elif not q and c == '(': depth += 1
        elif not q and c == ')':
            depth -= 1
            if depth == 0: return j
    raise ValueError("paren: " + s)

def _strip_comment(line):
    q = False
    for j, c in enumerate(line):
        if c == '"': q = not q
        elif c == "'" and not q: return line[:j].rstrip(), line[j:]
    return line, ""

def transform(code):
    code = code.replace("\r\n", "\n")
    code = re.sub(r'^Attribute VB_Name.*\n', '', code, flags=re.M)
    code = re.sub(r' _\n\s*', ' ', code)          # 行継続をつなぐ
    code = code.replace('vbObjectError + 1000', '-2147220504')   # LibreOffice は定数式の vbObjectError を受け付けない
    out = []
    for line in code.split("\n"):
        body, cmt = _strip_comment(line)
        body = body.replace('CreateObject("Scripting.Dictionary")', 'New Cs_辞書')
        # X.Item(k) = v → Call X.Pr_設定(k, v)
        m = re.match(r'^(\s*)([\w\.]+)\.Item\(', body)
        if m:
            o = body.index('(', m.end() - 1)
            c = _match_paren(body, o)
            rest = body[c + 1:]
            if rest.startswith(' = '):
                key = body[o + 1:c]
                val = rest[3:].replace('.Item(', '.Fn_取得(')
                body = f"{m.group(1)}Call {m.group(2)}.Pr_設定({key}, {val})"
        body = body.replace('.Item(', '.Fn_取得(')
        body = body.replace('MsgBox(', 'Fn_試験応答(')
        m = re.match(r'^(\s*)MsgBox (.*)$', body)
        if m: body = f"{m.group(1)}Call Pr_試験表示({m.group(2)})"
        m = re.match(r'^(\s*)(If .* Then )MsgBox (.*)$', body)
        if m: body = f"{m.group(1)}{m.group(2)}Call Pr_試験表示({m.group(3)})"
        if '.Tab.Color = ' in body:                 # シート見出しの色は LibreOffice に無い
            body = re.sub(r'\S+\.Tab\.Color = .*', "Rem タブ色は試験では省略", body)
        body = body.replace('hh:nn', 'hh:mm')         # LibreOffice は分を nn で書けない（Excel は nn が分）
        if '.Buttons.Add(' in body:
            body = re.sub(r'Set (\w+) = .*', r'Set \1 = Nothing: Exit Sub', body)
        out.append(body + ("" if not cmt else " " + cmt))
    return "Option VBASupport 1\n" + "\n".join(out)

def start_office():
    subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore",
                      "--accept=socket,host=localhost,port=2002;urp;"],
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def connect():
    local = uno.getComponentContext()
    res = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for k in range(80):
        try:
            return res.resolve("uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext")
        except Exception:
            if k == 0: start_office()
            time.sleep(0.5)
    raise SystemExit("LibreOffice に接続できません")

class Book:
    def __init__(self):
        ctx = connect()
        desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self.doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0,
                                             (pv("Hidden", False), pv("MacroExecutionMode", 4)))
        libs = self.doc.BasicLibraries
        if not libs.hasByName("Standard"): libs.createLibrary("Standard")
        lib = libs.getByName("Standard")
        mods = [("Cs_辞書", (HERE / "Cs_辞書.cls.txt").read_text(encoding="utf-8")),
                ("Mod_試験", (HERE / "Mod_試験.bas.txt").read_text(encoding="utf-8"))]
        for p in sorted(SRC.glob("*.bas")):
            mods.append((p.stem, transform(p.read_text(encoding="utf-8"))))
        for n, c in mods:
            if lib.hasByName(n): lib.removeByName(n)
            lib.insertByName(n, c)
        self.sp = self.doc.getScriptProvider()

    def call(self, proc, *args):
        mod = self.find_module(proc)
        s = self.sp.getScript(f"vnd.sun.star.script:Standard.{mod}.{proc}?language=Basic&location=document")
        return s.invoke(tuple(args), (), ())[0]

    def find_module(self, proc):
        lib = self.doc.BasicLibraries.getByName("Standard")
        for n in lib.getElementNames():
            if re.search(r'^\s*(Public |Private )?(Sub|Function) ' + re.escape(proc) + r'\b', lib.getByName(n), re.M):
                return n
        raise KeyError(proc)

    def run(self, proc, answers=None):
        """メインを1つ実行し、出たメッセージを返す"""
        self.call("Fn_試験記録取出")
        if answers is not None:
            self.set_answers(answers)
        try:
            self.call(proc)
        except Exception as e:
            return "[EXC]" + str(e)
        return self.call("Fn_試験記録取出")

    def set_answers(self, answers):
        code = "Sub Pr_試験応答設定()\n    Va_試験応答 = Array(" + ",".join(str(a) for a in answers) + ")\n    Lg_応答位置 = 0\nEnd Sub\n"
        lib = self.doc.BasicLibraries.getByName("Standard")
        if lib.hasByName("Mod_試験応答"): lib.removeByName("Mod_試験応答")
        lib.insertByName("Mod_試験応答", "Option VBASupport 1\n" + code)
        self.call("Pr_試験応答設定")

    def sheet_names(self):
        return list(self.doc.Sheets.ElementNames)

    def table(self, name, header_row=1):
        """シートの使用範囲を行のリストで返す（見出し行から）"""
        sh = self.doc.Sheets.getByName(name)
        cur = sh.createCursor(); cur.gotoEndOfUsedArea(False)
        last_r, last_c = cur.RangeAddress.EndRow, cur.RangeAddress.EndColumn
        rng = sh.getCellRangeByPosition(0, header_row - 1, last_c, last_r)
        return [list(r) for r in rng.DataArray]

    def cell(self, name, addr):
        c = self.doc.Sheets.getByName(name).getCellRangeByName(addr)
        return c.String if c.Type.value == "TEXT" else c.Value

    def set_cell(self, name, addr, value):
        c = self.doc.Sheets.getByName(name).getCellRangeByName(addr)
        if isinstance(value, str): c.String = value
        else: c.Value = value

    def close(self):
        self.doc.close(True)
