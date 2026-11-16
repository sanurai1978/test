# モジュールごとに LibreOffice でコンパイルが通るかを調べる（共通定義＋対象モジュールだけを読み込む）
import sys, os; sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lo_vba import pv, connect, transform, SRC, HERE
ctx = connect(); desk = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
names = sys.argv[1:] or [p.stem for p in sorted(SRC.glob("*.bas"))]
for n in names:
    doc = desk.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True), pv("MacroExecutionMode", 4)))
    libs = doc.BasicLibraries
    if not libs.hasByName("Standard"): libs.createLibrary("Standard")
    lib = libs.getByName("Standard")
    mods = {"Cs_辞書": (HERE / "Cs_辞書.cls.txt").read_text(encoding="utf-8"), "Mod_試験": (HERE / "Mod_試験.bas.txt").read_text(encoding="utf-8")}
    for m in {"Mod_共通定義", n}: mods[m] = transform((SRC / (m + ".bas")).read_text(encoding="utf-8"))
    mods[n] += "\nFunction Fn_生存() As String\n    Fn_生存 = \"ok\"\nEnd Function\n"
    for k, v in mods.items(): lib.insertByName(k, v)
    try: r = doc.getScriptProvider().getScript(f"vnd.sun.star.script:Standard.{n}.Fn_生存?language=Basic&location=document").invoke((), (), ())[0]
    except Exception as e: r = "EXC " + str(e)[:300]
    print(n, "OK" if r == "ok" else "NG " + repr(r)); doc.close(True)
