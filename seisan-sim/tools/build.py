# src/*.bas（UTF-8・LF。編集用の原本）から vba/*.bas（Shift_JIS・CRLF。VBE に取り込む用）を作る。
# Shift_JIS にできない文字が混ざっていたら、ファイル名と行番号を出して止める。
import pathlib, sys
root = pathlib.Path(__file__).resolve().parent.parent
src, out = root / "src", root / "vba"
out.mkdir(exist_ok=True)
ng = False
for p in sorted(src.glob("*.bas")) + sorted(src.glob("*.cls")):
    text = p.read_text(encoding="utf-8").replace("\r\n", "\n")
    for no, line in enumerate(text.split("\n"), 1):
        try:
            line.encode("cp932")
        except UnicodeEncodeError as e:
            print(f"{p.name}:{no}: Shift_JIS にできない文字 {line[e.start]!r}")
            ng = True
    if not ng:
        (out / p.name).write_bytes(text.replace("\n", "\r\n").encode("cp932"))
sys.exit(1 if ng else 0)
