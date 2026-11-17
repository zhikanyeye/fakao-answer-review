# -*- coding: utf-8 -*-
"""提取 docx 全部段落与表格文本，写入 utf-8 文件（随后用 Read 工具读取，切勿 print 中文）。

用法: python extract_docx.py <input.docx> <output.txt>
"""
import sys
import docx


def extract(path, out):
    d = docx.Document(path)
    lines = [p.text.strip() for p in d.paragraphs]
    for tbl in d.tables:
        for row in tbl.rows:
            lines.append(" | ".join(c.text.strip() for c in row.cells))
    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    # 仅输出 ASCII 状态信息，避免 Windows 控制台 GBK 报错
    print("OK paragraphs=%d tables=%d -> %s" % (len(d.paragraphs), len(d.tables), out))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: python extract_docx.py <input.docx> <output.txt>")
        sys.exit(1)
    extract(sys.argv[1], sys.argv[2])
