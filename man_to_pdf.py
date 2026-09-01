import pdfplumber
import pandas as pd
from collections import defaultdict
from pathlib import Path
import re

# =========================
# 输入文件
# =========================

INPUT_FILE = "man.pdf"

# 根据输入 PDF 文件名自动生成输出 Excel 文件名
# 例如 sv.pdf -> sv_车辆配置.xlsx
pdf_path = Path(INPUT_FILE)
OUTPUT_FILE = pdf_path.with_name(pdf_path.stem + "_Bestellunterlagen.xlsx")

# =========================
# Code格式
# =========================

CODE_RE = re.compile(r'^(0P\w+|DE\w+|TD\d+|ZS\w+|ZLS\d+)$')

HEADER = [
    "MAN Truck",
    "Region West",
    "Bestellnummer",
    "Albert Ziegler",
    "Datenstand",
    "Seite von"
]

END_WORDS = [
    "Technische Daten",
    "Preiszusammenstellung"
]


def clean(txt):
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


rows = []
start = False
buffer = ""

with pdfplumber.open(INPUT_FILE) as pdf:

    # 第6页开始，pdf.pages索引从0开始
    for page in pdf.pages[0:]:

        words = page.extract_words(
            x_tolerance=2,
            y_tolerance=2,
            keep_blank_chars=False
        )

        # -----------------------
        # 按Y坐标组成一行
        # -----------------------

        lines = defaultdict(list)

        for w in words:
            y = round(w["top"], 1)
            lines[y].append(w)

        for y in sorted(lines):

            ws = sorted(lines[y], key=lambda x: x["x0"])

            # 去掉页眉
            txt = " ".join(i["text"] for i in ws)

            if any(h in txt for h in HEADER):
                continue

            if any(e in txt for e in END_WORDS):
                break

            # 找开始
            if "Fahrzeug-Charakteristika" in txt:
                start = True
                continue

            if not start:
                continue

            # -----------------------
            # 找Code
            # -----------------------

            code = None
            desc_words = []

            for w in ws:

                t = clean(w["text"])

                if CODE_RE.match(t):
                    code = t
                else:
                    desc_words.append(t)

            if code:

                desc = " ".join(desc_words)

                if buffer:
                    desc = buffer + " " + desc
                    buffer = ""

                rows.append([code, desc])

            else:

                # 长描述缓存
                if len(txt) > 5:
                    buffer += " " + txt

# -----------------------
# 输出Excel
# -----------------------

df = pd.DataFrame(rows, columns=["Code", "Description"])

df.to_excel(OUTPUT_FILE, index=False)

print(df.head())
print()
print("共识别:", len(df))
print("Excel已生成:", OUTPUT_FILE)