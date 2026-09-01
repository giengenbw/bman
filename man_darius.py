import pdfplumber
import pandas as pd
from collections import defaultdict
import re
from pathlib import Path

# =========================
# 基本配置
# =========================

PDF_FILE = "FG Prüfung.pdf"

pdf_path = Path(PDF_FILE)
OUT_FILE = pdf_path.with_name(pdf_path.stem + "_Codes.xlsx")


# VCode：5位数字/字母
# 例如：0P0H7, 0P2UD, DE06K, ZFERA, ZW010, ZCAVL
VCODE_RE = re.compile(r"^([A-Z0-9]{5}|-)$")

# Bemerkung von AZG 可能出现的值
STATUS_WORDS = {
    "OK",
    "Entfall",
    "NatZu",
    "hinzufügen",
}

# 需要跳过的页眉/页脚/标题
SKIP_KEYWORDS = [
    "Kunde / TYP:",
    "MAN-Code",
    "Datenabgriff",
    "MAN TRUCK",
    "Bemerkung von AZG",
    "Info**",
    "VCode",
    "Name",
    "Preis**",
    "Albert Ziegler GmbH",
    "Seite"
]

# 分类标题，例如 x** Basis-Charakteristika**
CATEGORY_RE = re.compile(r"^x\*\*.*\*\*$")


def clean(txt):
    if txt is None:
        return ""
    txt = txt.replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


def is_skip_line(txt):
    """判断是否跳过该行"""
    if not txt:
        return True

    if any(k in txt for k in SKIP_KEYWORDS):
        return True

    if CATEGORY_RE.match(txt):
        return True

    return False


def find_records_in_line(words):
    """
    从一行 words 中按坐标识别记录：
    Bemerkung von AZG | VCode | Name
    """
    records = []

    # 按 x 坐标排序
    words = sorted(words, key=lambda w: w["x0"])

    tokens = []
    for w in words:
        t = clean(w["text"])
        if t:
            tokens.append({
                "text": t,
                "x0": w["x0"],
                "x1": w["x1"],
                "top": w["top"]
            })

    if not tokens:
        return records

    # 找这一行里的所有 VCode
    code_positions = []
    for i, tok in enumerate(tokens):
        if VCODE_RE.match(tok["text"]):
            code_positions.append(i)

    if not code_positions:
        return records

    for idx, code_i in enumerate(code_positions):

        code = tokens[code_i]["text"]

        # 下一个 code 的位置，用于截断当前 name
        next_code_i = code_positions[idx + 1] if idx + 1 < len(code_positions) else len(tokens)

        # -------------------------
        # 找 Bemerkung
        # -------------------------
        bemerkung = ""

        # 优先找 code 左边最近的状态词
        for j in range(code_i - 1, -1, -1):
            if tokens[j]["text"] in STATUS_WORDS:
                bemerkung = tokens[j]["text"]
                break

        # -------------------------
        # 找 Name
        # -------------------------
        name_tokens = []

        for j in range(code_i + 1, next_code_i):
            t = tokens[j]["text"]

            # 如果遇到状态词，说明可能进入下一条记录
            if t in STATUS_WORDS:
                break

            # 跳过价格/数量标记
            if t in {"x**", "Preis**"}:
                continue

            name_tokens.append(t)

        name = clean(" ".join(name_tokens))

        # 有 code 就保留
        records.append({
            "Bemerkung von AZG": bemerkung,
            "VCode": code,
            "Name": name
        })

    return records


rows = []
last_row_index = None

with pdfplumber.open(PDF_FILE) as pdf:

    for page_no, page in enumerate(pdf.pages, start=1):

        words = page.extract_words(
            x_tolerance=2,
            y_tolerance=2,
            keep_blank_chars=False
        )

        if not words:
            continue

        # =========================
        # 按 Y 坐标分行
        # =========================
        lines = defaultdict(list)

        for w in words:
            y = round(w["top"], 1)
            lines[y].append(w)

        for y in sorted(lines):

            line_words = sorted(lines[y], key=lambda w: w["x0"])
            line_txt = clean(" ".join(w["text"] for w in line_words))

            if is_skip_line(line_txt):
                continue

            # 尝试从这一行识别记录
            records = find_records_in_line(line_words)

            if records:
                for r in records:
                    rows.append(r)
                    last_row_index = len(rows) - 1

            else:
                # 如果没有新的 code，但这一行可能是上一条 Name 的换行描述
                # 例如长 Beschreibung 被 PDF 自动换行
                if last_row_index is not None:
                    # 排除明显不是描述的内容
                    if len(line_txt) > 3 and not CATEGORY_RE.match(line_txt):
                        rows[last_row_index]["Name"] = clean(
                            rows[last_row_index]["Name"] + " " + line_txt
                        )


# =========================
# 去重和输出
# =========================

df = pd.DataFrame(
    rows,
    columns=[
        "Bemerkung von AZG",
        "VCode",
        "Name"
    ]
)

# 去掉完全重复行
df = df.drop_duplicates().reset_index(drop=True)

# 输出 Excel
df.to_excel(OUT_FILE, index=False)

# 控制台输出所有 code
print(df)

print()
print("共识别 Code 数量:", len(df))

print()
print("所有 VCode:")
for code in df["VCode"].tolist():
    print(code)