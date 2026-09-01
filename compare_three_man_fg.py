import pdfplumber
import pandas as pd
from collections import defaultdict
from pathlib import Path
from datetime import datetime
import re
import sys


# ==================================================
# 手动指定三个 PDF
# ==================================================

MAN_OLD_PDF = "man old.pdf"
MAN_NEW_PDF = "man.pdf"
FG_PDF = "FG Prüfung Goch.pdf"

OUTPUT_FILE = "Vergleich_MAN_old_MAN_new_FG_Pruefung.xlsx"


# ==================================================
# MAN Bestellunterlagen Code格式
# ==================================================

MAN_CODE_RE = re.compile(
    r'^(0P\w+|DE\w+|TD\d+|ZS\w+|ZLS\d+|Z[A-Z0-9]{4,5})$'
)


MAN_HEADER = [
    "MAN Truck",
    "MAN Truck & Bus",
    "Region West",
    "Bestellnummer",
    "Albert Ziegler",
    "Datenstand",
    "Seite von",
    "Angebotsnummer"
]

MAN_END_WORDS = [
    "Technische Daten",
    "Preiszusammenstellung"
]


# ==================================================
# FG Prüfung Code格式
# ==================================================

FG_VCODE_RE = re.compile(r"^([A-Z0-9]{5}|-)$")

FG_STATUS_WORDS = {
    "OK",
    "Entfall",
    "NatZu",
    "hinzufügen",
}

FG_SKIP_KEYWORDS = [
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

FG_CATEGORY_RE = re.compile(r"^x\*\*.*\*\*$")


# ==================================================
# 通用清理函数
# ==================================================

def clean(txt):
    if txt is None:
        return ""

    txt = str(txt).replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)

    return txt.strip()


# ==================================================
# MAN PDF 提取逻辑
# ==================================================

def extract_man_pdf(input_file, source_name):
    """
    从 MAN Bestellunterlagen PDF 中提取：
    Code | Description | Page | Source

    逻辑尽量贴近你当前可以工作的版本：
    - 从第一页开始扫描
    - 遇到 Fahrzeug-Charakteristika 后开始
    - 按坐标分行
    - 一行中识别 Code
    """

    rows = []
    start = False
    buffer = ""

    with pdfplumber.open(input_file) as pdf:

        for page_no, page in enumerate(pdf.pages, start=1):

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False
            )

            if not words:
                continue

            lines = defaultdict(list)

            for w in words:
                y = round(w["top"], 1)
                lines[y].append(w)

            for y in sorted(lines):

                ws = sorted(lines[y], key=lambda x: x["x0"])

                txt = clean(" ".join(i["text"] for i in ws))

                if not txt:
                    continue

                if any(h in txt for h in MAN_HEADER):
                    continue

                if any(e in txt for e in MAN_END_WORDS):
                    break

                if "Fahrzeug-Charakteristika" in txt:
                    start = True
                    continue

                if not start:
                    continue

                code = None
                desc_words = []

                for w in ws:

                    t = clean(w["text"])

                    if MAN_CODE_RE.match(t):
                        code = t
                    else:
                        desc_words.append(t)

                if code:

                    desc = clean(" ".join(desc_words))

                    if buffer:
                        desc = clean(buffer + " " + desc)
                        buffer = ""

                    rows.append({
                        "Code": code,
                        "Description": desc,
                        "Page": page_no,
                        "Source": source_name
                    })

                else:

                    if len(txt) > 5:
                        buffer = clean(buffer + " " + txt)

    df = pd.DataFrame(
        rows,
        columns=[
            "Code",
            "Description",
            "Page",
            "Source"
        ]
    )

    df = df.drop_duplicates().reset_index(drop=True)

    return df


# ==================================================
# FG Prüfung PDF 提取逻辑
# ==================================================

def is_fg_skip_line(txt):

    if not txt:
        return True

    if any(k in txt for k in FG_SKIP_KEYWORDS):
        return True

    if FG_CATEGORY_RE.match(txt):
        return True

    return False


def find_fg_records_in_line(words):
    """
    从 FG Prüfung 一行中识别：
    Bemerkung von AZG | VCode | Name
    """

    records = []

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

    code_positions = []

    for i, tok in enumerate(tokens):

        if FG_VCODE_RE.match(tok["text"]):
            code_positions.append(i)

    if not code_positions:
        return records

    for idx, code_i in enumerate(code_positions):

        code = tokens[code_i]["text"]

        next_code_i = (
            code_positions[idx + 1]
            if idx + 1 < len(code_positions)
            else len(tokens)
        )

        bemerkung = ""

        for j in range(code_i - 1, -1, -1):

            if tokens[j]["text"] in FG_STATUS_WORDS:
                bemerkung = tokens[j]["text"]
                break

        name_tokens = []

        for j in range(code_i + 1, next_code_i):

            t = tokens[j]["text"]

            if t in FG_STATUS_WORDS:
                break

            if t in {"x**", "Preis**"}:
                continue

            name_tokens.append(t)

        name = clean(" ".join(name_tokens))

        records.append({
            "Bemerkung": bemerkung,
            "Code": code,
            "Description": name
        })

    return records


def extract_fg_pdf(input_file, source_name):
    """
    从 FG Prüfung PDF 中提取：
    Code | Description | Bemerkung | Page | Source
    """

    rows = []
    last_row_index = None

    with pdfplumber.open(input_file) as pdf:

        for page_no, page in enumerate(pdf.pages, start=1):

            words = page.extract_words(
                x_tolerance=2,
                y_tolerance=2,
                keep_blank_chars=False
            )

            if not words:
                continue

            lines = defaultdict(list)

            for w in words:
                y = round(w["top"], 1)
                lines[y].append(w)

            for y in sorted(lines):

                line_words = sorted(lines[y], key=lambda w: w["x0"])

                line_txt = clean(
                    " ".join(w["text"] for w in line_words)
                )

                if is_fg_skip_line(line_txt):
                    continue

                records = find_fg_records_in_line(line_words)

                if records:

                    for r in records:

                        rows.append({
                            "Code": r["Code"],
                            "Description": r["Description"],
                            "Bemerkung": r["Bemerkung"],
                            "Page": page_no,
                            "Source": source_name
                        })

                        last_row_index = len(rows) - 1

                else:

                    if last_row_index is not None:

                        if len(line_txt) > 3 and not FG_CATEGORY_RE.match(line_txt):

                            rows[last_row_index]["Description"] = clean(
                                rows[last_row_index]["Description"] + " " + line_txt
                            )

    df = pd.DataFrame(
        rows,
        columns=[
            "Code",
            "Description",
            "Bemerkung",
            "Page",
            "Source"
        ]
    )

    df = df.drop_duplicates().reset_index(drop=True)

    return df


# ==================================================
# 字典构建
# ==================================================

def build_code_dict(df):
    """
    Code -> info
    同一个 Code 重复时保留第一次。
    """

    result = {}

    for _, row in df.iterrows():

        code = clean(row.get("Code", ""))

        if not code:
            continue

        if code == "-":
            continue

        if code not in result:

            result[code] = {
                "Description": clean(row.get("Description", "")),
                "Bemerkung": clean(row.get("Bemerkung", "")),
                "Page": row.get("Page", ""),
                "Source": clean(row.get("Source", ""))
            }

    return result


# ==================================================
# 生成 DataFrame
# ==================================================

def rows_for_codes(codes, source_dict, desc_col="Description", page_col="Page"):

    rows = []

    for code in source_dict.keys():

        if code in codes:

            rows.append({
                "Code": code,
                desc_col: source_dict[code].get("Description", ""),
                page_col: source_dict[code].get("Page", ""),
                "Bemerkung": source_dict[code].get("Bemerkung", "")
            })

    return pd.DataFrame(rows)


# ==================================================
# 主程序
# ==================================================

def main():

    man_old_pdf = Path(MAN_OLD_PDF)
    man_new_pdf = Path(MAN_NEW_PDF)
    fg_pdf = Path(FG_PDF)

    for p in [man_old_pdf, man_new_pdf, fg_pdf]:

        if not p.exists():
            raise FileNotFoundError(f"文件不存在: {p}")

    print("MAN OLD PDF:", man_old_pdf.name)
    print("MAN NEW PDF:", man_new_pdf.name)
    print("FG Prüfung PDF:", fg_pdf.name)
    print()

    # ==================================================
    # 提取三个文件
    # ==================================================

    print("正在解析 MAN OLD ...")
    df_man_old = extract_man_pdf(
        man_old_pdf,
        "MAN_OLD"
    )

    print("正在解析 MAN NEW ...")
    df_man_new = extract_man_pdf(
        man_new_pdf,
        "MAN_NEW"
    )

    print("正在解析 FG Prüfung ...")
    df_fg = extract_fg_pdf(
        fg_pdf,
        "FG_PRUEFUNG"
    )

    print()
    print("MAN OLD 提取:", len(df_man_old))
    print("MAN NEW 提取:", len(df_man_new))
    print("FG Prüfung 提取:", len(df_fg))
    print()

    # ==================================================
    # 建立字典和集合
    # ==================================================

    man_old_dict = build_code_dict(df_man_old)
    man_new_dict = build_code_dict(df_man_new)
    fg_dict = build_code_dict(df_fg)

    old_codes = set(man_old_dict.keys())
    new_codes = set(man_new_dict.keys())
    fg_codes = set(fg_dict.keys())

    # ==================================================
    # 三文件关系
    # ==================================================

    in_all_three = old_codes & new_codes & fg_codes

    only_old = old_codes - new_codes - fg_codes
    only_new = new_codes - old_codes - fg_codes
    only_fg = fg_codes - old_codes - new_codes

    old_new_only = (old_codes & new_codes) - fg_codes
    old_fg_only = (old_codes & fg_codes) - new_codes
    new_fg_only = (new_codes & fg_codes) - old_codes

    # ==================================================
    # 独立 Code 汇总
    # ==================================================

    unique_rows = []

    for code in man_old_dict.keys():

        if code in only_old:

            unique_rows.append({
                "Unique_In": "MAN_OLD_ONLY",
                "Code": code,
                "Description": man_old_dict[code]["Description"],
                "Bemerkung": "",
                "Page": man_old_dict[code]["Page"]
            })

    for code in man_new_dict.keys():

        if code in only_new:

            unique_rows.append({
                "Unique_In": "MAN_NEW_ONLY",
                "Code": code,
                "Description": man_new_dict[code]["Description"],
                "Bemerkung": "",
                "Page": man_new_dict[code]["Page"]
            })

    for code in fg_dict.keys():

        if code in only_fg:

            unique_rows.append({
                "Unique_In": "FG_ONLY",
                "Code": code,
                "Description": fg_dict[code]["Description"],
                "Bemerkung": fg_dict[code]["Bemerkung"],
                "Page": fg_dict[code]["Page"]
            })

    df_unique = pd.DataFrame(
        unique_rows,
        columns=[
            "Unique_In",
            "Code",
            "Description",
            "Bemerkung",
            "Page"
        ]
    )

    # ==================================================
    # Details 全量表
    # ==================================================

    all_codes_ordered = []

    for d in [man_old_dict, man_new_dict, fg_dict]:

        for code in d.keys():

            if code not in all_codes_ordered:
                all_codes_ordered.append(code)

    detail_rows = []

    for code in all_codes_ordered:

        in_old = code in old_codes
        in_new = code in new_codes
        in_fg = code in fg_codes

        count = sum([in_old, in_new, in_fg])

        if count == 3:
            status = "IN_ALL_THREE"
        elif count == 1:

            if in_old:
                status = "ONLY_IN_MAN_OLD"
            elif in_new:
                status = "ONLY_IN_MAN_NEW"
            else:
                status = "ONLY_IN_FG"

        elif count == 2:

            if in_old and in_new:
                status = "IN_MAN_OLD_AND_NEW"
            elif in_old and in_fg:
                status = "IN_MAN_OLD_AND_FG"
            else:
                status = "IN_MAN_NEW_AND_FG"

        else:
            status = "UNKNOWN"

        detail_rows.append({
            "Code": code,

            "In_MAN_OLD": "Ja" if in_old else "Nein",
            "MAN_OLD_Description": man_old_dict.get(code, {}).get("Description", ""),
            "MAN_OLD_Page": man_old_dict.get(code, {}).get("Page", ""),

            "In_MAN_NEW": "Ja" if in_new else "Nein",
            "MAN_NEW_Description": man_new_dict.get(code, {}).get("Description", ""),
            "MAN_NEW_Page": man_new_dict.get(code, {}).get("Page", ""),

            "In_FG": "Ja" if in_fg else "Nein",
            "FG_Bemerkung": fg_dict.get(code, {}).get("Bemerkung", ""),
            "FG_Description": fg_dict.get(code, {}).get("Description", ""),
            "FG_Page": fg_dict.get(code, {}).get("Page", ""),

            "Status": status
        })

    df_details = pd.DataFrame(detail_rows)

    # ==================================================
    # 各分类 Sheet
    # ==================================================

    df_only_old = rows_for_codes(
        only_old,
        man_old_dict,
        desc_col="MAN_OLD_Description",
        page_col="MAN_OLD_Page"
    )

    df_only_new = rows_for_codes(
        only_new,
        man_new_dict,
        desc_col="MAN_NEW_Description",
        page_col="MAN_NEW_Page"
    )

    df_only_fg = rows_for_codes(
        only_fg,
        fg_dict,
        desc_col="FG_Description",
        page_col="FG_Page"
    )

    df_all_three = df_details[
        df_details["Status"] == "IN_ALL_THREE"
    ].copy()

    df_old_new = df_details[
        df_details["Status"] == "IN_MAN_OLD_AND_NEW"
    ].copy()

    df_old_fg = df_details[
        df_details["Status"] == "IN_MAN_OLD_AND_FG"
    ].copy()

    df_new_fg = df_details[
        df_details["Status"] == "IN_MAN_NEW_AND_FG"
    ].copy()

    # ==================================================
    # Summary
    # ==================================================

    df_summary = pd.DataFrame([
        ["MAN OLD PDF", man_old_pdf.name],
        ["MAN NEW PDF", man_new_pdf.name],
        ["FG Prüfung PDF", fg_pdf.name],
        ["MAN OLD Code数量", len(old_codes)],
        ["MAN NEW Code数量", len(new_codes)],
        ["FG Prüfung Code数量", len(fg_codes)],
        ["三个文件都有", len(in_all_three)],
        ["只在 MAN OLD", len(only_old)],
        ["只在 MAN NEW", len(only_new)],
        ["只在 FG Prüfung", len(only_fg)],
        ["只在 MAN OLD + MAN NEW", len(old_new_only)],
        ["只在 MAN OLD + FG", len(old_fg_only)],
        ["只在 MAN NEW + FG", len(new_fg_only)],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ], columns=["项目", "值"])

    # ==================================================
    # 输出 Excel
    # ==================================================

    with pd.ExcelWriter(
        OUTPUT_FILE,
        engine="openpyxl"
    ) as writer:

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        df_unique.to_excel(
            writer,
            sheet_name="Unique_Codes",
            index=False
        )

        df_only_old.to_excel(
            writer,
            sheet_name="Only_in_MAN_OLD",
            index=False
        )

        df_only_new.to_excel(
            writer,
            sheet_name="Only_in_MAN_NEW",
            index=False
        )

        df_only_fg.to_excel(
            writer,
            sheet_name="Only_in_FG",
            index=False
        )

        df_all_three.to_excel(
            writer,
            sheet_name="In_All_Three",
            index=False
        )

        df_old_new.to_excel(
            writer,
            sheet_name="In_MAN_OLD_and_NEW",
            index=False
        )

        df_old_fg.to_excel(
            writer,
            sheet_name="In_MAN_OLD_and_FG",
            index=False
        )

        df_new_fg.to_excel(
            writer,
            sheet_name="In_MAN_NEW_and_FG",
            index=False
        )

        df_details.to_excel(
            writer,
            sheet_name="Details",
            index=False
        )

        df_man_old.to_excel(
            writer,
            sheet_name="Extracted_MAN_OLD",
            index=False
        )

        df_man_new.to_excel(
            writer,
            sheet_name="Extracted_MAN_NEW",
            index=False
        )

        df_fg.to_excel(
            writer,
            sheet_name="Extracted_FG",
            index=False
        )

    # ==================================================
    # 控制台输出
    # ==================================================

    print("比较完成")
    print("Excel已生成:", OUTPUT_FILE)
    print()
    print("MAN OLD Code数量:", len(old_codes))
    print("MAN NEW Code数量:", len(new_codes))
    print("FG Prüfung Code数量:", len(fg_codes))
    print()
    print("三个文件都有:", len(in_all_three))
    print("只在 MAN OLD:", len(only_old))
    print("只在 MAN NEW:", len(only_new))
    print("只在 FG Prüfung:", len(only_fg))
    print("只在 MAN OLD + MAN NEW:", len(old_new_only))
    print("只在 MAN OLD + FG:", len(old_fg_only))
    print("只在 MAN NEW + FG:", len(new_fg_only))


if __name__ == "__main__":

    try:
        main()

    except Exception as e:
        print("程序出错:")
        print(e)
        sys.exit(1)