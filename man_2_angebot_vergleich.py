import pdfplumber
import pandas as pd
from collections import defaultdict
from pathlib import Path
from datetime import datetime
import re
import sys


# ==================================================
# 手动指定两个 PDF
# ==================================================

OLD_PDF = "man old.pdf"
NEW_PDF = "man.pdf"

OUTPUT_PREFIX = "MAN_PDF_Bestellunterlagen_Vergleich"


# ==================================================
# Code格式
# ==================================================

CODE_RE = re.compile(
    r'^(0P\w+|DE\w+|TD\d+|ZS\w+|ZLS\d+|Z[A-Z0-9]{4,5})$'
)


HEADER = [
    "MAN Truck",
    "MAN Truck & Bus",
    "Region West",
    "Bestellnummer",
    "Albert Ziegler",
    "Datenstand",
    "Seite von",
    "Angebotsnummer"
]


END_WORDS = [
    "Technische Daten",
    "Preiszusammenstellung",
    "Fahrzeugpreis"
]


def clean(txt):
    if txt is None:
        return ""

    txt = str(txt).replace("\xa0", " ")
    txt = re.sub(r"\s+", " ", txt)

    return txt.strip()


def normalize_heading(txt):
    """
    用来判断真正章节标题。
    去掉星号、空格等。
    """
    txt = clean(txt)
    txt = txt.replace("*", "")
    txt = txt.strip()

    return txt


def is_real_start_line(txt):
    """
    只识别真正的章节标题：
    Fahrzeug-Charakteristika

    不识别目录里的：
    Fahrzeug-Charakteristika 4
    """

    heading = normalize_heading(txt)

    return heading == "Fahrzeug-Charakteristika"


def is_end_line(txt):
    """
    识别真正结束章节。
    """
    heading = normalize_heading(txt)

    for word in END_WORDS:
        if heading.startswith(word):
            return True

    return False


def is_header_line(txt):
    return any(h in txt for h in HEADER)


def extract_records_from_line(ws, buffer, page_no):
    """
    从一行中提取 Code 和 Description。

    兼容：
    1. Beschreibung ... Code
    2. Code Beschreibung ...
    3. 一行多个 Code
    """

    records = []

    tokens = []

    for w in ws:
        t = clean(w["text"])

        if t:
            tokens.append(t)

    if not tokens:
        return records, buffer

    code_positions = []

    for i, t in enumerate(tokens):
        if CODE_RE.match(t):
            code_positions.append(i)

    if not code_positions:
        return records, buffer

    # ==================================================
    # 一行只有一个 Code
    # ==================================================

    if len(code_positions) == 1:

        code_i = code_positions[0]
        code = tokens[code_i]

        # Code在行首
        if code_i == 0:
            desc_tokens = tokens[code_i + 1:]

        # Code在行尾或中间
        else:
            desc_tokens = tokens[:code_i]

        desc = clean(" ".join(desc_tokens))

        if buffer:
            desc = clean(buffer + " " + desc)
            buffer = ""

        records.append({
            "Code": code,
            "Description": desc,
            "Page": page_no
        })

        return records, buffer

    # ==================================================
    # 一行多个 Code
    # ==================================================

    for idx, code_i in enumerate(code_positions):

        code = tokens[code_i]

        next_code_i = (
            code_positions[idx + 1]
            if idx + 1 < len(code_positions)
            else len(tokens)
        )

        # 如果 Code 在行首或中间，描述取 Code 后到下一个 Code 前
        desc_tokens = tokens[code_i + 1:next_code_i]

        # 如果是最后一个 Code 且后面没描述，则取前面的文本
        if not desc_tokens and code_i > 0:
            desc_tokens = tokens[:code_i]

        desc = clean(" ".join(desc_tokens))

        if buffer:
            desc = clean(buffer + " " + desc)
            buffer = ""

        records.append({
            "Code": code,
            "Description": desc,
            "Page": page_no
        })

    return records, buffer


def extract_man_codes_from_pdf(input_file):
    """
    通用 MAN Bestellunterlagen PDF 提取逻辑。
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

                txt = clean(
                    " ".join(i["text"] for i in ws)
                )

                if not txt:
                    continue

                # 页眉跳过
                if is_header_line(txt):
                    continue

                # 真正开始章节
                if is_real_start_line(txt):
                    start = True
                    buffer = ""
                    continue

                if not start:
                    continue

                # 真正结束章节
                if is_end_line(txt):
                    start = False
                    break

                records, buffer = extract_records_from_line(
                    ws,
                    buffer,
                    page_no
                )

                if records:
                    rows.extend(records)

                else:
                    # 长描述缓存
                    if len(txt) > 5:
                        buffer = clean(buffer + " " + txt)

    df = pd.DataFrame(
        rows,
        columns=[
            "Code",
            "Description",
            "Page"
        ]
    )

    df = df.drop_duplicates().reset_index(drop=True)

    return df


def build_code_dict(df):
    """
    建立 Code -> Description/Page 字典。
    同一个 Code 多次出现时，只保留第一次。
    """

    result = {}

    for _, row in df.iterrows():

        code = clean(row.get("Code", ""))

        if not code:
            continue

        desc = clean(row.get("Description", ""))
        page = row.get("Page", "")

        if code not in result:
            result[code] = {
                "Description": desc,
                "Page": page
            }

    return result


def main():

    old_pdf = Path(OLD_PDF)
    new_pdf = Path(NEW_PDF)

    if not old_pdf.exists():
        raise FileNotFoundError(f"OLD PDF不存在: {old_pdf}")

    if not new_pdf.exists():
        raise FileNotFoundError(f"NEW PDF不存在: {new_pdf}")

    print("OLD PDF:", old_pdf.name)
    print("NEW PDF:", new_pdf.name)
    print()

    # ==================================================
    # 提取
    # ==================================================

    print("正在解析 OLD PDF ...")
    df_old_extract = extract_man_codes_from_pdf(old_pdf)

    print("正在解析 NEW PDF ...")
    df_new_extract = extract_man_codes_from_pdf(new_pdf)

    print()
    print("OLD 提取数量:", len(df_old_extract))
    print("NEW 提取数量:", len(df_new_extract))
    print()

    # ==================================================
    # 建立字典
    # ==================================================

    old_dict = build_code_dict(df_old_extract)
    new_dict = build_code_dict(df_new_extract)

    old_codes = set(old_dict.keys())
    new_codes = set(new_dict.keys())

    # ==================================================
    # 对比
    # ==================================================

    only_in_old = old_codes - new_codes
    only_in_new = new_codes - old_codes
    matched = old_codes & new_codes

    changed_description = []

    for code in matched:

        old_desc = old_dict.get(code, {}).get("Description", "")
        new_desc = new_dict.get(code, {}).get("Description", "")

        if old_desc != new_desc:
            changed_description.append(code)

    # ==================================================
    # Only in OLD
    # ==================================================

    only_old_rows = []

    for code in old_dict.keys():

        if code in only_in_old:
            only_old_rows.append({
                "Code": code,
                "Old_Description": old_dict[code]["Description"],
                "Old_Page": old_dict[code]["Page"]
            })

    df_only_old = pd.DataFrame(
        only_old_rows,
        columns=[
            "Code",
            "Old_Description",
            "Old_Page"
        ]
    )

    # ==================================================
    # Only in NEW
    # ==================================================

    only_new_rows = []

    for code in new_dict.keys():

        if code in only_in_new:
            only_new_rows.append({
                "Code": code,
                "New_Description": new_dict[code]["Description"],
                "New_Page": new_dict[code]["Page"]
            })

    df_only_new = pd.DataFrame(
        only_new_rows,
        columns=[
            "Code",
            "New_Description",
            "New_Page"
        ]
    )

    # ==================================================
    # Matched
    # ==================================================

    matched_rows = []

    for code in old_dict.keys():

        if code in matched:
            matched_rows.append({
                "Code": code,
                "Old_Description": old_dict[code]["Description"],
                "New_Description": new_dict[code]["Description"],
                "Old_Page": old_dict[code]["Page"],
                "New_Page": new_dict[code]["Page"]
            })

    df_matched = pd.DataFrame(
        matched_rows,
        columns=[
            "Code",
            "Old_Description",
            "New_Description",
            "Old_Page",
            "New_Page"
        ]
    )

    # ==================================================
    # Changed Description
    # ==================================================

    changed_rows = []

    for code in old_dict.keys():

        if code in changed_description:
            changed_rows.append({
                "Code": code,
                "Old_Description": old_dict[code]["Description"],
                "New_Description": new_dict[code]["Description"],
                "Old_Page": old_dict[code]["Page"],
                "New_Page": new_dict[code]["Page"]
            })

    df_changed = pd.DataFrame(
        changed_rows,
        columns=[
            "Code",
            "Old_Description",
            "New_Description",
            "Old_Page",
            "New_Page"
        ]
    )

    # ==================================================
    # Details
    # ==================================================

    detail_rows = []

    all_codes_ordered = []

    for code in old_dict.keys():
        if code not in all_codes_ordered:
            all_codes_ordered.append(code)

    for code in new_dict.keys():
        if code not in all_codes_ordered:
            all_codes_ordered.append(code)

    for code in all_codes_ordered:

        in_old = code in old_codes
        in_new = code in new_codes

        old_desc = old_dict.get(code, {}).get("Description", "")
        new_desc = new_dict.get(code, {}).get("Description", "")

        old_page = old_dict.get(code, {}).get("Page", "")
        new_page = new_dict.get(code, {}).get("Page", "")

        if in_old and in_new:

            if old_desc == new_desc:
                status = "MATCHED"
            else:
                status = "CHANGED_DESCRIPTION"

        elif in_old and not in_new:
            status = "REMOVED_FROM_NEW"

        elif in_new and not in_old:
            status = "ADDED_IN_NEW"

        else:
            status = "UNKNOWN"

        detail_rows.append({
            "Code": code,
            "Old_Description": old_desc,
            "New_Description": new_desc,
            "Old_Page": old_page,
            "New_Page": new_page,
            "In_OLD": "Ja" if in_old else "Nein",
            "In_NEW": "Ja" if in_new else "Nein",
            "Status": status
        })

    df_details = pd.DataFrame(
        detail_rows,
        columns=[
            "Code",
            "Old_Description",
            "New_Description",
            "Old_Page",
            "New_Page",
            "In_OLD",
            "In_NEW",
            "Status"
        ]
    )

    # ==================================================
    # Summary
    # ==================================================

    df_summary = pd.DataFrame([
        ["OLD PDF", old_pdf.name],
        ["NEW PDF", new_pdf.name],
        ["OLD 提取Code数量", len(df_old_extract)],
        ["NEW 提取Code数量", len(df_new_extract)],
        ["OLD 去重Code数量", len(old_codes)],
        ["NEW 去重Code数量", len(new_codes)],
        ["Matched数量", len(matched)],
        ["Only in OLD数量", len(only_in_old)],
        ["Only in NEW数量", len(only_in_new)],
        ["Changed Description数量", len(changed_description)],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ], columns=["项目", "值"])

    # ==================================================
    # 输出 Excel
    # ==================================================

    output_file = Path(
        f"{OUTPUT_PREFIX}_{old_pdf.stem}_vs_{new_pdf.stem}.xlsx"
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        df_only_old.to_excel(
            writer,
            sheet_name="Only_in_OLD",
            index=False
        )

        df_only_new.to_excel(
            writer,
            sheet_name="Only_in_NEW",
            index=False
        )

        df_matched.to_excel(
            writer,
            sheet_name="Matched",
            index=False
        )

        df_changed.to_excel(
            writer,
            sheet_name="Changed_Description",
            index=False
        )

        df_details.to_excel(
            writer,
            sheet_name="Details",
            index=False
        )

        df_old_extract.to_excel(
            writer,
            sheet_name="Extracted_OLD",
            index=False
        )

        df_new_extract.to_excel(
            writer,
            sheet_name="Extracted_NEW",
            index=False
        )

    print("比较完成")
    print("Excel已生成:", output_file)
    print()
    print("OLD PDF:", old_pdf.name)
    print("NEW PDF:", new_pdf.name)
    print()
    print("OLD 提取Code数量:", len(df_old_extract))
    print("NEW 提取Code数量:", len(df_new_extract))
    print("OLD 去重Code数量:", len(old_codes))
    print("NEW 去重Code数量:", len(new_codes))
    print("Matched:", len(matched))
    print("Only in OLD:", len(only_in_old))
    print("Only in NEW:", len(only_in_new))
    print("Changed Description:", len(changed_description))


if __name__ == "__main__":

    try:
        main()

    except Exception as e:
        print("程序出错:")
        print(e)
        sys.exit(1)