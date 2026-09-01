import pandas as pd
import pdfplumber
from pathlib import Path
from datetime import datetime
import re
import sys


CODE_FILE_KEYWORD = "Codes"
BESTELL_FILE_KEYWORD = "man"

FIVE_CODE_RE = re.compile(r"^[A-Z0-9]{5}$")


def clean_value(value):

    if pd.isna(value):
        return ""

    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def find_one_xlsx(keyword):

    files = sorted([
        p for p in Path(".").glob("*.xlsx")
        if keyword.lower() in p.name.lower()
        and not p.name.startswith("~$")
    ])

    if not files:
        raise FileNotFoundError(
            f"未找到名称包含 '{keyword}' 的xlsx文件"
        )

    return files[0]


def find_one_pdf(keyword):

    files = sorted([
        p for p in Path(".").glob("*.pdf")
        if keyword.lower() in p.name.lower()
    ])

    if not files:
        raise FileNotFoundError(
            f"未找到名称包含 '{keyword}' 的PDF文件"
        )

    return files[0]


def load_pdf_text(pdf_file):

    text = ""

    with pdfplumber.open(pdf_file) as pdf:

        for page in pdf.pages:

            try:
                page_text = page.extract_text()

                if page_text:
                    text += "\n" + page_text

            except:
                pass

    text = text.lower()

    text_clean = (
        text
        .replace(".", "")
        .replace(",", "")
        .replace(" ", "")
    )

    return text, text_clean


def check_pdf_text(name, pdf_text, pdf_text_clean):

    if not name:
        return False, "Name为空"

    name_lower = name.lower()

    # ==================================================
    # 1 Ohne ...
    # ==================================================

    if name_lower.startswith("ohne "):

        keyword = (
            name_lower
            .replace("ohne ", "")
            .strip()
        )

        keyword = keyword.split(",")[0]

        if keyword not in pdf_text:

            return (
                True,
                f"PDF中未发现 '{keyword}'"
            )

        return (
            False,
            f"PDF中发现 '{keyword}'"
        )

    # ==================================================
    # 2 寻找数字参数
    # ==================================================

    numbers = re.findall(
        r"\d+(?:[.,/x]\d+)*",
        name
    )

    for num in numbers:

        test_num = (
            num
            .replace(".", "")
            .replace(",", "")
            .replace(" ", "")
        )

        if test_num in pdf_text_clean:

            return (
                True,
                f"PDF中发现参数 {num}"
            )

    # ==================================================
    # 3 常见消防底盘参数
    # ==================================================

    keywords = re.findall(
        r"[A-Za-zÄÖÜäöüß]{5,}",
        name
    )

    blacklist = {
        "zulässige",
        "vorderachse",
        "hinterachse",
        "ausführung",
        "fahrerhaus",
        "rahmen",
        "normal",
        "seite",
        "ohne"
    }

    for word in keywords:

        word = word.lower()

        if word in blacklist:
            continue

        if word in pdf_text:

            return (
                True,
                f"PDF中发现关键词 '{word}'"
            )

    # ==================================================
    # 4 特殊检查
    # ==================================================

    if "4x2" in name_lower:

        if "4x2" in pdf_text:

            return (
                True,
                "PDF中发现4x2"
            )

    if "305/70r19" in name_lower:

        if "305/70" in pdf_text:

            return (
                True,
                "PDF中发现轮胎规格"
            )

    return (
        False,
        "PDF中未发现匹配内容"
    )


def main():

    codes_file = find_one_xlsx(CODE_FILE_KEYWORD)

    bestell_file = find_one_xlsx(
        BESTELL_FILE_KEYWORD
    )

    pdf_file = find_one_pdf(
        BESTELL_FILE_KEYWORD
    )

    print("Codes:", codes_file.name)
    print("Bestellunterlagen:", bestell_file.name)
    print("PDF:", pdf_file.name)

    df_codes = pd.read_excel(
        codes_file,
        dtype=str,
        engine="openpyxl"
    )

    df_bestell = pd.read_excel(
        bestell_file,
        header=None,
        dtype=str,
        engine="openpyxl"
    )

    pdf_text, pdf_text_clean = load_pdf_text(
        pdf_file
    )

    # ==================================================
    # Bestellunterlagen所有Code
    # ==================================================

    bestell_codes = set()

    for value in df_bestell.iloc[:, 0]:

        code = clean_value(value).upper()

        if FIVE_CODE_RE.match(code):

            bestell_codes.add(code)

    # ==================================================
    # 检查
    # ==================================================

    result_rows = []

    for idx, row in df_codes.iterrows():

        line_no = idx + 2

        bemerkung = clean_value(row.iloc[0])

        vcode = clean_value(row.iloc[1]).upper()

        name = ""

        if len(row) >= 3:
            name = clean_value(row.iloc[2])

        if not bemerkung or not vcode:
            continue

        if not FIVE_CODE_RE.match(vcode):

            result_rows.append({
                "Zeile": line_no,
                "Bemerkung": bemerkung,
                "VCode": vcode,
                "Name": name,
                "Status": "SKIP",
                "Detail": "非五位Code"
            })

            continue

        exists = vcode in bestell_codes

        # ==========================================
        # OK / hinzufügen
        # ==========================================

        if bemerkung in ["OK", "hinzufügen"]:

            if exists:

                status = "PASS_CODE"

                detail = (
                    "Code存在于Bestellunterlagen.xlsx"
                )

            else:

                pdf_match, pdf_detail = (
                    check_pdf_text(
                        name,
                        pdf_text,
                        pdf_text_clean
                    )
                )

                if pdf_match:

                    status = "PASS_TEXT"
                    detail = pdf_detail

                else:

                    status = "FAIL"
                    detail = (
                        "Code未找到且PDF未匹配"
                    )

        # ==========================================
        # Entfall
        # ==========================================

        elif bemerkung == "Entfall":

            if exists:

                status = "FAIL"

                detail = (
                    "Entfall Code仍存在"
                )

            else:

                status = "PASS"

                detail = (
                    "Entfall Code不存在"
                )

        else:

            status = "SKIP"

            detail = (
                f"未知Bemerkung={bemerkung}"
            )

        result_rows.append({
            "Codes文件行号": line_no,
            "Bemerkung von AZG": bemerkung,
            "VCode": vcode,
            "Name": name,
            "Code存在": "Ja" if exists else "Nein",
            "检查结果": status,
            "详细说明": detail
        })

    df_result = pd.DataFrame(result_rows)

    # ==================================================
    # Summary
    # ==================================================

    pass_code = int(
        (df_result["检查结果"] == "PASS_CODE").sum()
    )

    pass_text = int(
        (df_result["检查结果"] == "PASS_TEXT").sum()
    )

    pass_normal = int(
        (df_result["检查结果"] == "PASS").sum()
    )

    fail_count = int(
        (df_result["检查结果"] == "FAIL").sum()
    )

    skip_count = int(
        (df_result["检查结果"] == "SKIP").sum()
    )

    summary = pd.DataFrame([
        ["Codes文件", codes_file.name],
        ["Bestellunterlagen文件", bestell_file.name],
        ["PDF文件", pdf_file.name],
        ["PASS_CODE", pass_code],
        ["PASS_TEXT", pass_text],
        ["PASS", pass_normal],
        ["FAIL", fail_count],
        ["SKIP", skip_count],
        ["生成时间",
         datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
    ], columns=["项目", "值"])

    output_file = Path(
        f"man_check_output_{codes_file.stem}.xlsx"
    )

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        print(summary)

        df_result.to_excel(
            writer,
            sheet_name="Details",
            index=False
        )

        df_result[
            df_result["检查结果"] == "FAIL"
        ].to_excel(
            writer,
            sheet_name="FAIL_only",
            index=False
        )

    print()
    print("完成")
    print("输出文件:", output_file)


if __name__ == "__main__":

    try:

        main()

    except Exception as e:

        print("程序出错:")
        print(e)

        sys.exit(1)