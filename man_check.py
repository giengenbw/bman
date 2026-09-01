import pandas as pd
from pathlib import Path
from datetime import datetime
import re
import sys

# ============================================================
# man_check.py
# 功能：
# 1. 自动查找当前目录下名称包含 "Codes" 的 xlsx 文件
# 2. 自动查找当前目录下名称包含 "Bestellunterlagen" 的 xlsx 文件
# 3. 对 Codes 文件逐行检查：
#    - A列 = OK 或 hinzufügen：B列 Code 必须存在于 Bestellunterlagen 文件 A列
#    - A列 = Entfall：B列 Code 必须不存在于 Bestellunterlagen 文件 A列
# 4. 输出详细检查结果 Excel
# ============================================================

CODE_FILE_KEYWORD = "Codes"
BESTELL_FILE_KEYWORD = "Bestellunterlagen"

# 只检查五位 Code；横杠 - 不作为五位 Code 检查
FIVE_CODE_RE = re.compile(r"^[A-Z0-9]{5}$")


def clean_value(value):
    """清理 Excel 单元格内容"""
    if pd.isna(value):
        return ""

    value = str(value).replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def find_one_xlsx(keyword):
    """
    在当前目录查找文件名包含 keyword 的 xlsx 文件
    如果找到多个，默认使用排序后的第一个
    """
    files = sorted([
        p for p in Path(".").glob("*.xlsx")
        if keyword.lower() in p.name.lower()
        and not p.name.startswith("~$")
    ])

    if len(files) == 0:
        raise FileNotFoundError(f"未找到名称包含 '{keyword}' 的 xlsx 文件")

    if len(files) > 1:
        print(f"发现多个名称包含 '{keyword}' 的 xlsx 文件，将使用第一个：")
        for p in files:
            print(" -", p.name)
        print()

    return files[0]


def main():

    # --------------------------------------------------------
    # 自动查找输入文件
    # --------------------------------------------------------

    codes_file = find_one_xlsx(CODE_FILE_KEYWORD)
    bestell_file = find_one_xlsx(BESTELL_FILE_KEYWORD)

    print("Codes 文件:", codes_file.name)
    print("Bestellunterlagen 文件:", bestell_file.name)
    print()

    # --------------------------------------------------------
    # 读取 Excel
    # --------------------------------------------------------
    # header=None 表示按 A/B/C 列位置读取，不依赖表头名称

    df_codes = pd.read_excel(
        codes_file,
        header=None,
        dtype=str,
        engine="openpyxl"
    )

    df_bestell = pd.read_excel(
        bestell_file,
        header=None,
        dtype=str,
        engine="openpyxl"
    )

    # 至少需要 A、B 两列
    if df_codes.shape[1] < 2:
        raise ValueError("Codes 文件至少需要 A列 Bemerkung 和 B列 VCode")

    if df_bestell.shape[1] < 1:
        raise ValueError("Bestellunterlagen 文件至少需要 A列 Code")

    # --------------------------------------------------------
    # 提取 Bestellunterlagen A列中的 Code 集合
    # --------------------------------------------------------

    bestell_codes = set()

    for value in df_bestell.iloc[:, 0].tolist():
        code = clean_value(value).upper()

        if FIVE_CODE_RE.match(code):
            bestell_codes.add(code)

    # --------------------------------------------------------
    # 逐行检查 Codes 文件
    # --------------------------------------------------------

    result_rows = []

    for excel_row_no, row in df_codes.iterrows():

        # Excel 实际行号 = DataFrame index + 1
        line_no = excel_row_no + 1

        bemerkung = clean_value(row.iloc[0])
        vcode = clean_value(row.iloc[1]).upper()

        name = ""
        if df_codes.shape[1] >= 3:
            name = clean_value(row.iloc[2])

        # 如果 A列和B列都为空，跳过
        if not bemerkung and not vcode:
            continue

        # 只处理 A列B列都不为空的行
        if not bemerkung or not vcode:
            result_rows.append({
                "Codes文件行号": line_no,
                "Bemerkung von AZG": bemerkung,
                "VCode": vcode,
                "Name": name,
                "检查规则": "A列或B列为空，无法检查",
                "Bestellunterlagen中是否存在": "未检查",
                "检查结果": "WARNUNG",
                "详细说明": "该行不是 A列B列都不为空，已标记为警告"
            })
            continue

        # 跳过表头行
        if bemerkung.lower() in {"bemerkung von azg", "bemerkung"}:
            continue

        if vcode.lower() in {"vcode", "code"}:
            continue

        # 只检查五位 Code；例如 '-' 不检查，但保留在 output 中
        if not FIVE_CODE_RE.match(vcode):
            result_rows.append({
                "Codes文件行号": line_no,
                "Bemerkung von AZG": bemerkung,
                "VCode": vcode,
                "Name": name,
                "检查规则": "非五位Code，不参与检查",
                "Bestellunterlagen中是否存在": "未检查",
                "检查结果": "SKIP",
                "详细说明": "VCode 不是五位数字/字母，例如 '-'，未参与存在性检查"
            })
            continue

        exists = vcode in bestell_codes
        exists_text = "是" if exists else "否"

        # ----------------------------------------------------
        # 规则判断
        # ----------------------------------------------------

        if bemerkung in {"OK", "hinzufügen"}:

            rule = "OK/hinzufügen：Bestellunterlagen A列中必须存在该Code"
            passed = exists

            if passed:
                detail = "符合：该Code存在"
            else:
                detail = "不符合：该Code应该存在，但未找到"

        elif bemerkung == "Entfall":

            rule = "Entfall：Bestellunterlagen A列中必须不存在该Code"
            passed = not exists

            if passed:
                detail = "符合：该Code不存在"
            else:
                detail = "不符合：该Code为Entfall，但仍在Bestellunterlagen中找到"

        else:

            rule = "未知Bemerkung，不参与OK/Entfall规则判断"
            passed = None
            detail = f"Bemerkung 为 '{bemerkung}'，不在 OK / hinzufügen / Entfall 范围内"

        if passed is True:
            check_result = "PASS"
        elif passed is False:
            check_result = "FAIL"
        else:
            check_result = "SKIP"

        result_rows.append({
            "Codes文件行号": line_no,
            "Bemerkung von AZG": bemerkung,
            "VCode": vcode,
            "Name": name,
            "检查规则": rule,
            "Bestellunterlagen中是否存在": exists_text,
            "检查结果": check_result,
            "详细说明": detail
        })

    df_result = pd.DataFrame(result_rows)

    # --------------------------------------------------------
    # 汇总表
    # --------------------------------------------------------

    if len(df_result) > 0:
        pass_count = int((df_result["检查结果"] == "PASS").sum())
        fail_count = int((df_result["检查结果"] == "FAIL").sum())
        skip_count = int((df_result["检查结果"] == "SKIP").sum())
        warn_count = int((df_result["检查结果"] == "WARNUNG").sum())
    else:
        pass_count = 0
        fail_count = 0
        skip_count = 0
        warn_count = 0

    total = len(df_result)

    df_summary = pd.DataFrame([
        ["Codes文件", codes_file.name],
        ["Bestellunterlagen文件", bestell_file.name],
        ["Bestellunterlagen A列五位Code数量", len(bestell_codes)],
        ["检查总行数", total],
        ["PASS", pass_count],
        ["FAIL", fail_count],
        ["SKIP", skip_count],
        ["WARNUNG", warn_count],
        ["生成时间", datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
    ], columns=["项目", "值"])

    # --------------------------------------------------------
    # 输出文件名
    # --------------------------------------------------------

    output_file = Path(
        f"man_check_output_{codes_file.stem}_vs_{bestell_file.stem}.xlsx"
    )

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        df_summary.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        df_result.to_excel(
            writer,
            sheet_name="Details",
            index=False
        )

        # FAIL 单独输出一张表，方便查看
        df_fail = df_result[df_result["检查结果"] == "FAIL"].copy()

        df_fail.to_excel(
            writer,
            sheet_name="FAIL_only",
            index=False
        )

    print("检查完成")
    print("输出文件:", output_file.name)
    print()
    print("Summary:")
    print("PASS:", pass_count)
    print("FAIL:", fail_count)
    print("SKIP:", skip_count)
    print("WARNUNG:", warn_count)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("程序出错:", e)
        sys.exit(1)