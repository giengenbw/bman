# BMAN Bedienungsanleitung / BMAN 使用说明

> [!CAUTION]
> **Unternehmensvertraulichkeit / 公司隐私提醒**  
> Verwenden Sie dieses Werkzeug nur in einer von Ihrem Unternehmen genehmigten Umgebung. Laden Sie keine echten PDF-, Excel- oder Ergebnisdateien auf GitHub, in öffentliche Cloud-Dienste, Messenger oder KI-Dienste hoch. Die Dateien können Bestell-, Fahrzeug-, Kunden-, Lieferanten-, Preis- oder Projektdaten enthalten.  
> 请仅在公司批准的环境中使用本工具。不要把真实的 PDF、Excel 或检查结果上传到 GitHub、公共网盘、聊天软件或 AI 服务。这些文件可能包含订单、车辆、客户、供应商、价格或项目数据。

Diese Anleitung richtet sich an Personen ohne Python-Erfahrung. Sie beschreibt die empfohlene Reihenfolge: zuerst `man_to_pdf.py` und `man_darius.py`, danach `man_check.py`.  
> 本说明面向没有 Python 使用经验的人员。推荐顺序：先运行 `man_to_pdf.py` 和 `man_darius.py`，然后运行 `man_check.py`。

---

## 1. Was macht das Programm? / 程序的作用

Der Ablauf besteht aus drei Skripten:  
整个流程由三个脚本组成：

1. **`man_to_pdf.py`**  
   Liest `man.pdf` und erzeugt `man_Bestellunterlagen.xlsx`.  
   读取 `man.pdf`，生成 `man_Bestellunterlagen.xlsx`。

2. **`man_darius.py`**  
   Liest `FG Prüfung.pdf` und erzeugt `FG Prüfung_Codes.xlsx`.  
   读取 `FG Prüfung.pdf`，生成 `FG Prüfung_Codes.xlsx`。

3. **`man_check.py`**  
   Vergleicht die Datei mit `Codes` im Namen mit der Datei mit `Bestellunterlagen` im Namen und erzeugt einen Prüfbericht.  
   比较文件名中包含 `Codes` 的 Excel 与文件名中包含 `Bestellunterlagen` 的 Excel，并生成检查报告。

### Prüflogik / 检查逻辑

- `OK` oder `hinzufügen`: Der fünfstellige Code muss in den Bestellunterlagen vorhanden sein.  
  `OK` 或 `hinzufügen`：五位 Code 必须存在于 Bestellunterlagen 中。
- `Entfall`: Der fünfstellige Code darf in den Bestellunterlagen nicht vorhanden sein.  
  `Entfall`：五位 Code 不应存在于 Bestellunterlagen 中。
- Andere oder nicht prüfbare Einträge werden als `SKIP` oder `WARNUNG` ausgegeben.  
  其他状态或无法检查的内容会显示为 `SKIP` 或 `WARNUNG`。

---

## 2. Voraussetzungen / 使用前准备

### Benötigte Dateien / 所需文件

Legen Sie diese Dateien alle in **denselben lokalen Ordner**:  
请把以下文件放在**同一个本地文件夹**中：

```text
BMAN/
├── man_to_pdf.py
├── man_darius.py
├── man_check.py
├── man.pdf
└── FG Prüfung.pdf
```

> [!IMPORTANT]
> Die beiden PDF-Dateinamen müssen genau `man.pdf` und `FG Prüfung.pdf` lauten. Andernfalls findet das jeweilige Skript die Datei nicht.  
> 两个 PDF 的文件名必须准确写成 `man.pdf` 和 `FG Prüfung.pdf`，否则脚本无法找到文件。

### Python installieren / 安装 Python

1. Lassen Sie Python nach Möglichkeit über die interne IT oder das genehmigte Softwareportal installieren.  
   建议通过公司 IT 或公司批准的软件中心安装 Python。
2. Unter Windows bei der Installation unbedingt **Add Python to PATH** aktivieren.  
   Windows 安装时请务必勾选 **Add Python to PATH**。
3. Öffnen Sie danach die **Eingabeaufforderung**: Windows-Taste drücken, `cmd` eingeben und Enter drücken.  
   安装后打开**命令提示符**：按 Windows 键，输入 `cmd`，然后按回车。
4. Prüfen Sie die Installation:  
   检查是否安装成功：

```bat
python --version
```

Falls dieser Befehl nicht funktioniert, versuchen Sie:  
如果该命令无效，请尝试：

```bat
py --version
```

---

## 3. Benötigte Pakete installieren / 安装所需组件

Öffnen Sie den BMAN-Ordner im Datei-Explorer. Klicken Sie oben in die Adressleiste, geben Sie `cmd` ein und drücken Sie Enter. Dadurch öffnet sich die Eingabeaufforderung direkt im richtigen Ordner.  
在文件资源管理器中打开 BMAN 文件夹。点击顶部地址栏，输入 `cmd` 并按回车，即可直接在当前文件夹打开命令提示符。

Installieren Sie danach einmalig die benötigten Pakete:  
然后只需安装一次以下组件：

```bat
python -m pip install pdfplumber pandas openpyxl
```

Falls Ihr Computer `py` statt `python` verwendet:  
如果电脑使用 `py` 而不是 `python`：

```bat
py -m pip install pdfplumber pandas openpyxl
```

> [!NOTE]
> In einer Unternehmensumgebung kann die Installation aus dem Internet gesperrt sein. Wenden Sie sich dann an die interne IT. Verwenden Sie keine nicht genehmigten Downloadquellen.  
> 公司环境可能禁止联网安装。如遇此情况，请联系内部 IT，不要使用未经批准的下载来源。

---

## 4. Ausführung in der richtigen Reihenfolge / 按正确顺序运行

### Schritt 1: Bestellunterlagen auslesen / 第一步：读取订单资料

Stellen Sie sicher, dass `man.pdf` im selben Ordner liegt. Führen Sie aus:  
确认 `man.pdf` 在同一文件夹中，然后运行：

```bat
python man_to_pdf.py
```

Erwartete Ausgabedatei:  
预期生成文件：

```text
man_Bestellunterlagen.xlsx
```

### Schritt 2: Darius-Prüf-PDF auslesen / 第二步：读取 Darius 检查 PDF

Stellen Sie sicher, dass `FG Prüfung.pdf` im selben Ordner liegt. Führen Sie aus:  
确认 `FG Prüfung.pdf` 在同一文件夹中，然后运行：

```bat
python man_darius.py
```

Erwartete Ausgabedatei:  
预期生成文件：

```text
FG Prüfung_Codes.xlsx
```

### Schritt 3: Dateien vergleichen / 第三步：比较文件

Schließen Sie die erzeugten Excel-Dateien vor der Prüfung. Führen Sie danach aus:  
检查前请关闭已生成的 Excel 文件，然后运行：

```bat
python man_check.py
```

Das Skript sucht automatisch im aktuellen Ordner nach:  
脚本会自动在当前文件夹查找：

- einer `.xlsx`-Datei mit `Codes` im Dateinamen;
- einer `.xlsx`-Datei mit `Bestellunterlagen` im Dateinamen.

Der Ergebnisname beginnt mit:  
结果文件名以以下内容开头：

```text
man_check_output_
```

Der Prüfbericht enthält normalerweise:  
检查报告通常包含：

- `Summary`: Zusammenfassung / 汇总
- `Details`: alle Prüfergebnisse / 全部检查结果
- `FAIL_only`: nur fehlerhafte Einträge / 仅错误项目

---

## 5. Empfohlener Komplettablauf / 推荐的一次性完整操作

Führen Sie die folgenden Befehle nacheinander aus:  
依次运行以下命令：

```bat
python man_to_pdf.py
python man_darius.py
python man_check.py
```

Wenn Ihr System `py` verwendet:  
如果系统使用 `py`：

```bat
py man_to_pdf.py
py man_darius.py
py man_check.py
```

Prüfen Sie anschließend zuerst das Blatt `Summary` und danach das Blatt `FAIL_only`.  
完成后先查看 `Summary` 工作表，再查看 `FAIL_only` 工作表。

---

## 6. Häufige Fehler / 常见问题

### `python` wurde nicht gefunden / 找不到 `python`

- Versuchen Sie `py` statt `python`.  
  尝试使用 `py` 代替 `python`。
- Starten Sie die Eingabeaufforderung nach der Python-Installation neu.  
  安装 Python 后重新打开命令提示符。
- Lassen Sie die PATH-Einstellung von der IT prüfen.  
  请 IT 检查 PATH 设置。

### `No module named ...` / 显示缺少模块

Installieren Sie die Pakete erneut:  
重新安装所需组件：

```bat
python -m pip install pdfplumber pandas openpyxl
```

### `FileNotFoundError` / 找不到文件

Prüfen Sie:  
请检查：

- Liegen Skript und PDF im selben Ordner? / 脚本和 PDF 是否在同一文件夹？
- Heißen die PDFs genau `man.pdf` und `FG Prüfung.pdf`? / PDF 是否准确命名？
- Ist die Dateiendung wirklich `.pdf`? / 文件扩展名是否确实为 `.pdf`？

### Falsche Excel-Datei wird geprüft / 检查了错误的 Excel 文件

`man_check.py` verwendet bei mehreren Treffern die alphabetisch zuerst gefundene passende Datei. Verschieben Sie alte oder nicht benötigte `Codes`- und `Bestellunterlagen`-Dateien vor dem Start in einen anderen Ordner.  
如果有多个匹配文件，`man_check.py` 会使用按名称排序后的第一个文件。运行前请把旧的或不需要的 `Codes`、`Bestellunterlagen` 文件移到其他文件夹。

### Excel-Ausgabedatei kann nicht geschrieben werden / 无法写入 Excel

Schließen Sie die betreffende Excel-Datei und starten Sie das Skript erneut.  
关闭相关 Excel 文件，然后重新运行脚本。

---

## 7. Datenschutz-Check vor und nach der Nutzung / 使用前后的隐私检查

Vor dem Start und vor jeder Weitergabe:  
运行前以及发送文件前，请检查：

- Nur genehmigte lokale oder interne Unternehmensspeicher verwenden.  
  仅使用公司批准的本地或内部存储位置。
- Keine Original-PDFs oder Ergebnisdateien in das Git-Repository committen.  
  不要把原始 PDF 或结果文件提交到 Git 仓库。
- Keine Dateien per privater E-Mail, privatem Cloud-Speicher oder Messenger versenden.  
  不要通过私人邮箱、个人网盘或聊天软件发送文件。
- Keine Screenshots mit Bestellnummern, Namen, Preisen oder Codes öffentlich teilen.  
  不要公开分享包含订单号、姓名、价格或 Code 的截图。
- Ergebnisdateien nur an berechtigte Personen weitergeben.  
  结果文件只发送给有权限的人员。
- Nicht mehr benötigte Arbeitskopien gemäß interner Löschrichtlinie entfernen.  
  按公司内部删除规定清理不再需要的工作副本。

### Empfohlene `.gitignore`-Einträge / 推荐的 `.gitignore` 设置

```gitignore
# Vertrauliche Eingabe- und Ergebnisdateien
*.pdf
*.xlsx
*.xls
*.csv
*.msg
*.eml

# Zugangsdaten und lokale Einstellungen
.env
.env.*
*.key
*.pem
credentials.*
secrets.*
token.*

# Temporäre Dateien
~$*
*.tmp
*.log
__pycache__/
```

> [!WARNING]
> Eine Datei bleibt möglicherweise in der Git-Historie enthalten, auch wenn sie später normal gelöscht wird. Wenn vertrauliche Daten versehentlich committed oder hochgeladen wurden, informieren Sie sofort die zuständige IT- oder Informationssicherheitsstelle.  
> 即使之后普通删除文件，该文件仍可能保留在 Git 历史中。如果误提交或上传了机密数据，请立即联系公司 IT 或信息安全部门。

---

## 8. Technischer Hinweis / 技术说明

Die drei Skripte wurden syntaktisch geprüft. Die tatsächliche Erkennungsqualität hängt jedoch vom Aufbau und von der Textqualität der jeweiligen PDF-Dateien ab. Prüfen Sie das erzeugte Excel-Ergebnis daher immer fachlich.  
三个脚本已通过 Python 语法检查，但实际识别质量仍取决于 PDF 的格式与文本质量。因此必须人工审核生成的 Excel 结果。

Dieses Werkzeug unterstützt die Prüfung, ersetzt aber keine fachliche Freigabe.  
本工具用于辅助检查，不能替代专业人员的最终审核与批准。
