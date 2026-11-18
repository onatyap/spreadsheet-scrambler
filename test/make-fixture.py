"""Builds test/fixture.xlsx: a small workbook of invented data with the features the scrambler has to handle.
openpyxl writes every text inline, so if LibreOffice is installed the same workbook is also saved through it as
test/fixture-sst.xlsx, which keeps text in a shared strings list and charts with cached values, the way Excel does.

    python3 test/make-fixture.py      (needs openpyxl; LibreOffice for the second file)
"""
import datetime as dt
import os
import re
import shutil
import subprocess
import tempfile
import zipfile

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import Rule
from openpyxl.packaging.custom import StringProperty
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Sales"
ws.append(["Customer", "Region", "Revenue", "Share", "Signed", "Active", "Status", "Double", "Contact"])
rows = [
    ("Acme Trading LLC", "North", 1234567, 0.153, dt.datetime(2026, 3, 1), True, "Yes"),
    ("Globex Holdings", "South", 98765, 0.0042, dt.datetime(2025, 12, 31, 14, 30), False, "Overdue"),
    ("شركة الاتصالات", "North", -450, 1.05, dt.datetime(2026, 7, 15), True, "No"),
    ("Acme Trading LLC", "East", 0, 0.5, dt.datetime(2024, 2, 29), False, "Yes"),
]
for r, row in enumerate(rows, start=2):
    ws.append(row)
    ws[f"H{r}"] = f'=IF(G{r}="Yes",C{r}*2,0)'
ws["I2"] = CellRichText([TextBlock(InlineFont(b=True), "Bilal "), "Haddad"])
ws["I3"] = "#N/A"
ws["I3"].data_type = "e"
ws["C7"] = "=SUM(C2:C5)"
ws["C8"] = '=COUNTIF(B2:B5,"North")'
ws["C9"] = '=VLOOKUP(A2,Lookup!A2:B4,2,FALSE)'
ws["A2"].comment = Comment("Very secret note about this customer", "Jane Example")
ws["A3"].hyperlink = "https://intranet.example.org/customers/globex"
for r in range(2, 6):
    ws[f"E{r}"].number_format = "yyyy-mm-dd hh:mm"
    ws[f"D{r}"].number_format = "0.0%"
dv = DataValidation(type="list", formula1='"Yes,No,Overdue,Disputed"')
ws.add_data_validation(dv)
dv.add("G2:G5")
ws.conditional_formatting.add("G2:G5", Rule(type="containsText", operator="containsText", text="Overdue",
                                            formula=['NOT(ISERROR(SEARCH("Overdue",G2)))'],
                                            dxf=DifferentialStyle(fill=PatternFill(bgColor="FFC7CE"))))
ws.auto_filter.ref = "A1:I5"
ws.auto_filter.add_filter_column(1, ["North"])
wb.defined_names["Flagship"] = DefinedName("Flagship", attr_text='"Acme Trading LLC"')
chart = BarChart()
chart.add_data(Reference(ws, min_col=3, min_row=1, max_row=5), titles_from_data=True)
ws.add_chart(chart, "K2")

lookup = wb.create_sheet("Lookup")
lookup.append(["Name", "Code"])
for name, code in [("Acme Trading LLC", 4711), ("Globex Holdings", 815), ("Initech Partners", 1999)]:
    lookup.append([name, code])
lookup.add_table(Table(displayName="Codes", ref="A1:B4"))

notes = wb.create_sheet("Notes")
notes.sheet_state = "hidden"
notes["A1"] = "Board approved the Initech Partners deal"
notes["A2"] = 42

wb.properties.creator = "Jane Example"
wb.properties.title = "Quarterly customers"
wb.custom_doc_props.append(StringProperty(name="Client", value="Acme Trading LLC"))
here = os.path.dirname(os.path.abspath(__file__))
wb.save(os.path.join(here, "fixture.xlsx"))

soffice = shutil.which("soffice") or "/Applications/LibreOffice.app/Contents/MacOS/soffice"
if os.path.exists(soffice):
    with tempfile.TemporaryDirectory() as work:
        subprocess.run([soffice, f"-env:UserInstallation=file://{work}/profile", "--headless", "--convert-to",
                        "xlsx:Calc MS Excel 2007 XML", "--outdir", work, os.path.join(here, "fixture.xlsx")],
                       check=True, capture_output=True, timeout=300)
        # LibreOffice writes the machine's locale into the file; replace it so the fixture says nothing about who built it
        with zipfile.ZipFile(os.path.join(work, "fixture.xlsx")) as src, \
                zipfile.ZipFile(os.path.join(here, "fixture-sst.xlsx"), "w", zipfile.ZIP_DEFLATED) as out:
            for item in src.infolist():
                data = src.read(item.filename)
                if item.filename == "docProps/core.xml":
                    data = re.sub(rb"<dc:language>[^<]*</dc:language>", b"<dc:language>en-US</dc:language>", data)
                out.writestr(item, data)
else:
    print("LibreOffice not found, so fixture-sst.xlsx was not rebuilt")
