from pathlib import Path
import json
import openpyxl
import re

SOURCES = [
    ("/Users/aadityabalakrishnan/Desktop/Product tank.xlsx", "Product managers", "ProductTank"),
    ("/Users/aadityabalakrishnan/Desktop/TPF.xlsx", "Product managers", "The Product Folks"),
    ("/Users/aadityabalakrishnan/Desktop/AIC.xlsx", "Founders", "The AI Collective"),
    ("/Users/aadityabalakrishnan/Desktop/Madrasters.xlsx", "AI practitioners", "Madrasters"),
]

records = []
for filename, category, source in SOURCES:
    sheet = openpyxl.load_workbook(filename, read_only=True, data_only=True).active
    for row in sheet.iter_rows(min_row=2, values_only=True):
        number, name, push_name = (list(row) + [None, None, None])[:3]
        if number is None:
            continue
        raw_number = str(number).strip()
        if not raw_number.isdigit():
            continue
        digits = raw_number
        supplied_display = str(name or push_name or '').strip()
        display = supplied_display
        display = re.sub(r'[\U0001F000-\U0001FAFF\U00002600-\U000027BF]+', '', display).strip()
        if supplied_display and not re.search(r'\w', display, re.UNICODE):
            continue
        if display == '~@v!π@$#~':
            continue
        if display == '-Prince Anto' and digits == '918939039381' and source == 'Madrasters':
            continue
        records.append({
            "name": display,
            "phone": "+" + digits,
            "category": category,
            "source": source,
        })

Path("network-data.js").write_text("window.NETWORK_CONTACTS = " + json.dumps(records, ensure_ascii=False, separators=(',', ':')) + ";\n", encoding="utf-8")
print(f"Created network-data.js with {len(records)} contact records.")
