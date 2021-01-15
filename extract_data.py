import csv
import json
import sys
from itertools import groupby
from docx import Document

doc = Document(sys.argv[1])

rows = []

for table in doc.tables:
    if table.rows[0].cells[0].text not in ["Field Number", "147"]:
        continue

    for table_row in table.rows:
        if not table_row.cells[0].text:
            continue
        if table_row.cells[0].text == "Field Number":
            continue
        if table_row.cells[0].text.startswith("Note"):
            continue

        row = [
            " ".join(
                "".join(e.text for e in para._element.xpath(".//w:r"))
                for para in cell.paragraphs
            )
            .strip()
            .replace("–", "-")
            for cell in table_row.cells
        ]

        row = [item for item, _ in groupby(row)]
        rows.append(row)

# with open("tmp.csv", "w") as f:
#     writer = csv.writer(f)
#     writer.writerows(rows)

# with open("v1.1.csv") as f:
#     rows = list(csv.reader(f))

records = []

last_id = 0

for id, group in groupby(rows, lambda row: row[0]):
    # print(id)
    id = int(id)
    assert id == last_id + 1
    last_id = id

    group = list(group)
    name = group[0][1].strip().upper()
    if not (name.endswith("_COD") or name in ["BMI_STAGE", "SEV_OBESITY"]):
        continue

    for row in group:
        assert row[1].strip().upper() == name

    record = {
        "id": id,
        "name": name,
        "details": group[0][2],
        "criteria": group[0][3],
    }

    if len(group) == 1:
        assert name == "BMI_COD"
        title = "BMI"
    else:
        title = group[1][2]
        if title[0] == "(":
            assert title[-1] == ")"
            title = title[1:-1]
    record["title"] = title

    records.append(record)


with open(sys.argv[2], "w") as f:
    json.dump(records, f, indent=2)
