from __future__ import annotations

import json
import pathlib
import re

import openpyxl


ROOT = pathlib.Path(__file__).resolve().parents[1]
HTML_PATH = ROOT / "index.html"
DATA_PATH = ROOT / "map-data.js"
WORKBOOK_PATH = ROOT / "Skipton - CGI partners travel.xlsx"
SHEET_NAME = "Final View"

EXTRA_CENTROIDS = {
    "NG1": [52.95478264984716, -1.1484459204892967],
    "YO42": [53.92394883152171, -0.7900222105978265],
}

FULL_POSTCODE_LOCATIONS = {
    "LS25 6PQ": [53.795993, -1.23728],
    "YO42 2LY": [53.932113, -0.773118],
    "PR1 9HS": [53.739046, -2.72244],
    "HA2 9QL": [51.569997, -0.381045],
}


def clean_text(value):
    if value is None:
        return None
    text = str(value).replace("\xa0", " ").strip()
    return text or None


def clean_postcode(value):
    text = clean_text(value)
    if not text or text.upper() in {"N/A", "NA", "NONE"}:
        return None
    return re.sub(r"\s+", " ", text.upper())


def is_full_postcode(postcode):
    return bool(postcode and re.match(r"^[A-Z]{1,2}\d[A-Z\d]?\s?\d[A-Z]{2}$", postcode))


def outward_code(postcode):
    if not postcode:
        return None

    compact = re.sub(r"\s+", "", postcode.upper())
    match = re.match(r"^([A-Z]{1,2}\d[A-Z\d]?)(\d[A-Z]{2})$", compact)
    return match.group(1) if match else compact


def clean_number(value):
    if value is None:
        return None

    number = float(value)
    return int(number) if number.is_integer() else number


def spreadsheet_rows():
    workbook = openpyxl.load_workbook(WORKBOOK_PATH, data_only=True)
    sheet = workbook[SHEET_NAME]
    headers = [cell.value for cell in sheet[1]]
    column = {header: index for index, header in enumerate(headers)}

    rows = []
    for raw_row in sheet.iter_rows(min_row=2, values_only=True):
        name = clean_text(raw_row[column["Name"]])
        if not name:
            continue

        postcode = clean_postcode(raw_row[column["Postcode"]])
        rows.append(
            {
                "name": name,
                "postcode": postcode,
                "outward": outward_code(postcode),
                "is_full": is_full_postcode(postcode),
                "role": clean_text(raw_row[column["Role"]]),
                "commute": clean_number(raw_row[column["Commute (hrs)"]]),
                "owner": clean_text(raw_row[column["Owner"]]),
                "expectation": clean_text(raw_row[column["Expectation based on location"]]),
                "spoken_to": clean_text(raw_row[column["Spoken to"]]),
                "outcome": clean_text(raw_row[column["Outcome"]]),
            }
        )

    return rows


def current_centroids(html):
    if DATA_PATH.exists():
        data_text = DATA_PATH.read_text()
        match = re.search(
            r'"districtCentroids": (.*?),\n  "fullPostcodeLocations"',
            data_text,
            flags=re.S,
        )
        if match:
            centroids = json.loads(match.group(1))
            centroids.update(EXTRA_CENTROIDS)
            return centroids

    match = re.search(
        r"const districtCentroids = (.*?);\nconst (?:fullPostcodeLocations|skipton(?:Address)?)\b",
        html,
        flags=re.S,
    )
    if not match:
        raise RuntimeError("Could not find districtCentroids block in map data or index.html")

    centroids = json.loads(match.group(1))
    centroids.update(EXTRA_CENTROIDS)
    return centroids


def main():
    html = HTML_PATH.read_text()
    rows = spreadsheet_rows()
    centroids = current_centroids(html)

    payload = {
        "colleagues": rows,
        "districtCentroids": dict(sorted(centroids.items())),
        "fullPostcodeLocations": dict(sorted(FULL_POSTCODE_LOCATIONS.items())),
    }
    generated = (
        "// Generated from Final View in Skipton - CGI partners travel.xlsx.\n"
        "window.skiptonMapData = "
        + json.dumps(payload, indent=2, ensure_ascii=True)
        + ";\n"
    )

    if DATA_PATH.exists() and DATA_PATH.read_text() == generated:
        print(f"{DATA_PATH.name} is already up to date with {len(rows)} spreadsheet rows")
    else:
        DATA_PATH.write_text(generated)
        print(f"Wrote {len(rows)} spreadsheet rows into {DATA_PATH.name}")

    print(f"Centroid districts: {len(centroids)}")


if __name__ == "__main__":
    main()
