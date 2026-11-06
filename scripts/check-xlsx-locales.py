"""Check the factual native locale table against the primary MS-LCID table.

Pass --source to reuse a downloaded page. No Office implementation code is
downloaded or used; only assigned numeric language IDs and language tags.
"""
from html.parser import HTMLParser
from pathlib import Path
from urllib.request import urlopen
import argparse
import re

ROOT = Path(__file__).resolve().parent.parent
URL = "https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-lcid/63d3d639-7fd2-4afb-abbe-0d5b5551eef8"


class LanguageTable(HTMLParser):
    def __init__(self):
        super().__init__(); self.rows = []; self.row = []; self.cell = None

    def handle_starttag(self, tag, attrs):
        if tag == "tr": self.row = []
        if tag in ["th", "td"]: self.cell = ""

    def handle_data(self, data):
        if self.cell is not None: self.cell += data

    def handle_endtag(self, tag):
        if tag in ["th", "td"] and self.cell is not None:
            self.row.append(" ".join(self.cell.split())); self.cell = None
        if tag == "tr" and self.row: self.rows.append(self.row)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--source", type=Path); args = parser.parse_args()
    if args.source: page = args.source.read_text()
    else:
        with urlopen(URL, timeout=45) as response: page = response.read(1_000_000).decode("utf-8")
    table = LanguageTable(); table.feed(page); rows = []
    for row in table.rows:
        if len(row) == 2 and re.fullmatch(r"0x[0-9A-Fa-f]{4}", row[0]) and re.fullmatch(r"[a-zA-Z]{2,3}(?:-[a-zA-Z0-9]{2,8})*(?:_tradnl)?", row[1]):
            rows.append((int(row[0], 16), row[1].removesuffix("_tradnl")))
    assert len(dict(rows)) == len(rows), "Duplicate assigned language ID"
    source = (ROOT / "packages/core/src/workbook/native-locales.ts").read_text()
    actual = {int(key, 16): tag for key, tag in re.findall(r'0x([0-9a-f]{4}): "([^"]+)"', source)}
    assert len(rows) == 412 and actual == dict(rows), "The primary assigned language table differs from the recorded mappings"
    print(f"Verified {len(rows)} assigned native locale mappings against MS-LCID")


if __name__ == "__main__": main()
