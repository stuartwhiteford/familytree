"""
Parse Whiteford Almanac HTML files and generate a family tree CSV.
"""
import csv
import os
import re
from bs4 import BeautifulSoup

ALMANAC_DIR = os.path.join(os.path.dirname(__file__), "almanac")
OUTPUT_CSV = os.path.join(os.path.dirname(__file__), "data", "whiteford_family_tree.csv")

# Patterns applied to plain text of person blocks
BORN_YEAR_RE = re.compile(
    r"(?:was born|born)[^.]{0,80}?\b(1[0-9]{3}|20[0-2][0-9])\b",
    re.IGNORECASE,
)
DIED_YEAR_RE = re.compile(
    r"(?:died|d\.)[^.]{0,80}?\b(1[0-9]{3}|20[0-2][0-9])\b",
    re.IGNORECASE,
)
BORN_PLACE_RE = re.compile(
    r"(?:was born|born)[^.]*?\bin\s+([A-Z][^.,]{3,60}?)(?:\.|,|\s+(?:and|He|She|They|Parents|Children|aged)\b)",
    re.IGNORECASE,
)
DIED_PLACE_RE = re.compile(
    r"(?:died|d\.)[^.]*?\bin\s+([A-Z][^.,]{3,60}?)(?:\.|,|\s+(?:and|He|She|They|Parents|Children|aged)\b)",
    re.IGNORECASE,
)
# "married to NAME on DATE in PLACE" — we just want the href
MARRIED_RE = re.compile(r"married to\s*$", re.IGNORECASE)


def href_to_pid(href: str) -> str:
    """Extract person ID from href like 'd32.htm#P3204' -> 'P3204'."""
    if href and "#P" in href:
        return href.split("#")[1]
    return ""


def clean(text: str) -> str:
    return " ".join(text.split()).strip()


def parse_block_html(block_html: str, anchor_id: str) -> dict:
    """Parse a single person's HTML block (string from anchor to <HR>)."""
    soup = BeautifulSoup(block_html, "html.parser")

    # Name from <B> tag
    b = soup.find("b")
    name = clean(b.get_text()) if b else ""

    plain = clean(soup.get_text())

    birth_year = ""
    m = BORN_YEAR_RE.search(plain)
    if m:
        birth_year = m.group(1)

    death_year = ""
    m = DIED_YEAR_RE.search(plain)
    if m:
        death_year = m.group(1)

    birth_place = ""
    m = BORN_PLACE_RE.search(plain)
    if m:
        birth_place = clean(m.group(1))

    death_place = ""
    m = DIED_PLACE_RE.search(plain)
    if m:
        death_place = clean(m.group(1))

    # Find spouse: first <A> that is immediately after "married to"
    spouse_id = ""
    # Work on raw HTML to find the link following "married to"
    married_link_re = re.compile(
        r"married to\s*<a\s[^>]*href=['\"]([^'\"]+)['\"]", re.IGNORECASE
    )
    mm = married_link_re.search(block_html)
    if mm:
        spouse_id = href_to_pid(mm.group(1))

    # Find parents: text "Parents: <A>Father</A> and <A>Mother</A>"
    father_id = ""
    mother_id = ""
    parents_re = re.compile(
        r"Parents?:\s*((?:<a\s[^>]*>.*?</a>|[^<P])*)", re.IGNORECASE | re.DOTALL
    )
    pm = parents_re.search(block_html)
    if pm:
        parents_html = pm.group(1)
        parent_links = re.findall(
            r'<a\s[^>]*href=[\'"]([^\'"]+)[\'"]', parents_html, re.IGNORECASE
        )
        pids = [href_to_pid(h) for h in parent_links if href_to_pid(h)]
        if pids:
            father_id = pids[0]
        if len(pids) >= 2:
            mother_id = pids[1]

    # Notes: plain text without name, truncated
    notes = plain
    if name and notes.startswith(name):
        notes = notes[len(name):].lstrip(" .")
    notes = notes[:500]

    return {
        "person_id": anchor_id,
        "name": name,
        "birth_year": birth_year,
        "death_year": death_year,
        "birth_place": birth_place,
        "death_place": death_place,
        "father_id": father_id,
        "mother_id": mother_id,
        "spouse_id": spouse_id,
        "notes": notes,
    }


def get_html_files() -> list:
    skip = {"index.htm", "fowsndx.htm", "fowndx.htm"}
    files = []
    for fn in sorted(os.listdir(ALMANAC_DIR)):
        if fn.endswith(".htm") and fn not in skip:
            files.append(os.path.join(ALMANAC_DIR, fn))
    return files


def parse_all_files() -> dict:
    records: dict = {}

    # Split file into person blocks using regex on raw HTML
    anchor_split_re = re.compile(
        r'<A\s+NAME="(P\d+)"', re.IGNORECASE
    )
    hr_re = re.compile(r"<HR\b", re.IGNORECASE)

    for filepath in get_html_files():
        try:
            with open(filepath, encoding="utf-8", errors="replace") as f:
                html = f.read()
        except OSError:
            continue

        # Find all person anchors and split the HTML into blocks
        for m in anchor_split_re.finditer(html):
            pid = m.group(1)
            start = m.start()
            # Find the next <HR> after this anchor
            hr_match = hr_re.search(html, m.end())
            end = hr_match.start() if hr_match else len(html)
            block_html = html[start:end]

            record = parse_block_html(block_html, pid)

            if pid not in records:
                records[pid] = record
            else:
                # Merge: fill empty fields from later occurrences
                existing = records[pid]
                for key in ("name", "birth_year", "death_year", "birth_place",
                            "death_place", "father_id", "mother_id", "spouse_id"):
                    if not existing[key] and record[key]:
                        existing[key] = record[key]

    return records


def write_csv(records: dict) -> None:
    os.makedirs(os.path.dirname(OUTPUT_CSV), exist_ok=True)
    fieldnames = [
        "person_id", "name", "birth_year", "death_year",
        "birth_place", "death_place", "father_id", "mother_id",
        "spouse_id", "notes",
    ]
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for record in sorted(records.values(), key=lambda r: int(r["person_id"][1:])):
            writer.writerow(record)
    print(f"Wrote {len(records)} records to {OUTPUT_CSV}")


if __name__ == "__main__":
    print("Parsing HTML files...")
    records = parse_all_files()
    write_csv(records)
