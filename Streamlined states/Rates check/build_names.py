"""Builds the place names the SST rates check shows beside each code.

The Streamlined rate files carry codes only (state FIPS, county FIPS, Census
place FIPS, and each state's own special district codes), so the macro carries
names for the states this check covers: Ohio, Washington, Nebraska, Oklahoma.

Sources, all in "state files\\names sources":
- Census 2020 county and place code lists (www2.census.gov/geo/docs/reference/codes2020/)
- WA DOR location code list, Rates_26Q4.csv (dor.wa.gov downloadable database)
- OK OTC Currentcsv.csv (COPO names) and a place-to-COPO map pulled from the OU
  rate locator database OTCRates4thQtr2026 (oktax.csa.ou.edu/oktax)
- NE DOR rate chart, July 1, 2026 (Good Life districts)
- Ohio transit codes: matched to counties from the Ohio SST boundary file
  OHB2026Q4SEP17 on Oct 4, 2026 (each code sits in one county)

Writes "state files\\place names.csv" (the list, one row per code) and puts the
same list into SSTRates.bas between the place names markers.

    python build_names.py
"""
import csv
import re
import unicodedata
from pathlib import Path

HERE = Path(__file__).parent
SRC = HERE / "state files" / "names sources"
RATES = HERE / "state files"
OUT_CSV = RATES / "place names.csv"
BAS = HERE / "SSTRates.bas"
STATES = {"39": "oh", "31": "ne", "40": "ok", "53": "wa"}

# Ohio's transit codes, by the county each one sits in (boundary file, Oct 4, 2026).
# 87100, 93000, 94000, 96000 and 98000 cover part of a county: the part a neighbor's
# transit authority reaches (TARTA into Wood, COTA into Fairfield, Licking, Delaware, Union).
OH_TRANSIT = {
    "02000": "Allen County transit",
    "18000": "Cuyahoga County transit (GCRTA)",
    "25000": "Franklin County transit (COTA)",
    "31000": "Hamilton County transit (Metro)",
    "43000": "Lake County transit (Laketran)",
    "48000": "Lucas County transit (TARTA)",
    "50000": "Mahoning County transit (WRTA)",
    "57000": "Montgomery County transit (RTA)",
    "67000": "Portage County transit (PARTA)",
    "76000": "Stark County transit (SARTA)",
    "77000": "Summit County transit (METRO)",
    "87100": "Wood County, TARTA part",
    "93000": "Fairfield County, COTA part",
    "94000": "Licking County, COTA part",
    "96000": "Delaware County, COTA part",
    "98000": "Union County, COTA part",
}
# Nebraska's chart (July 1, 2026) lists GL801 and GL601 to GL605. The rate file has
# GL801 to GL805. Only GL801 lines up, so the rest keep their codes.
NE_GOOD_LIFE = {
    "GL801": "Good Life District, Avenue One opted-in retailers",
    "GL802": "Good Life District GL802",
    "GL803": "Good Life District GL803",
    "GL804": "Good Life District GL804",
    "GL805": "Good Life District GL805",
}


def plain(s):
    """ASCII only, so the names paste into the VBA editor the same everywhere."""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("|", "/").replace(";", ",").replace('"', "'")
    return re.sub(r"\s+", " ", s).strip()


def census(fips, ab):
    counties, places = {}, {}
    for line in (SRC / ("st%s_%s_cou2020.txt" % (fips, ab))).read_text(encoding="utf-8").splitlines()[1:]:
        p = line.split("|")
        counties[p[2]] = p[4]
    for line in (SRC / ("st%s_%s_place2020.txt" % (fips, ab))).read_text(encoding="utf-8").splitlines()[1:]:
        p = line.split("|")
        places[p[2]] = (p[4], p[5])          # name, type (INCORPORATED PLACE or CENSUS DESIGNATED PLACE)
    return counties, places


def codes_in_files(fips):
    """Every (type, code) any of this state's rate files in "state files" mentions."""
    seen = set()
    for f in RATES.glob("*.csv"):
        if f.name == OUT_CSV.name:
            continue
        for row in csv.reader(f.read_text(encoding="utf-8-sig").splitlines()):
            if len(row) >= 9 and row[0].strip().isdigit() and row[0].strip().zfill(2) == fips:
                t = row[1].strip().zfill(2)
                c = row[2].strip().upper()
                if c.isdigit():
                    c = c.zfill(3 if t == "00" else 2 if t == "45" else 5)
                seen.add((t, c))
    return seen


def title_words(s):
    """OKLAHOMA CITY -> Oklahoma City, for the few OK names that only the OTC list has."""
    return " ".join(w.capitalize() if w.isalpha() else w for w in s.split(" "))


def build():
    rows = []
    for fips, ab in STATES.items():
        counties, places = census(fips, ab)
        seen = codes_in_files(fips)
        st = ab.upper()
        for code, name in counties.items():
            rows.append((st, "00", code, name, "Census 2020 counties"))
        for code, (name, typ) in places.items():
            # Ohio has no city rates, so its places stay out unless a file names one
            if (typ == "INCORPORATED PLACE" and fips != "39") or ("01", code) in seen:
                rows.append((st, "01", code, name, "Census 2020 places"))
        if fips == "39":
            for code, name in OH_TRANSIT.items():
                rows.append((st, "63", code, name, "OH boundary file OHB2026Q4SEP17, by county"))
        if fips == "31":
            for code, name in NE_GOOD_LIFE.items():
                rows.append((st, "63", code, name, "NE DOR rate chart, July 1, 2026"))
        if fips == "40":
            copo = {}
            for r in csv.reader((SRC / "OK OTC Currentcsv.csv").read_text(encoding="latin-1").splitlines()):
                if len(r) > 3 and r[0] == "STS":
                    copo[r[1].zfill(4)] = r[2].strip()
            for r in csv.DictReader((SRC / "OK place to COPO (from OTCRates4thQtr2026).csv").open(encoding="utf-8")):
                code, cp = r["PlaceFIPS"], r["CityCOPO"]
                if ("01", code) in seen and code not in places and cp in copo:
                    rows.append((st, "01", code, title_words(copo[cp]), "OK OTC COPO list, by the OU rate locator database"))
        if fips == "53":
            for r in csv.DictReader((SRC / "WA DOR Rates_26Q4.csv").open(encoding="utf-8-sig")):
                rows.append((st, "63", "L" + r["Code"], "%s (location code %s)" % (r["Name"].strip(), r["Code"]),
                             "WA DOR location codes, Rates_26Q4"))
    rows = [(s, t, c, plain(n), src) for (s, t, c, n, src) in rows]
    rows.sort(key=lambda r: (r[0], r[1], r[2]))
    with OUT_CSV.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["State", "Type", "Code", "Name", "Source"])
        w.writerows(rows)
    return rows


def vba_block(rows):
    """One function per state, each a long string: type|code|name;type|code|name;..."""
    out = ["' ---- place names (made by build_names.py from the Census and state lists) ----", ""]
    for st in ("OH", "NE", "OK", "WA"):
        entries = ["%s|%s|%s" % (t, c, n) for (s, t, c, n, _) in rows if s == st]
        # A procedure has a size limit, so big states get split across a few functions
        chunks = [entries[i:i + 300] for i in range(0, len(entries), 300)]
        for k, chunk in enumerate(chunks, 1):
            out.append("Private Function Names%s%d() As String" % (st, k))
            out.append("    Dim s As String")
            line = ""
            for e in chunk:
                piece = e + ";"
                if len(line) + len(piece) > 900:
                    out.append('    s = s & "%s"' % line)
                    line = ""
                line += piece
            if line:
                out.append('    s = s & "%s"' % line)
            out.append("    Names%s%d = s" % (st, k))
            out.append("End Function")
            out.append("")
        parts = " & ".join("Names%s%d()" % (st, k) for k in range(1, len(chunks) + 1))
        out.append("Private Function Names%s() As String" % st)
        out.append("    Names%s = %s" % (st, parts))
        out.append("End Function")
        out.append("")
    out.append("' ---- end of place names ----")
    return "\r\n".join(out)


def inject(block):
    text = BAS.read_text(encoding="utf-8")
    start = text.index("' ---- place names (made by build_names.py")
    end = text.index("' ---- end of place names ----") + len("' ---- end of place names ----")
    BAS.write_text(text[:start] + block + text[end:], encoding="utf-8", newline="")


if __name__ == "__main__":
    rows = build()
    print("%d names written to %s" % (len(rows), OUT_CSV.name))
    for st in ("OH", "NE", "OK", "WA"):
        print("  %s: %d" % (st, sum(1 for r in rows if r[0] == st)))
    if BAS.exists():
        inject(vba_block(rows))
        print("names put into %s" % BAS.name)
