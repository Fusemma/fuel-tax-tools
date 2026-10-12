"""Reference address matcher for the Missouri DOR boundary file.

Given a ship-to street, zip and optional plus-4, find its 13-digit
jurisdiction code (city-county-district, like 37000-051-000) the way the
DOR's guide says to: address row first, then zip+4 row, then 5-digit zip
row. This Python copy is the test oracle; the VBA copy that ships must
give the same answers on the same inputs (see test_match.py).

Kept to plain string work (split, replace, dictionaries) so it ports to
VBA without regular expressions.
"""
import csv, io, zipfile, collections

# Boundary file columns (0-based). The file's own header row is shifted
# after the first district column, so positions are fixed here instead of
# read from the header. The jurisdiction code sits under the header named
# "Special Tax District Code (#20)".
C_TYPE, C_BEGIN, C_END, C_LOW, C_HIGH, C_PARITY = 0, 1, 2, 3, 4, 5
C_PRE, C_NAME, C_SUF, C_POST = 6, 7, 8, 9
C_CITY, C_ZIP, C_PLUS4 = 14, 15, 16
C_ZLO, C_XLO, C_ZHI, C_XHI = 17, 18, 19, 20
C_COUNTY, C_PLACE = 24, 25
C_JCODE = 69

DIRS = {"NORTH": "N", "SOUTH": "S", "EAST": "E", "WEST": "W",
        "NORTHEAST": "NE", "NORTHWEST": "NW", "SOUTHEAST": "SE", "SOUTHWEST": "SW"}
DIR_SET = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}

# USPS Publication 28 street suffixes, full word or common variant -> standard.
SUFFIX_WORDS = """
ALLEY ALY|ALLEE ALY|ALLY ALY|ANNEX ANX|ANEX ANX|ARCADE ARC|AVENUE AVE|AV AVE|AVEN AVE|AVENU AVE
AVN AVE|AVNUE AVE|BAYOU BYU|BEACH BCH|BEND BND|BLUFF BLF|BLUF BLF|BOTTOM BTM|BOULEVARD BLVD
BOUL BLVD|BOULV BLVD|BRANCH BR|BRIDGE BRG|BROOK BRK|BYPASS BYP|CAMP CP|CANYON CYN|CAPE CPE
CAUSEWAY CSWY|CENTER CTR|CENTRE CTR|CENT CTR|CNTR CTR|CNTER CTR|CIRCLE CIR|CIRC CIR|CIRCL CIR
CRCL CIR|CRCLE CIR|CLIFF CLF|CLUB CLB|COMMON CMN|CORNER COR|CORNERS CORS|COURSE CRSE|COURT CT
CRT CT|COVE CV|CREEK CRK|CRESCENT CRES|CREST CRST|CROSSING XING|CRSSNG XING|CROSSROAD XRD
CURVE CURV|DALE DL|DAM DM|DIVIDE DV|DRIVE DR|DRIV DR|DRV DR|ESTATE EST|ESTATES EST|ESTS EST
EXPRESSWAY EXPY|EXPR EXPY|EXPRESS EXPY|EXPW EXPY|EXTENSION EXT|EXTN EXT|EXTNSN EXT|FALLS FLS
FERRY FRY|FIELD FLD|FIELDS FLDS|FLAT FLT|FOREST FRST|FORGE FRG|FORK FRK|FORT FT|FREEWAY FWY
GARDEN GDN|GARDENS GDNS|GATEWAY GTWY|GLEN GLN|GREEN GRN|GROVE GRV|HARBOR HBR|HAVEN HVN
HEIGHTS HTS|HEIGHT HTS|HT HTS|HIGHWAY HWY|HIGHWY HWY|HIWAY HWY|HIWY HWY|HWAY HWY|HILL HL
HILLS HLS|HOLLOW HOLW|HOLLOWS HOLW|HLLW HOLW|INLET INLT|ISLAND IS|JUNCTION JCT|JCTION JCT
KNOLL KNL|LAKE LK|LAKES LKS|LANDING LNDG|LNDNG LNDG|LANE LN|MANOR MNR|MEADOW MDW|MEADOWS MDWS
MILL ML|MILLS MLS|MOUNT MT|MOUNTAIN MTN|ORCHARD ORCH|PARKS PARK|PARKWAY PKWY|PARKWY PKWY
PKWAY PKWY|PKY PKWY|PINE PNE|PINES PNES|PLACE PL|PLAIN PLN|PLAINS PLNS|PLAZA PLZ|PLZA PLZ
POINT PT|POINTS PTS|PORT PRT|PRAIRIE PR|RANCH RNCH|RIDGE RDG|RDGE RDG|RIVER RIV|ROAD RD
ROUTE RTE|RT RTE|SHORE SHR|SHORES SHRS|SPRING SPG|SPRINGS SPGS|SQUARE SQ|SQR SQ|SQU SQ
STATION STA|STREET ST|STR ST|STRT ST|SUMMIT SMT|TERRACE TER|TERR TER|TRACE TRCE|TRACK TRAK
TRAFFICWAY TRFY|TRAIL TRL|TRAILS TRL|TRLS TRL|TUNNEL TUNL|TURNPIKE TPKE|VALLEY VLY|VIEW VW
VILLAGE VLG|VILLE VL|VISTA VIS|WELL WL|WELLS WLS|SAINT ST|INTERSTATE I|IH I|COUNTY CO
BUSINESS BUS|PRIVATE PVT|MISSOURI MO
"""
SYN = dict(DIRS)
for pair in SUFFIX_WORDS.replace("\n", "|").split("|"):
    pair = pair.strip()
    if pair:
        a, b = pair.split(" ")
        SYN[a] = b
SUFFIX_SET = set(b for b in SYN.values()) | {"ST", "RD", "DR", "AVE", "LN", "CT", "BLVD", "TER",
    "PL", "CIR", "HWY", "TRL", "PKWY", "WAY", "TRFY", "PLZ", "RTE", "LOOP", "RDG", "CTR", "EXPY",
    "SQ", "PT", "PASS", "EXT", "CV", "PARK", "RUN", "HL", "EST", "MALL", "XING", "LNDG", "SPUR",
    "HOLW", "WALK", "CRK", "TRCE", "BND", "HTS", "VW", "HLS", "MNR", "VLY", "PATH", "MDWS", "GRV",
    "VLG", "ALY", "CRST", "ROW", "FRST", "BLF", "GLN", "HVN", "PIKE", "OVAL"}
SUFFIX_SET -= DIR_SET | {"I", "CO", "BUS", "PVT", "MO"}

# Words that start a unit (apartment, suite, lot); the unit and anything after it drop.
UNIT_WORDS = {"APT", "APARTMENT", "STE", "SUITE", "UNIT", "LOT", "TRLR", "TRAILER", "OFC",
              "OFFICE", "FL", "FLOOR", "RM", "ROOM", "BSMT", "BLDG", "BUILDING", "SPC", "SPACE",
              "DEPT", "#", "NO", "REAR", "FRNT", "LOWR", "UPPR", "LBBY"}


def cut_unit(tok):
    """Drop a unit ('Ste 5', 'Lot 4', 'Rear') and everything after it. Never
    the first word, so a street named Lot Rd or No Name Rd survives."""
    for k, w in enumerate(tok):
        if w in UNIT_WORDS and k > 0:
            return tok[:k]
    return tok


def tokens(text):
    """Uppercase, drop periods and apostrophes, turn anything else that isn't
    a letter, digit or # into a space, map synonyms. Plain A-Z only, so the
    VBA copy (which can't lean on Python's upper()) gives the same words."""
    t = []
    for c in text or "":
        if c in ".'":
            continue
        if "a" <= c <= "z":
            c = chr(ord(c) - 32)
        t.append(c if ("A" <= c <= "Z" or "0" <= c <= "9" or c == "#") else " ")
    t = "".join(t).replace("#", " # ")
    return [SYN.get(w, w) for w in t.split()]


def is_route_id(w):
    """A highway id: a number (50, 13A) or one to three letters (K, CC, AB)."""
    if not w:
        return False
    if w[0].isdigit():
        return all(c.isalnum() for c in w)
    return w.isalpha() and len(w) <= 3 and w not in SUFFIX_SET and w not in DIR_SET


def canon(tok):
    """Collapse the many ways people write the same road to one spelling.
    STATE HWY K, STATE RTE K, STATE RD K, ROUTE K, MO 13, US HWY 61, US 61 -> HWY <id>;
    COUNTY ROAD 12, CO RD 12 -> CR 12; FARM ROAD 12 -> FR 12."""
    out, i, n = [], 0, len(tok)
    while i < n:
        w = tok[i]
        nxt = tok[i + 1] if i + 1 < n else ""
        nxt2 = tok[i + 2] if i + 2 < n else ""
        if w in ("STATE", "MO", "US") and nxt in ("HWY", "RTE", "RD") and is_route_id(nxt2):
            out += ["HWY", nxt2]; i += 3; continue
        if w in ("MO", "US") and is_route_id(nxt) and nxt[0].isdigit():
            out += ["HWY", nxt]; i += 2; continue
        if w in ("HWY", "RTE") and is_route_id(nxt):
            out += ["HWY", nxt]; i += 2; continue
        if w == "CO" and nxt == "RD" and is_route_id(nxt2):
            out += ["CR", nxt2]; i += 3; continue
        if w == "FARM" and nxt == "RD" and is_route_id(nxt2):
            out += ["FR", nxt2]; i += 3; continue
        out.append(w); i += 1
    return out


def one(word):
    """A direction or suffix field from the state's file, standardized ('' if blank)."""
    t = tokens(word)
    return t[0] if t else ""


def split_street(line):
    """'1234 N Main St Ste 5' -> ('1234', 1234, ['N','MAIN','ST'])."""
    tok = tokens(line)
    num_s, num_i = "", None
    if tok and tok[0][0].isdigit():
        first = tok.pop(0)
        digits = ""
        for c in first:
            if c.isdigit():
                digits += c
            else:
                break
        num_s, num_i = first, int(digits)
        # '123 1/2 Main' -> the fraction became '1 2'; drop it
        if len(tok) >= 2 and tok[0] == "1" and tok[1] == "2":
            tok = tok[2:]
    return num_s, num_i, canon(cut_unit(tok))


def readings(tok):
    """Every sensible way to split the ship-to's street into
    (direction before, name, suffix, direction after). Whether the last word
    is a suffix or part of the name ('Henley Lake' vs 'Henley Lake Rd') can't
    be told from the words alone, so both readings get tried."""
    out = []
    n = len(tok)
    for use_pre in (True, False):
        for use_post in (True, False):
            for use_suf in (True, False):
                a, z = 0, n
                pre = post = suf = ""
                if use_pre:
                    if z - a > 1 and tok[a] in DIR_SET:
                        pre = tok[a]; a += 1
                    else:
                        continue
                if use_post:
                    if z - a > 1 and tok[z - 1] in DIR_SET:
                        post = tok[z - 1]; z -= 1
                    else:
                        continue
                if use_suf:
                    if z - a > 1 and tok[z - 1] in SUFFIX_SET and not (
                            z - a == 2 and tok[a] in ("HWY", "CR", "FR", "I")):
                        suf = tok[z - 1]; z -= 1
                    else:
                        continue
                out.append((pre, tuple(tok[a:z]), suf, post))
    return out


def fits(reading, row):
    """The reading names the same street as the row, and nothing it says
    contradicts the row (a direction or suffix it leaves out is fine)."""
    pre, name, suf, post = reading
    if name != row["name"]:
        return False
    if suf and suf != row["suf"]:
        return False
    said = {d for d in (pre, post) if d}
    has = {d for d in (row["pre"], row["post"]) if d}
    if said and has and not (said & has):
        return False
    return True


def contains(row, num_s, num_i):
    """Is house number inside the state's range, on the right side of the street?"""
    low, high, parity = row["low"], row["high"], row["parity"]
    if num_i is None:
        return False
    if low.isdigit() and high.isdigit():
        lo, hi = min(int(low), int(high)), max(int(low), int(high))  # a few rows run backwards
        if not (lo <= num_i <= hi):
            return False
        if parity == "O" and num_i % 2 == 0:
            return False
        if parity == "E" and num_i % 2 == 1:
            return False
        return True
    # lettered numbers like 1616B: exact text match only
    return num_s == low or num_s == high


def width(row):
    """How many house numbers a row covers; a single address is 0."""
    if row["low"].isdigit() and row["high"].isdigit():
        return abs(int(row["high"]) - int(row["low"]))
    return 0


def narrowest(hits):
    """The state lists wide ranges, then single addresses inside them that sit
    in a different place (a house on a city line, a building in a district).
    On the same street the narrower row is the specific exception, so it wins.
    Different streets (W 46th St and W 46th Ter) never knock each other out."""
    by_street = collections.defaultdict(list)
    for h in hits:
        by_street[h["full"]].append(h)
    out = []
    for group in by_street.values():
        w = min(width(h) for h in group)
        out += [h for h in group if width(h) == w]
    return out


def city_key(city):
    """Postal city, standardized so the state's own spellings meet:
    SAINT LOUIS = ST. LOUIS, O'FALLON = O FALLON."""
    t = []
    for c in city or "":
        if c in ".'":
            continue
        if "a" <= c <= "z":
            c = chr(ord(c) - 32)
        t.append(c if ("A" <= c <= "Z" or "0" <= c <= "9") else " ")
    words = {"SAINT": "ST", "SAINTE": "STE", "MOUNT": "MT", "FORT": "FT"}
    return "".join(words.get(w, w) for w in "".join(t).split())


# County codes as the DOR writes them in jurisdiction codes, names from the
# Census Bureau. The DOR uses 193 for Ste. Genevieve County (the Census moved
# it to 186).
COUNTIES = dict(x.split(" ", 1) for x in """001 Adair|003 Andrew|005 Atchison|007 Audrain|009 Barry|011 Barton|013 Bates
|015 Benton|017 Bollinger|019 Boone|021 Buchanan|023 Butler|025 Caldwell|027 Callaway|029 Camden|031 Cape Girardeau
|033 Carroll|035 Carter|037 Cass|039 Cedar|041 Chariton|043 Christian|045 Clark|047 Clay|049 Clinton|051 Cole
|053 Cooper|055 Crawford|057 Dade|059 Dallas|061 Daviess|063 DeKalb|065 Dent|067 Douglas|069 Dunklin|071 Franklin
|073 Gasconade|075 Gentry|077 Greene|079 Grundy|081 Harrison|083 Henry|085 Hickory|087 Holt|089 Howard|091 Howell
|093 Iron|095 Jackson|097 Jasper|099 Jefferson|101 Johnson|103 Knox|105 Laclede|107 Lafayette|109 Lawrence|111 Lewis
|113 Lincoln|115 Linn|117 Livingston|119 McDonald|121 Macon|123 Madison|125 Maries|127 Marion|129 Mercer|131 Miller
|133 Mississippi|135 Moniteau|137 Monroe|139 Montgomery|141 Morgan|143 New Madrid|145 Newton|147 Nodaway|149 Oregon
|151 Osage|153 Ozark|155 Pemiscot|157 Perry|159 Pettis|161 Phelps|163 Pike|165 Platte|167 Polk|169 Pulaski
|171 Putnam|173 Ralls|175 Randolph|177 Ray|179 Reynolds|181 Ripley|183 St. Charles|185 St. Clair|193 Ste. Genevieve
|187 St. Francois|189 St. Louis|195 Saline|197 Schuyler|199 Scotland|201 Scott|203 Shannon|205 Shelby|207 Stoddard
|209 Stone|211 Sullivan|213 Taney|215 Texas|217 Vernon|219 Warren|221 Washington|223 Wayne|225 Webster|227 Worth
|229 Wright|510 St. Louis City""".replace("\n", "").split("|"))


def county_key(name):
    """'St. Louis County', 'ST LOUIS CO', 'Saint Louis' -> 'STLOUIS';
    'St. Louis City' -> 'STLOUISCITY'."""
    t = []
    for c in name or "":
        if c in ".'":
            continue
        if "a" <= c <= "z":
            c = chr(ord(c) - 32)
        t.append(c if ("A" <= c <= "Z" or "0" <= c <= "9") else " ")
    swap = {"SAINT": "ST", "SAINTE": "STE"}
    return "".join(swap.get(w, w) for w in "".join(t).split() if w not in ("COUNTY", "CO", "CNTY"))


COUNTY_BY_KEY = {county_key(v): k for k, v in COUNTIES.items()}
COUNTY_BY_KEY.update({"STGENEVIEVE": "193", "STLOUISCO": "189"})


def county_code(name):
    """IT's county column -> the DOR's 3-digit county code, '' if unknown."""
    k = county_key(name)
    if k.isdigit() and len(k) <= 3:
        return k.zfill(3)
    return COUNTY_BY_KEY.get(k, "")


def county_agrees(name, code):
    """Does the list's county name fit the state's county code? A county the
    names can't be read for, and a plain 'St. Louis' (city or county), pass."""
    cc = county_code(name)
    if not cc:
        return True
    if county_key(name) == "STLOUIS" and code in ("189", "510"):
        return True
    return cc == code


class Boundary:
    def __init__(self, path, as_of):
        """as_of is the date the answers are for, like '20261001'. Rows whose
        begin and end dates don't cover it are skipped, per the DOR guide."""
        self.by_zip = collections.defaultdict(list)
        self.by_city_name = collections.defaultdict(list)    # (city key, street name) -> rows
        self.by_county_name = collections.defaultdict(list)  # (county code, street name) -> rows
        self.zip4 = []
        self.zip5 = []
        if path.lower().endswith(".zip"):
            zf = zipfile.ZipFile(path)
            f = io.TextIOWrapper(zf.open(zf.namelist()[0]), encoding="latin-1", newline="")
        else:
            f = open(path, newline="", encoding="latin-1")
        with f:
            r = csv.reader(f)
            next(r)
            for row in r:
                if not (row[C_BEGIN] <= as_of <= row[C_END]):
                    continue
                t = row[C_TYPE]
                if t == "A":
                    pre, suf, post = one(row[C_PRE]), one(row[C_SUF]), one(row[C_POST])
                    # a few hundred rows carry a unit inside the street name
                    # ('LINDBERGH BLVD REAR UNIT'); cut it the same way as a ship-to's
                    name = canon(cut_unit(tokens(row[C_NAME])))
                    rec = {
                        "low": row[C_LOW].strip(), "high": row[C_HIGH].strip(), "parity": row[C_PARITY].strip(),
                        "pre": pre, "name": tuple(name), "suf": suf, "post": post,
                        "full": tuple(canon([x for x in [pre] + name + [suf, post] if x])),
                        "code": row[C_JCODE], "city": row[C_CITY], "zip": row[C_ZIP],
                        "raw": (row[C_PRE].strip(), row[C_NAME].strip(), row[C_SUF].strip(), row[C_POST].strip()),
                        "said":" ".join(x for x in (row[C_PRE], row[C_NAME], row[C_SUF], row[C_POST]) if x.strip())}
                    self.by_zip[row[C_ZIP]].append(rec)
                    self.by_city_name[(city_key(row[C_CITY]), rec["name"])].append(rec)
                    self.by_county_name[(row[C_JCODE][6:9], rec["name"])].append(rec)
                elif t == "4":
                    self.zip4.append((row[C_ZLO] + row[C_XLO], row[C_ZHI] + row[C_XHI], row[C_JCODE]))
                elif t == "Z":
                    self.zip5.append((row[C_ZLO], row[C_ZHI], row[C_JCODE]))

    def match(self, street, zip5, plus4=""):
        """Returns (how, code, note). how is one of
        EXACT  - the street as written, number in the state's range
        CLOSE  - same street with a word left off or spelled another way
        STREET - right street, number outside the state's ranges, and the
                 whole street in this zip is one code
        ZIP4 / ZIP - no street match; fell back to the zip+4 or zip row
                 (the zip row is the lowest rate in the zip)
        MIXED  - the state's rows disagree for what was written
        NONE   - zip not in the file"""
        zip5 = (zip5 or "").strip()[:5]
        plus4 = (plus4 or "").strip()
        num_s, num_i, tok = split_street(street)
        full = tuple(tok)
        rows = self.by_zip.get(zip5, [])

        hits = narrowest([r for r in rows if r["full"] == full and contains(r, num_s, num_i)])
        codes = sorted(set(r["code"] for r in hits))
        if len(codes) == 1:
            return "EXACT", codes[0], self._say(hits[0])
        if len(codes) > 1:
            return "MIXED", "", "the state's file has more than one code here: " + ", ".join(codes)

        reads = readings(tok)
        street_rows = [r for r in rows if any(fits(rd, r) for rd in reads)]
        hits = narrowest([r for r in street_rows if contains(r, num_s, num_i)])
        codes = sorted(set(r["code"] for r in hits))
        if len(codes) == 1:
            return "CLOSE", codes[0], self._say(hits[0])
        if len(codes) > 1:
            return "MIXED", "", "more than one street fits what was written: " + ", ".join(codes)

        codes = sorted(set(r["code"] for r in street_rows))
        if len(codes) == 1 and num_i is not None:
            return "STREET", codes[0], "number not in the state's ranges; the whole street here is " + codes[0]

        if len(plus4) == 4:
            z9 = zip5 + plus4
            for lo, hi, code in self.zip4:
                if lo <= z9 <= hi:
                    return "ZIP4", code, "zip+4 row"
        for zlo, zhi, code in self.zip5:
            if zlo <= zip5 <= zhi:
                return "ZIP", code, "5-digit zip row, lowest rate in the zip"
        return "NONE", "", "zip not in the state's file"

    def match_ship_to(self, street, city, county=""):
        """match_city plus the county cross-check: if the list's county names a
        different county than the state's row, it goes to the map."""
        how, code, note, where = self.match_city(street, city, county)
        if code and county and not county_agrees(county, code[6:9]):
            return ("MIXED", "", f"the list says {county}, but the state's row is in "
                    f"{COUNTIES.get(code[6:9], code[6:9])} ({code})", where)
        return how, code, note, where

    def match_city(self, street, city, county=""):
        """For a list with no zip (IT's ship-to list). The street is looked for
        under the city on the mail (the state's City Name is the postal city);
        if it isn't placed there and a county is given, anywhere in that county.
        Returns (how, code, note, where): how is EXACT, CLOSE, STREET or MIXED
        as in match, or NONE; where is 'city', 'county' or ''."""
        num_s, num_i, tok = split_street(street)
        if not tok:
            return "NONE", "", "no street on that line", ""
        reads = readings(tok)
        names = sorted(set(rd[1] for rd in reads))
        ck = city_key(city)
        rows = [r for n in names for r in self.by_city_name.get((ck, n), [])]
        res = self._match_rows(num_s, num_i, tok, reads, rows)
        if res[0] in ("EXACT", "CLOSE", "MIXED"):
            return res + ("city",)
        cc = county_code(county) if county else ""
        if cc:
            rows = [r for n in names for r in self.by_county_name.get((cc, n), [])]
            res2 = self._match_rows(num_s, num_i, tok, reads, rows)
            if res[0] == "STREET":
                # the house number isn't on that street under that city: either
                # a house the file doesn't have yet, or the city on the mail is
                # another name for a different town. The county has to agree.
                if res2[0] == "NONE":
                    return res + ("city",)
                if res2[1] == res[1]:
                    return (res2 + ("county",)) if res2[0] in ("EXACT", "CLOSE") else (res + ("city",))
                return ("MIXED", "", "that street under that city is " + res[1] + ", but the county's rows say " +
                        (res2[1] or "more than one code"), "county")
            if res2[0] != "NONE":
                return res2 + ("county",)
        return res + (("city",) if res[0] == "STREET" else ("",))

    def _match_rows(self, num_s, num_i, tok, reads, rows):
        """The street steps of match, on rows already picked by name."""
        full = tuple(tok)
        hits = narrowest([r for r in rows if r["full"] == full and contains(r, num_s, num_i)])
        codes = sorted(set(r["code"] for r in hits))
        if len(codes) == 1:
            return "EXACT", codes[0], self._say(hits[0])
        if len(codes) > 1:
            return "MIXED", "", "the state's file has more than one code here: " + ", ".join(codes)
        street_rows = [r for r in rows if any(fits(rd, r) for rd in reads)]
        hits = narrowest([r for r in street_rows if contains(r, num_s, num_i)])
        codes = sorted(set(r["code"] for r in hits))
        if len(codes) == 1:
            return "CLOSE", codes[0], self._say(hits[0])
        if len(codes) > 1:
            return "MIXED", "", "more than one street fits what was written: " + ", ".join(codes)
        codes = sorted(set(r["code"] for r in street_rows))
        if len(codes) == 1 and num_i is not None:
            return "STREET", codes[0], "that house number isn't in the state's file; the rest of the street is " + codes[0]
        if street_rows and num_i is None:
            return "NONE", "", "no house number on that line"
        if street_rows:
            return "NONE", "", "the street is there, but not that house number"
        return "NONE", "", "street not found there"

    @staticmethod
    def _say(r):
        return f"{r['low']}-{r['high']} {r['said']}, {r['city']}"
