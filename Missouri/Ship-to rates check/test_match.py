"""Tests the reference matcher on the real Q4 2026 boundary file.

1. The DOR guide's own example must come back right.
2. Addresses built from random rows of the state's file, written the way
   the state writes them, must match their row (EXACT).
3. The same addresses written the messy ways people type them (spelled-out
   words, periods, suite numbers, Hwy/Route/State Rd, dropped directions)
   must land on the same jurisdiction code.
Run: python test_match.py [sample size]
"""
import sys, random, collections, time
from pathlib import Path
from mo_match import Boundary, SYN, DIRS, contains, narrowest

HERE = Path(__file__).parent
N = int(sys.argv[1]) if len(sys.argv) > 1 else 3000
rng = random.Random(20260930)

t0 = time.time()
b = Boundary(str(HERE / "state files" / "dor-boundary-database-q4-2026.zip"), "20261001")
print(f"loaded in {time.time() - t0:.0f}s: {sum(len(v) for v in b.by_zip.values())} address rows, "
      f"{len(b.zip4)} zip+4 rows, {len(b.zip5)} zip rows")

fails = 0
def check(name, got, want):
    global fails
    ok = got == want
    if not ok:
        fails += 1
    print(("PASS " if ok else "FAIL ") + name + ("" if ok else f"  got {got!r} want {want!r}"))

# 1. the guide's example (page 9 and 13)
how, code, note = b.match("201 W Capitol Ave", "65101", "6809")
check("guide example 201 W Capitol Ave 65101 -> 37000-051-000", (how, code), ("EXACT", "37000-051-000"))
how, code, note = b.match("201 West Capitol Avenue", "65101")
check("same, spelled out", (how, code), ("EXACT", "37000-051-000"))

# highways, written the ways people write them (row: 3200-3298 even, HIGHWAY 50 E, Linn 65051)
for s in ("3200 Highway 50 E", "3200 Hwy 50 East", "3200 US Hwy 50 E", "3200 US-50 E", "3200 State Hwy 50 E"):
    check(f"{s!r} 65051 -> 00000-151-002", b.match(s, "65051")[1], "00000-151-002")
check("'3201 Hwy 50 E' is the odd side, not in the even row", b.match("3201 Hwy 50 E", "65051")[0] != "EXACT", True)
# a street whose last word looks like a suffix, with the real suffix left off
check("'5400 Heege Rd' 63123 -> 00000-189-026", b.match("5400 Heege Rd", "63123")[1], "00000-189-026")
check("'5400 Heege' (no Rd) 63123 -> 00000-189-026", b.match("5400 Heege", "63123")[1], "00000-189-026")
# units
check("'6211 N Lindbergh Blvd Rear Unit' 63042", b.match("6211 N Lindbergh Blvd Rear Unit", "63042")[:2],
      ("EXACT", "31276-189-000"))
check("'201 W Capitol Ave Suite 4' 65101", b.match("201 W Capitol Ave Suite 4", "65101")[1], "37000-051-000")
# a single house carved out of a wider range wins; its neighbors keep the range's code
check("'15933 Clayton Rd' 63011 (carved-out house) -> 14176-189-000",
      b.match("15933 Clayton Rd", "63011")[1], "14176-189-000")
check("'15935 Clayton Rd' 63011 (the range around it) -> 21898-189-000",
      b.match("15935 Clayton Rd", "63011")[1], "21898-189-000")
check("'174 Clarkson Rd Ste 125' 63011 -> 21898-189-000", b.match("174 Clarkson Rd Ste 125", "63011")[1],
      "21898-189-000")
# genuinely unclear: W 46th St and W 46th Ter are different streets with different codes
check("'395 W 46th' 64112 with no suffix stays MIXED", b.match("395 W 46th", "64112")[0], "MIXED")
# nothing to match: falls back to the zip row, never pretends
check("made-up street falls back to ZIP", b.match("123 Nonexistent Street", "65101")[0], "ZIP")
check("no house number falls back to ZIP", b.match("Hwy 63 and County Road 200", "65101")[0], "ZIP")
check("zip outside Missouri -> NONE", b.match("1 Main St", "46802")[0], "NONE")

# 2 and 3. random rows
FULL = {}
for k, v in SYN.items():
    if v not in FULL and len(k) > len(v) and k not in ("SAINT",):
        FULL[v] = k
FULL.update({v: k for k, v in DIRS.items()})
FULL["ST"] = "STREET"

def pick_number(low, high, parity):
    lo, hi = int(low), int(high)
    for _ in range(50):
        n = rng.randint(lo, hi)
        if parity == "O" and n % 2 == 0: continue
        if parity == "E" and n % 2 == 1: continue
        return n
    return None

def messy(pre, name, suf, post):
    words = []
    if pre: words.append(FULL.get(pre, pre) if rng.random() < 0.5 else pre)
    nm = name
    up = name.upper()
    r = rng.random()
    if up.startswith("HIGHWAY ") and r < 0.6:
        nm = rng.choice(["Hwy ", "State Hwy ", "Route ", "State Rte ", "MO Hwy "]) + name[8:]
    elif up.startswith("STATE HIGHWAY ") and r < 0.6:
        nm = rng.choice(["Hwy ", "Highway ", "Route ", "State Rd ", "St Hwy "]).replace("St Hwy ", "State Hwy ") + name[14:]
    elif up.startswith("STATE ROUTE ") and r < 0.6:
        nm = rng.choice(["Hwy ", "Route ", "State Hwy "]) + name[12:]
    elif up.startswith("STATE ROAD ") and r < 0.6:
        nm = rng.choice(["Hwy ", "State Rd ", "Route "]) + name[11:]
    elif up.startswith("US HIGHWAY ") and r < 0.6:
        nm = rng.choice(["US Hwy ", "US ", "Hwy ", "U.S. Highway "]) + name[11:]
    elif up.startswith("COUNTY ROAD ") and r < 0.6:
        nm = rng.choice(["CR ", "Co Rd ", "County Rd ", "Co. Road "]) + name[12:]
    elif up.startswith("FARM ROAD ") and r < 0.6:
        nm = rng.choice(["FR ", "Farm Rd "]) + name[10:]
    words.append(nm)
    if suf and rng.random() > 0.15:
        s = FULL.get(suf.upper(), suf) if rng.random() < 0.5 else suf
        if rng.random() < 0.3 and len(s) <= 4: s += "."
        words.append(s)
    if post and rng.random() > 0.2:
        words.append(FULL.get(post, post) if rng.random() < 0.5 else post)
    text = " ".join(words)
    if rng.random() < 0.5: text = text.title()
    if rng.random() < 0.2: text += rng.choice([" Ste 200", " Apt 3", " #5", ", Suite B", " Unit 12", " Lot 4"])
    return text

pool = [(z, r) for z, rows in b.by_zip.items() for r in rows if r["low"].isdigit() and r["high"].isdigit()]
sample = rng.sample(pool, N)
canon_how = collections.Counter(); messy_how = collections.Counter()
agree = 0; disagree = []; canon_wrong = []
for z, r in sample:
    low, high, parity, code = r["low"], r["high"], r["parity"], r["code"]
    pre, name, suf, post = r["raw"]
    n = pick_number(low, high, parity)
    if n is None:
        continue
    street = " ".join(x for x in (str(n), pre, name, suf, post) if x.strip())
    h1, c1, _ = b.match(street, z)
    canon_how[h1] += 1
    # a narrower row on the same street (a carved-out house) rightly outranks the sampled row
    same = [x for x in b.by_zip[z] if x["full"] == r["full"] and contains(x, str(n), n)]
    want = sorted(set(x["code"] for x in narrowest(same)))
    if len(want) == 1:
        code = want[0]
    if c1 != code and h1 != "MIXED":
        canon_wrong.append((street, z, code, h1, c1))
    m = str(n) + " " + messy(pre.strip(), name.strip(), suf.strip(), post.strip())
    h2, c2, note = b.match(m, z)
    messy_how[h2] += 1
    if c2 == c1:
        agree += 1
    else:
        disagree.append((m, z, h1, c1, h2, c2, note))

total = sum(canon_how.values())
print(f"\nstate's spelling: {dict(canon_how)} of {total}")
check("state's spelling gives the row's own code (MIXED aside)", len(canon_wrong), 0)
for x in canon_wrong[:10]: print("   ", x)
print(f"messy spelling: {dict(messy_how)}")
print(f"messy agrees with state's spelling: {agree} of {total} ({100 * agree / total:.1f}%)")
check("messy spelling agrees at least 98% of the time", agree / total >= 0.98, True)
wrong_code = [d for d in disagree if d[5] and d[4] in ("EXACT", "CLOSE", "STREET")]
print(f"messy gave a DIFFERENT code while claiming a street match: {len(wrong_code)}")
for d in wrong_code[:15]: print("   ", d)
print("other disagreements (fell back or unsure):")
for d in [d for d in disagree if d not in wrong_code][:15]: print("   ", d)
# 4. no zip (IT's ship-to list): the same messy addresses, placed by street and
#    city, then county. City written the ways people write it (Saint/St./St).
print("\nno zip, by city:")
from mo_match import COUNTIES, city_key
def city_as_typed(c):
    c = c.strip().title()
    r = rng.random()
    if c.startswith("Saint "):
        c = rng.choice(["Saint ", "St. ", "St "]) + c[6:]
    elif c.startswith("St. "):
        c = rng.choice(["Saint ", "St. ", "St "]) + c[4:]
    return c.upper() if r < 0.3 else c
zip_cities = collections.defaultdict(set)
for z, rows in b.by_zip.items():
    for r in rows:
        zip_cities[z].add(r["city"].strip())
city_how = collections.Counter(); city_where = collections.Counter(); city_wrong = []
alt_how = collections.Counter(); alt_where = collections.Counter(); alt_wrong = []; alt_n = 0
rng2 = random.Random(7)
for z, r in sample:
    n = pick_number(r["low"], r["high"], r["parity"]) if r["low"].isdigit() and r["high"].isdigit() else None
    if n is None:
        continue
    pre, name, suf, post = r["raw"]
    by_zip_how, by_zip_code, _ = b.match(" ".join(x for x in (str(n), pre, name, suf, post) if x.strip()), z)
    if by_zip_how not in ("EXACT", "CLOSE"):
        continue
    m = str(n) + " " + messy(pre.strip(), name.strip(), suf.strip(), post.strip())
    county = COUNTIES.get(r["code"][6:9], "")
    how, code, note, where = b.match_ship_to(m, city_as_typed(r["city"]), county + " County")
    city_how[how] += 1; city_where[where] += 1
    if code and code != by_zip_code:
        city_wrong.append((m, r["city"], z, by_zip_code, how, code, where))
    # the same address under another mail city its zip also uses (a USPS alternate name)
    others = sorted(c for c in zip_cities[z] if city_key(c) != city_key(r["city"]))
    if others:
        alt_n += 1
        how, code, note, where = b.match_ship_to(m, rng2.choice(others), county)
        alt_how[how] += 1; alt_where[where] += 1
        if code and code != by_zip_code:
            alt_wrong.append((m, r["city"], z, by_zip_code, how, code, where))
total = sum(city_how.values())
placed = sum(v for k, v in city_how.items() if k in ("EXACT", "CLOSE", "STREET"))
print(f"  city as typed: {dict(city_how)} of {total}; found under {dict(city_where)}")
print(f"  placed {placed} of {total} ({100 * placed / total:.1f}%), wrong code: {len(city_wrong)}")
for x in city_wrong[:10]: print("   ", x)
check("no zip: wrong code never given when placed by city", len(city_wrong), 0)
check("no zip: at least 97% placed by street and city", placed / total >= 0.97, True)
print(f"  another mail city of the same zip ({alt_n} tried): {dict(alt_how)}; found under {dict(alt_where)}; "
      f"wrong code: {len(alt_wrong)}")
for x in alt_wrong[:10]: print("   ", x)

print(f"\n{fails} failures")
