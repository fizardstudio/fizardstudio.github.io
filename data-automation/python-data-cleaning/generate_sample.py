"""Generate a deliberately messy customer dataset for demoing the cleaner.

The data is 100% synthetic (random names, fake domains), so it is safe to
publish. Typical real-world problems are injected on purpose:

* inconsistent casing and stray whitespace
* phone numbers in many formats
* dates in several formats (and a few impossible ones)
* malformed e-mails
* exact duplicates and "near" duplicates (same person, different formatting)
* amounts stored as text with currency symbols / thousands separators
* blank required fields

Usage:
    python generate_sample.py --rows 200000 --out sample/customers_raw.csv
"""
from __future__ import annotations

import argparse
import csv
import random
from datetime import date, timedelta
from pathlib import Path

FIRST = ["James", "Mary", "John", "Patricia", "Robert", "Jennifer", "Michael", "Linda",
         "William", "Elizabeth", "David", "Barbara", "Richard", "Susan", "Joseph", "Jessica",
         "Thomas", "Sarah", "Charles", "Karen", "Ahmad", "Siti", "Budi", "Dewi", "Rizky", "Putri"]
LAST = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis",
        "Rodriguez", "Martinez", "Wilson", "Anderson", "Taylor", "Thomas", "Moore",
        "Santoso", "Wijaya", "Pratama", "Saputra", "Hidayat"]
CITIES = [("New York", "US"), ("Los Angeles", "US"), ("Chicago", "US"), ("London", "GB"),
          ("Manchester", "GB"), ("Toronto", "CA"), ("Sydney", "AU"), ("Jakarta", "ID"),
          ("Bandung", "ID"), ("Singapore", "SG")]
COUNTRY_VARIANTS = {
    "US": ["US", "USA", "United States", "united states", "U.S."],
    "GB": ["GB", "UK", "United Kingdom", "England"],
    "CA": ["CA", "Canada", "canada"],
    "AU": ["AU", "Australia"],
    "ID": ["ID", "Indonesia", "indonesia", "INA"],
    "SG": ["SG", "Singapore"],
}
DIAL = {"US": "1", "CA": "1", "GB": "44", "AU": "61", "ID": "62", "SG": "65"}
DOMAINS = ["example.com", "example.org", "mail.example.net", "demo.test"]


def messy_case(s: str, rnd: random.Random) -> str:
    r = rnd.random()
    if r < 0.15:
        return s.upper()
    if r < 0.30:
        return s.lower()
    if r < 0.40:
        return f"  {s} "
    return s


def phone_variants(cc: str, national: str, rnd: random.Random) -> str:
    dial = DIAL[cc]
    a, b, c = national[:3], national[3:6], national[6:]
    opts = [
        f"+{dial} {a} {b} {c}",
        f"+{dial}{national}",
        f"({a}) {b}-{c}",
        f"{a}-{b}-{c}",
        f"{a}.{b}.{c}",
        f"00{dial}{national}",
        f"0{national}" if cc in ("GB", "AU", "ID") else national,
    ]
    return rnd.choice(opts)


def date_variants(d: date, rnd: random.Random) -> str:
    opts = [d.isoformat(), d.strftime("%d/%m/%Y"), d.strftime("%m-%d-%Y"),
            d.strftime("%d %b %Y"), d.strftime("%Y/%m/%d"), d.strftime("%B %d, %Y")]
    return rnd.choice(opts)


def amount_variants(v: float, rnd: random.Random) -> str:
    opts = [f"{v:.2f}", f"${v:,.2f}", f"{v:,.2f}", f"USD {v:.2f}", f"{v:.0f}", f" {v:.2f} "]
    return rnd.choice(opts)


def make_person(i: int, rnd: random.Random) -> dict:
    first, last = rnd.choice(FIRST), rnd.choice(LAST)
    city, cc = rnd.choice(CITIES)
    national = "".join(str(rnd.randint(0, 9)) for _ in range(10))
    national = str(rnd.randint(2, 9)) + national[1:]
    signup = date(2019, 1, 1) + timedelta(days=rnd.randint(0, 2400))
    email = f"{first}.{last}{i}@{rnd.choice(DOMAINS)}".lower()
    return {
        "first": first, "last": last, "email": email, "cc": cc, "national": national,
        "city": city, "signup": signup, "spend": round(rnd.uniform(5, 5000), 2),
    }


def render(p: dict, cust_id: str, rnd: random.Random) -> dict:
    email = p["email"]
    r = rnd.random()
    if r < 0.03:
        email = email.replace("@", " at ")          # malformed
    elif r < 0.05:
        email = email.replace(".", ",", 1)           # typo
    elif r < 0.06:
        email = email.split("@")[0] + "@"            # domain missing
    elif r < 0.15:
        email = email.upper()
    signup = date_variants(p["signup"], rnd)
    if rnd.random() < 0.005:
        signup = "31/02/2021"                         # impossible date
    row = {
        "customer_id": cust_id,
        "full_name": messy_case(f"{p['first']} {p['last']}", rnd),
        "email": email,
        "phone": phone_variants(p["cc"], p["national"], rnd),
        "city": messy_case(p["city"], rnd),
        "country": rnd.choice(COUNTRY_VARIANTS[p["cc"]]),
        "signup_date": signup,
        "total_spend": amount_variants(p["spend"], rnd),
    }
    if rnd.random() < 0.02:
        row[rnd.choice(["full_name", "email", "phone"])] = ""   # missing value
    return row


def generate(rows: int, seed: int = 42):
    rnd = random.Random(seed)
    people = []
    out = []
    i = 0
    while len(out) < rows:
        r = rnd.random()
        if people and r < 0.04:
            # exact duplicate of an earlier row
            out.append(dict(rnd.choice(out)))
            continue
        if people and r < 0.10:
            # near duplicate: same person re-entered with different formatting
            p = rnd.choice(people)
            out.append(render(p, f"C{rnd.randint(10**6, 10**7 - 1)}", rnd))
            continue
        i += 1
        p = make_person(i, rnd)
        people.append(p)
        out.append(render(p, f"C{1000000 + i}", rnd))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--rows", type=int, default=1000)
    ap.add_argument("--seed", type=int, default=42)
    ap.add_argument("--out", type=Path, default=Path("sample/customers_raw.csv"))
    a = ap.parse_args()
    a.out.parent.mkdir(parents=True, exist_ok=True)
    data = generate(a.rows, a.seed)
    with a.out.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(data[0].keys()))
        w.writeheader()
        w.writerows(data)
    print(f"wrote {len(data):,} rows -> {a.out}")


if __name__ == "__main__":
    main()
