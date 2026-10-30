"""Column-level cleaning rules. Every function is vectorised (pandas Series in,
Series out) so the pipeline stays fast on files with hundreds of thousands
of rows."""
from __future__ import annotations

import pandas as pd

COUNTRY_MAP = {
    "us": "US", "usa": "US", "u.s.": "US", "u.s.a.": "US", "united states": "US",
    "united states of america": "US",
    "gb": "GB", "uk": "GB", "u.k.": "GB", "united kingdom": "GB", "england": "GB",
    "great britain": "GB",
    "ca": "CA", "canada": "CA",
    "au": "AU", "australia": "AU",
    "id": "ID", "ina": "ID", "indonesia": "ID",
    "sg": "SG", "singapore": "SG",
}
DIAL_CODES = {"US": "1", "CA": "1", "GB": "44", "AU": "61", "ID": "62", "SG": "65"}

# Order matters: the first format that parses wins. Slash dates are read as
# day/month (international style) and dash dates as month-day (US export
# style) - change this list to match the client's source system.
DATE_FORMATS = ["%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m-%d-%Y", "%d %b %Y", "%B %d, %Y"]

EMAIL_RE = r"^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$"


def squash_spaces(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.replace(r"\s+", " ", regex=True)


def title_case(s: pd.Series) -> pd.Series:
    return squash_spaces(s).str.title()


def normalize_country(s: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Return (iso2_code, is_unknown)."""
    key = squash_spaces(s).str.lower()
    iso = key.map(COUNTRY_MAP)
    unknown = iso.isna() & key.ne("")
    return iso.fillna(""), unknown


def normalize_email(s: pd.Series) -> tuple[pd.Series, pd.Series, pd.Series]:
    """Return (email, was_autofixed, is_invalid).

    Only two very safe repairs are made automatically (" at " -> "@" and a
    comma typed instead of a dot); everything else that does not look like an
    address is flagged, never guessed."""
    raw = squash_spaces(s).str.lower()
    fixed = raw.str.replace(r"\s+at\s+", "@", regex=True).str.replace(",", ".", regex=False)
    autofixed = fixed.ne(raw)
    invalid = fixed.ne("") & ~fixed.str.match(EMAIL_RE)
    return fixed, autofixed, invalid


def normalize_phone(s: pd.Series, country: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Convert phone numbers to E.164 (+<country code><number>).

    Returns (e164, is_invalid)."""
    raw = squash_spaces(s)
    has_plus = raw.str.startswith("+")
    digits = raw.str.replace(r"\D", "", regex=True)
    dial = country.map(DIAL_CODES).fillna("")

    intl_00 = ~has_plus & digits.str.startswith("00")
    digits = digits.where(~intl_00, digits.str[2:])
    international = has_plus | intl_00

    # national number: drop a single trunk "0" and prepend the country code
    national = digits.str.replace(r"^0", "", regex=True)
    starts_with_dial = pd.Series([bool(d) and n.startswith(d) for n, d in zip(national, dial)],
                                 index=national.index)
    already_prefixed = starts_with_dial & (national.str.len() - dial.str.len() >= 9)
    local = dial + national
    e164_digits = digits.where(international, local.where(~already_prefixed, national))

    e164 = "+" + e164_digits
    invalid = digits.ne("") & ~e164_digits.str.len().between(8, 15)
    invalid |= digits.ne("") & ~international & dial.eq("")
    e164 = e164.where(digits.ne("") & ~invalid, "")
    return e164, invalid & digits.ne("")


def parse_dates(s: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Return (datetime64 series, is_invalid)."""
    raw = squash_spaces(s)
    out = pd.Series(pd.NaT, index=s.index, dtype="datetime64[ns]")
    for fmt in DATE_FORMATS:
        todo = out.isna() & raw.ne("")
        if not todo.any():
            break
        out.loc[todo] = pd.to_datetime(raw[todo], format=fmt, errors="coerce")
    invalid = out.isna() & raw.ne("")
    return out, invalid


def parse_amount(s: pd.Series) -> tuple[pd.Series, pd.Series]:
    """'$1,234.50', 'USD 99', ' 12.00 ' -> float. Return (amount, is_invalid)."""
    raw = squash_spaces(s)
    cleaned = raw.str.replace(r"[^0-9.\-]", "", regex=True)
    num = pd.to_numeric(cleaned, errors="coerce")
    invalid = num.isna() & raw.ne("")
    return num.round(2), invalid
