import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from cleaner import rules  # noqa: E402
from cleaner.pipeline import clean_frame  # noqa: E402
from generate_sample import generate  # noqa: E402


def S(*v):
    return pd.Series(list(v))


def test_country_aliases():
    iso, unknown = rules.normalize_country(S("usa", " United Kingdom ", "INA", "Narnia", ""))
    assert iso.tolist() == ["US", "GB", "ID", "", ""]
    assert unknown.tolist() == [False, False, False, True, False]


def test_email_fix_and_validate():
    email, fixed, invalid = rules.normalize_email(S("A.B@Example.COM", "a at b.com", "x@", ""))
    assert email.tolist()[:2] == ["a.b@example.com", "a@b.com"]
    assert fixed.tolist() == [False, True, False, False]
    assert invalid.tolist() == [False, False, True, False]


def test_phone_e164():
    phones = S("(415) 555-0132", "+44 20 7946 0958", "0812 3456 7890", "0044 20 7946 0958",
               "6281234567890", "12", "")
    country = S("US", "GB", "ID", "GB", "ID", "US", "US")
    e164, invalid = rules.normalize_phone(phones, country)
    assert e164.tolist() == ["+14155550132", "+442079460958", "+6281234567890",
                             "+442079460958", "+6281234567890", "", ""]
    assert invalid.tolist() == [False, False, False, False, False, True, False]


def test_dates_multiple_formats():
    d, invalid = rules.parse_dates(S("2024-03-05", "05/03/2024", "03-05-2024", "5 Mar 2024",
                                     "March 05, 2024", "31/02/2021", ""))
    assert d.iloc[:5].dt.strftime("%Y-%m-%d").tolist() == ["2024-03-05"] * 5
    assert invalid.tolist() == [False] * 5 + [True, False]


def test_amounts():
    a, invalid = rules.parse_amount(S("$1,234.50", "USD 99", " 12 ", "n/a"))
    assert a.iloc[:3].tolist() == [1234.5, 99.0, 12.0]
    assert invalid.tolist() == [False, False, False, True]


def test_pipeline_removes_duplicates_and_rejects():
    raw = pd.DataFrame([
        ["C1", "jane doe", "JANE@EXAMPLE.COM", "(415) 555-0132", "sf", "USA", "2024-01-02", "$10"],
        ["C1", "jane doe", "JANE@EXAMPLE.COM", "(415) 555-0132", "sf", "USA", "2024-01-02", "$10"],
        ["C9", " Jane  Doe ", "jane@example.com", "+1 415 555 0132", "SF", "US", "01-05-2024", "10"],
        ["C2", "", "nobody@example.com", "", "x", "US", "2024-01-01", "1"],
        ["C3", "No Contact", "bad@", "", "x", "US", "2024-01-01", "1"],
    ], columns=["customer_id", "full_name", "email", "phone", "city", "country",
                "signup_date", "total_spend"])
    res = clean_frame(raw)
    assert res.clean["customer_id"].tolist() == ["C1"]
    assert res.clean.loc[0, "full_name"] == "Jane Doe"
    assert res.stats["exact_duplicates"] == 1
    assert res.stats["same_person_duplicates"] == 1
    assert set(res.rejected["customer_id"]) == {"C2", "C3"}
    assert set(res.duplicates["kept_customer_id"]) == {"C1"}


def test_row_accounting_on_generated_data():
    raw = pd.DataFrame(generate(5000, seed=7))
    res = clean_frame(raw)
    s = res.stats
    assert s["rows_clean"] + s["rows_rejected"] + s["duplicates_removed"] == s["rows_in"]
    assert res.clean["email"].replace("", pd.NA).dropna().is_unique
