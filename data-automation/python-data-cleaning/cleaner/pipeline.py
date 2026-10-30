"""End-to-end cleaning pipeline: load -> clean -> validate -> de-duplicate -> report."""
from __future__ import annotations

import time
from dataclasses import dataclass, field

import pandas as pd

from . import rules


@dataclass
class Result:
    clean: pd.DataFrame
    rejected: pd.DataFrame
    duplicates: pd.DataFrame
    stats: dict = field(default_factory=dict)


def _add_issue(issues: pd.Series, mask: pd.Series, label: str) -> pd.Series:
    return issues.where(~mask, issues + label + "; ")


def clean_frame(raw: pd.DataFrame) -> Result:
    t0 = time.perf_counter()
    df = raw.copy()
    n_in = len(df)
    issues = pd.Series("", index=df.index)

    df["customer_id"] = rules.squash_spaces(df["customer_id"]).str.upper()
    df["full_name"] = rules.title_case(df["full_name"])
    df["city"] = rules.title_case(df["city"])

    df["country"], bad_country = rules.normalize_country(df["country"])
    issues = _add_issue(issues, bad_country, "unknown country")

    df["email"], email_fixed, bad_email = rules.normalize_email(df["email"])
    issues = _add_issue(issues, email_fixed & ~bad_email, "email auto-fixed")
    issues = _add_issue(issues, bad_email, "invalid email")
    df.loc[bad_email, "email"] = ""

    raw_phone = df["phone"].copy()
    df["phone"], bad_phone = rules.normalize_phone(df["phone"], df["country"])
    issues = _add_issue(issues, bad_phone, "invalid phone")

    df["signup_date"], bad_date = rules.parse_dates(df["signup_date"])
    issues = _add_issue(issues, bad_date, "invalid date")

    df["total_spend"], bad_amount = rules.parse_amount(df["total_spend"])
    issues = _add_issue(issues, bad_amount, "invalid amount")

    # required fields
    missing_name = df["full_name"].eq("")
    no_contact = df["email"].eq("") & df["phone"].eq("")
    issues = _add_issue(issues, missing_name, "missing name")
    issues = _add_issue(issues, no_contact, "no valid email or phone")
    df["issues"] = issues.str.rstrip("; ")

    reject_mask = missing_name | no_contact
    rejected = raw.loc[reject_mask].copy()
    rejected["reason"] = df.loc[reject_mask, "issues"]
    df = df.loc[~reject_mask]

    # 1) exact duplicates (after normalisation, ignoring the id column)
    content_cols = [c for c in df.columns if c not in ("customer_id", "issues")]
    exact_dup = df.duplicated(subset=content_cols + ["customer_id"], keep="first")

    # 2) same person entered twice: match on e-mail, otherwise on phone + name
    key = df["email"].where(df["email"].ne(""), "tel:" + df["phone"] + "|" + df["full_name"].str.lower())
    df = df.assign(_key=key, _order=range(len(df)))
    df = df.sort_values(["_key", "signup_date", "_order"], na_position="last")
    person_dup = df.duplicated(subset="_key", keep="first") & ~exact_dup.reindex(df.index)

    keeper = df.groupby("_key")["customer_id"].transform("first")
    dup_mask = exact_dup.reindex(df.index) | person_dup
    duplicates = df.loc[dup_mask, ["customer_id", "full_name", "email", "phone"]].copy()
    duplicates["kept_customer_id"] = keeper[dup_mask]
    duplicates["match_type"] = "exact"
    duplicates.loc[person_dup[dup_mask], "match_type"] = "same email"
    duplicates.loc[person_dup[dup_mask] & df.loc[dup_mask, "_key"].str.startswith("tel:"), "match_type"] = "same phone + name"

    clean = (df.loc[~dup_mask].sort_values("_order")
               .drop(columns=["_key", "_order"]).reset_index(drop=True))

    stats = {
        "rows_in": n_in,
        "rows_clean": len(clean),
        "rows_rejected": len(rejected),
        "duplicates_removed": int(dup_mask.sum()),
        "exact_duplicates": int(exact_dup.sum()),
        "same_person_duplicates": int(person_dup.sum()),
        "emails_auto_fixed": int((email_fixed & ~bad_email).sum()),
        "invalid_emails_blanked": int(bad_email.sum()),
        "invalid_phones_blanked": int(bad_phone.sum()),
        "phones_normalised": int(raw_phone.fillna("").ne("").sum() - bad_phone.sum()),
        "invalid_dates": int(bad_date.sum()),
        "unknown_countries": int(bad_country.sum()),
        "seconds": round(time.perf_counter() - t0, 2),
    }
    return Result(clean, rejected.reset_index(drop=True), duplicates.reset_index(drop=True), stats)


def to_markdown(stats: dict, before: pd.DataFrame, after: pd.DataFrame, n: int = 5) -> str:
    pct = lambda x: f"{x / stats['rows_in']:.1%}" if stats["rows_in"] else "0%"
    lines = [
        "# Data Cleaning Report", "",
        "| Metric | Count | Share |", "|---|---:|---:|",
        f"| Rows in | {stats['rows_in']:,} | 100% |",
        f"| Clean rows out | {stats['rows_clean']:,} | {pct(stats['rows_clean'])} |",
        f"| Duplicates removed | {stats['duplicates_removed']:,} | {pct(stats['duplicates_removed'])} |",
        f"| &nbsp;&nbsp;exact copies | {stats['exact_duplicates']:,} | |",
        f"| &nbsp;&nbsp;same person, different formatting | {stats['same_person_duplicates']:,} | |",
        f"| Rejected (no name / no contact) | {stats['rows_rejected']:,} | {pct(stats['rows_rejected'])} |",
        f"| E-mails auto-fixed | {stats['emails_auto_fixed']:,} | |",
        f"| Invalid e-mails blanked | {stats['invalid_emails_blanked']:,} | |",
        f"| Phones converted to E.164 | {stats['phones_normalised']:,} | |",
        f"| Invalid dates flagged | {stats['invalid_dates']:,} | |",
        f"| Unknown countries flagged | {stats['unknown_countries']:,} | |",
        "", f"Processing time: **{stats['seconds']} s**", "",
        "## Before (first rows)", "", before.head(n).to_markdown(index=False, disable_numparse=True), "",
        "## After (first rows)", "", after.head(n).drop(columns=["issues"]).to_markdown(index=False, disable_numparse=True), "",
    ]
    return "\n".join(lines)
