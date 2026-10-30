# Data Cleaning Report

| Metric | Count | Share |
|---|---:|---:|
| Rows in | 1,000 | 100% |
| Clean rows out | 892 | 89.2% |
| Duplicates removed | 97 | 9.7% |
| &nbsp;&nbsp;exact copies | 39 | |
| &nbsp;&nbsp;same person, different formatting | 58 | |
| Rejected (no name / no contact) | 11 | 1.1% |
| E-mails auto-fixed | 47 | |
| Invalid e-mails blanked | 10 | |
| Phones converted to E.164 | 990 | |
| Invalid dates flagged | 4 | |
| Unknown countries flagged | 0 | |

Processing time: **0.06 s**

## Before (first rows)

| customer_id   | full_name         | email                               | phone          | city        | country   | signup_date       | total_spend   |
|:--------------|:------------------|:------------------------------------|:---------------|:------------|:----------|:------------------|:--------------|
| C1000001      | james rodriguez   | james.rodriguez1@example.com        | 521.819.6001   | london      | GB        | August 10, 2021   | $2,808.42     |
| C1000002      | Robert Miller     | robert.miller2@example.com          | +13161559407   | Toronto     | CA        | 2023/03/31        | 2762.44       |
| C1000003      | john davis        | john.davis3@example.org             | (447) 525-5341 | Los Angeles | US        | December 27, 2024 | $821.20       |
| C1000004      | richard rodriguez | richard.rodriguez4@mail.example.net | +15953767242   | Los Angeles | US        | 2025/04/18        | 3736.33       |
| C1000005      | Jennifer Thomas   | jennifer.thomas5@example.com        | (666) 978-4801 | Singapore   | SG        | 2022-10-25        | 1471          |

## After (first rows)

| customer_id   | full_name         | email                               | phone         | city        | country   | signup_date   | total_spend   |
|:--------------|:------------------|:------------------------------------|:--------------|:------------|:----------|:--------------|:--------------|
| C1000001      | James Rodriguez   | james.rodriguez1@example.com        | +445218196001 | London      | GB        | 2021-08-10    | 2808.42       |
| C1000002      | Robert Miller     | robert.miller2@example.com          | +13161559407  | Toronto     | CA        | 2023-03-31    | 2762.44       |
| C1000003      | John Davis        | john.davis3@example.org             | +14475255341  | Los Angeles | US        | 2024-12-27    | 821.20        |
| C1000004      | Richard Rodriguez | richard.rodriguez4@mail.example.net | +15953767242  | Los Angeles | US        | 2025-04-18    | 3736.33       |
| C1000005      | Jennifer Thomas   | jennifer.thomas5@example.com        | +656669784801 | Singapore   | SG        | 2022-10-25    | 1471.00       |
