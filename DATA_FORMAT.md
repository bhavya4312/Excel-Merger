# Data Format Reference

This document describes the exact structure of the input Excel files and the merged output, for use as a reference when making modifications.

## Input Files

Both input files follow the **MARG ERP 9+ Excel Report** format. They are outstanding (receivable) reports exported from MARG ERP accounting software.

| File | Entity | Example |
|---|---|---|
| Medical file | RAHUL MEDICAL & SURGICALS | `RMS KORBA.xlsx` |
| Surgical file | RAHUL SURGICAL | `RS KORBA.xlsx` |

### Column Layout (6 columns: A–F)

| Column | Header (Row 7) | Content |
|---|---|---|
| A | *(none)* | Misc flags (e.g., `2`), or `Rcpt:` lines |
| B | INVOICE | Invoice/bill number (e.g., `*RM-4498`, `RS-0093`) |
| C | DATE | Invoice date (e.g., `06-03-26`) |
| D | BILL VALUE | **Debit and Credit as a single space-separated string** (e.g., `"2920.00          0.00"`) |
| E | BALANCE | Outstanding balance (numeric) |
| F | O/D | Overdue days (numeric) |

### Row Types

The file is structured top-to-bottom as follows:

#### 1. Header Block (Rows 1–7)

Rows 1–6 contain business metadata (name, address, GSTIN, date). Row 7 is the column header row.

```
Row 1: ["                           RAHUL MEDICAL & SURGICALS", None, None, None, None, None]
Row 2: ["                MEDICAL COMPLEX, TELIPARA BILASPUR (C.G) 495001", ...]
Row 3: ["Phone : 9329428767 E-Mail : rahulmedicalbilaspur@gmail.com", ...]
Row 4: ["M.S.M.E. NO : ...", ...]
Row 5: ["             GSTIN : 22AJNPV2350G1Z7 Food Lic.No. : ...", ...]
Row 6: ["                         OUTSTANDING AS ON 16-04-2026", ...]
Row 7: [None, "INVOICE", "DATE", "BILL VALUE", "BALANCE", "O/D"]
```

#### 2. Grand Total (Row 9)

```
Row 9: [None, "TOTAL NO. OF", "BILLS : 194", "GRAND TOTAL :", 634391, ""]
```

#### 3. City Subtotal

Groups parties by city. Appears before the first party of that city.

```
[None, "DIPKA", "", "", 125495, ""]
```

**Pattern:** Column B = city name, Column E = city subtotal amount.

#### 4. Party Header

Identifies a customer/party. Column B starts with `-`.

```
[None, "-DR.ARVIND MAHANT", "", "DIPKA", 3892, ""]
[None, "-GEETA MEDICAL", "STORES", "DIPKA", 917, ""]
```

**Pattern:**
- Column B: `-` + party name (may continue into Column C, e.g., `"STORES"` is part of the name)
- Column D: City name
- Column E: Party outstanding total

> **Matching logic:** Parties across the two files are matched using a **normalized key** = `(party_name + city).lower().replace(" ", "")`. For example, `"GEETA MEDICAL STORES"` in `DIPKA` appears in both files and gets merged.

#### 5. Receipt Note (Optional)

Appears immediately after a party header if there's a recent receipt.

```
["Rcpt: 799.00 Dt.: 05-04-26 <11 Days>", None, None, None, None, None]
```

**Pattern:** Column A starts with `Rcpt:`. Columns B–F are all `None`. These rows have no debit/credit data — the receipt info is entirely in the Column A text.

#### 6. Transaction Rows

Individual invoice/bill entries under a party.

```
[2, "*RM-4498", "06-03-26", "2920.00          0.00", 2920, 41]
[None, "*RM-4680", "18-03-26", "155.00          0.00", 155, 29]
```

**Key detail — Column D (BILL VALUE):**
The debit and credit amounts are packed into a single string separated by whitespace, e.g., `"2920.00          0.00"`. The Python processing script splits this into separate Debit and Credit columns in the output.

**Negative debit values:**
Some rows have a negative debit value formatted as a single number, e.g., `"       -209.00"` (1 part, not 2). When this happens, the processing script flips it: sets Debit to `0.0` and puts the absolute value in the Credit column.

- Asterisk (`*`) prefix on invoice numbers (e.g., `*RM-4498`) indicates overdue bills.
- Column A sometimes contains `2` (meaning unclear — possibly a flag from MARG ERP). This value is stripped during processing.
- Column F is the number of overdue days.

#### 7. Separator Row

An all-`None` row appears between each party block.

```
[None, None, None, None, None, None]
```

#### 8. Footer (Last Row)

An ad/promo line from MARG ERP software. Filtered out during processing.

```
["Digital Purchase | ERP Ordering | Healthcare QRCode on bills for extra earnings | Call MARG 7771012366,...", ...]
```

### Parsing Edge Cases

> **Separator rows reset party context:**
> The parser resets `current_party_key = None` on separator rows (all-`None` rows). This prevents city subtotal rows (e.g., `[None, "KORBA", "", "", 508896, ""]`) and other non-transaction rows from leaking into the previous party's transaction list. Without this reset, any row appearing after a separator but before the next party header would be incorrectly attributed to the last party.

> **`Rcpt:` rows and debit processing:**
> Receipt rows have `None` in Column D and no bill value data. The parser correctly includes them in the output without attempting debit/credit conversion.

> **Footer filtering:**
> The MARG ERP footer/ad row (`"Digital Purchase | ERP Ordering..."`) is explicitly filtered out by checking if Column A starts with `"Digital Purchase"`.

---

## Merged Output Format

After processing, the app produces a 7-column merged report:

### Output Columns

| Column | Content |
|---|---|
| 1 (A) | Type (e.g., `Rcpt:...` lines) |
| 2 (B) | Bill/Invoice number |
| 3 (C) | Date |
| 4 (D) | Debit |
| 5 (E) | Credit |
| 6 (F) | Balance |
| 7 (G) | Days (overdue) |

> Note: The input's single `BILL VALUE` column (D) gets **split into two columns** — Debit (D) and Credit (E) — so the output has 7 columns vs the input's 6.

### Output Row Structure (per party)

Each matched party produces this block:

```
┌─────────────────────────────────────────────────────────────────┐
│  Party Name  (City)                              │ Total Amt   │  ← Blue fill, merged cols A–F
├─────────────────────────────────────────────────────────────────┤
│  Type │ Bill │ Date │ Debit │ Credit │ Balance │ Days          │  ← Grey sub-header
├─────────────────────────────────────────────────────────────────┤
│                    --- Medical ---                              │  ← Green fill (if has medical txns)
│  ...transaction rows from Medical file...                      │
├─────────────────────────────────────────────────────────────────┤
│                    --- Surgical ---                             │  ← Orange fill (if has surgical txns)
│  ...transaction rows from Surgical file...                     │
├─────────────────────────────────────────────────────────────────┤
│  Outstanding Summary | Medical: X | Surgical: Y | Total: Z    │  ← Dark blue fill, white text
├─────────────────────────────────────────────────────────────────┤
│  (empty row separator)                                         │
└─────────────────────────────────────────────────────────────────┘
```

### Color Coding (Excel)

| Section | Fill Color | Hex |
|---|---|---|
| Party header | Light blue | `#D9E1F2` |
| `--- Medical ---` section label | Light green | `#C6E0B4` |
| `--- Surgical ---` section label | Light orange | `#F8CBAD` |
| Column sub-headers (Type, Bill, etc.) | Light grey | `#F2F2F2` |
| Outstanding Summary | Dark blue (white text) | `#4472C4` |

### Number Formatting

- **Indian comma grouping** is applied to all monetary amounts (Debit, Credit, Balance, party totals): e.g., `12,34,567.00` instead of `1234567.00`.
- In Excel, the number format `#,##,##0.00` is applied to numeric cells in columns D (Debit), E (Credit), F (Balance), and G (party total).
- The Outstanding Summary text uses formatted numbers: `Outstanding Summary | Medical: 6,34,391.00 | Surgical: 5,08,100.00 | Total: 11,42,491.00`.

### Excel Output Details

- **Sheet name:** "Outstanding Report" (not default "Sheet1")
- **Output filename:** `Merged_Report_YYYY-MM-DD.xlsx` (includes the processing date)

### PDF Output

- The PDF mirrors the Excel layout and color scheme, rendered as a single continuous page using jsPDF + autoTable.
- **Title:** "Outstanding Report — DD Mon YYYY" displayed at the top of the page.
- The page height is dynamically calculated to fit all content without pagination.
- Numeric values in Debit, Credit, and Balance columns use Indian comma grouping.
- **Output filename:** `Merged_Report_YYYY-MM-DD.pdf`
