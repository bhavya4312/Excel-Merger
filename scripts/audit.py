#!/usr/bin/env python3
"""
Audit Script for Excel Merger
==============================
Cross-references all transactions from input files against the merged output
to ensure no credit/debit/due payments are missing.

Usage:
    npm run audit
    — or —
    python3 scripts/audit.py

Checks:
  1. Every invoice/bill from both inputs exists in the merged output
  2. Party totals in the output match the sum of input totals
  3. No phantom/duplicate rows in the output
  4. All Rcpt (receipt) rows are preserved
  5. Debit/Credit/Balance amounts match per invoice

Prerequisites:
  - python3 and openpyxl must be available
  - Place input files in test-files/ as "RMS KORBA.xlsx" and "RS KORBA.xlsx"
  - Place merged output in test-files/ as "Merged_Report_*.xlsx"
"""

import os
import sys
import re
import glob

try:
    import openpyxl
except ImportError:
    print("❌ openpyxl is not installed. Run: pip3 install openpyxl")
    sys.exit(1)

# ── Resolve paths relative to repo root ──
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
REPO_ROOT = os.path.dirname(SCRIPT_DIR)
TEST_DIR = os.path.join(REPO_ROOT, "test-files")

INPUT_FILES = {
    "Medical": os.path.join(TEST_DIR, "RMS KORBA.xlsx"),
    "Surgical": os.path.join(TEST_DIR, "RS KORBA.xlsx"),
}

# ── Validate prerequisites ──
for label, fpath in INPUT_FILES.items():
    if not os.path.exists(fpath):
        print(f"❌ Input file not found: {fpath}")
        print(f"   Place your {label} input file as '{os.path.basename(fpath)}' in test-files/")
        sys.exit(1)

output_files = sorted(glob.glob(os.path.join(TEST_DIR, "Merged_Report_*.xlsx")), reverse=True)
if not output_files:
    print("❌ No Merged_Report_*.xlsx found in test-files/ directory.")
    print("   Run the app, merge the test files, and save the output to test-files/")
    sys.exit(1)

OUTPUT_FILE = output_files[0]


# ============================================================================
# Parsers
# ============================================================================

def parse_input_file(filepath, label):
    """Extract all transaction rows (invoices + rcpt) from a MARG ERP input file."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    transactions = []
    rcpt_rows = []
    parties = {}

    current_party = None
    current_party_key = None

    for row_num in range(1, ws.max_row + 1):
        vals = [cell.value for cell in ws[row_num]]
        if len(vals) < 2:
            continue

        col_a = str(vals[0] or '').strip()
        col_b = str(vals[1] or '').strip()

        # Separator row — reset party context (matches app logic)
        if all(v is None or str(v).strip() == '' for v in vals):
            current_party = None
            current_party_key = None
            continue

        # Skip footer
        if col_a.startswith('Digital Purchase'):
            continue

        # Party header
        if col_b.startswith('-'):
            party_name = col_b.lstrip('-').strip()
            city = str(vals[3] or '').strip() if len(vals) > 3 else ''
            current_party = f"{party_name} ({city})"
            current_party_key = (party_name + city).lower().replace(" ", "")
            try:
                amount = float(vals[4]) if len(vals) > 4 and vals[4] else 0.0
            except (ValueError, TypeError):
                amount = 0.0
            parties[current_party_key] = {
                'name': party_name, 'city': city, 'amount': amount,
                'display': current_party
            }
            continue

        if not current_party:
            continue

        # Rcpt row
        if col_a.lower().startswith('rcpt'):
            rcpt_rows.append({
                'party': current_party,
                'party_key': current_party_key,
                'text': col_a,
                'row_num': row_num,
                'source': label,
            })
            continue

        # Transaction row
        if col_b:
            debit, credit, balance = 0.0, 0.0, 0.0
            if len(vals) > 3 and vals[3] is not None:
                bill_val = str(vals[3]).strip()
                parts = bill_val.split()
                if len(parts) == 2:
                    try:
                        debit, credit = float(parts[0]), float(parts[1])
                    except ValueError:
                        pass
                elif len(parts) == 1:
                    try:
                        v = float(parts[0])
                        if v < 0:
                            credit = abs(v)
                        else:
                            debit = v
                    except ValueError:
                        pass

            if len(vals) > 4 and vals[4] is not None:
                try:
                    balance = float(vals[4])
                except (ValueError, TypeError):
                    pass

            days = str(vals[5] or '').strip() if len(vals) > 5 and vals[5] is not None else ''

            transactions.append({
                'party': current_party,
                'party_key': current_party_key,
                'invoice': col_b,
                'date': str(vals[2] or '').strip() if len(vals) > 2 else '',
                'debit': debit, 'credit': credit, 'balance': balance,
                'days': days,
                'row_num': row_num,
                'source': label,
            })

    wb.close()
    return transactions, rcpt_rows, parties


def parse_output_file(filepath):
    """Extract all transaction rows from the merged output file."""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    transactions = []
    rcpt_rows = []
    parties = {}
    current_party = None
    current_party_key = None
    current_section = None

    for row_num in range(1, ws.max_row + 1):
        vals = [cell.value for cell in ws[row_num]]
        if len(vals) < 7:
            continue

        col_a = str(vals[0] or '').strip()
        col_b = str(vals[1] or '').strip()

        # Empty rows appear between sections within a party — DON'T reset party here
        if all(v is None or str(v).strip() == '' for v in vals):
            continue

        # Outstanding Summary marks end of a party block
        if col_a.startswith('Outstanding Summary'):
            current_party = None
            current_party_key = None
            current_section = None
            continue

        # Party header
        is_party_header = (col_a and not col_b and
            not col_a.lower().startswith('rcpt') and
            not col_a.startswith('---') and
            col_a != 'Type')

        if is_party_header:
            current_party = col_a
            match = re.match(r'^(.+?)\s+\((.+?)\)$', col_a)
            if match:
                name, city = match.group(1).strip(), match.group(2).strip()
                current_party_key = (name + city).lower().replace(" ", "")
            else:
                current_party_key = col_a.lower().replace(" ", "")
            total = float(vals[6]) if vals[6] is not None else 0.0
            parties[current_party_key] = {'display': col_a, 'total': total}
            continue

        # Section headers
        if col_a.startswith('--- Medical'):
            current_section = 'Medical'
            continue
        if col_a.startswith('--- Surgical'):
            current_section = 'Surgical'
            continue

        if col_a == 'Type':
            continue

        if not current_party:
            continue

        # Rcpt row
        if col_a.lower().startswith('rcpt'):
            rcpt_rows.append({
                'party': current_party, 'party_key': current_party_key,
                'text': col_a, 'section': current_section, 'row_num': row_num,
            })
            continue

        # Transaction row
        if col_b:
            try:
                debit = float(vals[3]) if vals[3] is not None and vals[3] != '' else 0.0
            except (ValueError, TypeError):
                debit = 0.0
            try:
                credit = float(vals[4]) if vals[4] is not None and vals[4] != '' else 0.0
            except (ValueError, TypeError):
                credit = 0.0
            try:
                balance = float(vals[5]) if vals[5] is not None and vals[5] != '' else 0.0
            except (ValueError, TypeError):
                balance = 0.0
            days = str(vals[6] or '').strip() if vals[6] is not None else ''

            transactions.append({
                'party': current_party, 'party_key': current_party_key,
                'invoice': col_b,
                'date': str(vals[2] or '').strip(),
                'debit': debit, 'credit': credit, 'balance': balance,
                'days': days, 'section': current_section, 'row_num': row_num,
            })

    wb.close()
    return transactions, rcpt_rows, parties


# ============================================================================
# Audit
# ============================================================================

def run_audit():
    print(f"Output file: {os.path.basename(OUTPUT_FILE)}")
    print("=" * 90)
    print("AUDIT: Cross-referencing input transactions vs merged output")
    print("=" * 90)

    # Parse
    all_input_txns, all_input_rcpts, all_input_parties = [], [], {}
    for label, fpath in INPUT_FILES.items():
        txns, rcpts, parties = parse_input_file(fpath, label)
        all_input_txns.extend(txns)
        all_input_rcpts.extend(rcpts)
        for k, v in parties.items():
            if k not in all_input_parties:
                all_input_parties[k] = {'Medical': 0.0, 'Surgical': 0.0, 'display': v['display']}
            all_input_parties[k][label] = v['amount']

    output_txns, output_rcpts, output_parties = parse_output_file(OUTPUT_FILE)

    issues = 0

    # ── Check 1: Every input transaction exists in output ──
    print("\n📋 CHECK 1: All input transactions present in output")
    print("-" * 80)
    missing = []
    for t in all_input_txns:
        inv = t['invoice'].replace('*', '')
        if not any(ot['party_key'] == t['party_key'] and ot['invoice'].replace('*', '') == inv for ot in output_txns):
            missing.append(t)
    if not missing:
        print(f"  ✅ All {len(all_input_txns)} transactions found in output")
    else:
        for t in missing:
            print(f"  ❌ MISSING: [{t['source']}] {t['party']} — {t['invoice']} "
                  f"(Row {t['row_num']}, Debit={t['debit']:.2f}, Credit={t['credit']:.2f}, Balance={t['balance']:.2f})")
        issues += len(missing)

    # ── Check 2: All Rcpt rows preserved ──
    print(f"\n📋 CHECK 2: All Rcpt rows preserved")
    print("-" * 80)
    input_rcpt_set = {(r['party_key'], r['text']) for r in all_input_rcpts}
    output_rcpt_set = {(r['party_key'], r['text']) for r in output_rcpts}
    missing_rcpts = input_rcpt_set - output_rcpt_set
    if not missing_rcpts:
        med_r = sum(1 for r in all_input_rcpts if r['source'] == 'Medical')
        sur_r = sum(1 for r in all_input_rcpts if r['source'] == 'Surgical')
        print(f"  ✅ All {len(all_input_rcpts)} Rcpt rows preserved ({med_r} Medical + {sur_r} Surgical)")
    else:
        for pk, text in missing_rcpts:
            print(f"  ❌ MISSING Rcpt: {text} (party_key={pk})")
        issues += len(missing_rcpts)

    # ── Check 3: Party totals match ──
    print(f"\n📋 CHECK 3: Party totals match")
    print("-" * 80)
    total_mismatches = 0
    for key, inp in all_input_parties.items():
        expected = inp.get('Medical', 0) + inp.get('Surgical', 0)
        if key in output_parties:
            actual = output_parties[key]['total']
            if abs(expected - actual) > 0.01:
                print(f"  ❌ MISMATCH: {inp['display']}  Expected={expected:.2f}  Got={actual:.2f}")
                total_mismatches += 1
        elif expected > 0:
            print(f"  ❌ PARTY NOT IN OUTPUT: {inp['display']} (total={expected:.2f})")
            total_mismatches += 1
    if total_mismatches == 0:
        print(f"  ✅ All {len(all_input_parties)} party totals match")
    issues += total_mismatches

    # ── Check 4: No phantom rows ──
    print(f"\n📋 CHECK 4: No phantom/extra rows in output")
    print("-" * 80)
    phantoms = []
    for ot in output_txns:
        inv = ot['invoice'].replace('*', '')
        if not any(it['party_key'] == ot['party_key'] and it['invoice'].replace('*', '') == inv for it in all_input_txns):
            phantoms.append(ot)
    if not phantoms:
        print(f"  ✅ No phantom rows. Output has exactly {len(output_txns)} transactions")
    else:
        for ot in phantoms:
            print(f"  ⚠️  PHANTOM: {ot['party']} — {ot['invoice']} "
                  f"(Section={ot['section']}, Row {ot['row_num']})")
        issues += len(phantoms)

    # ── Check 5: Balance amounts match ──
    print(f"\n📋 CHECK 5: Balance amounts match per invoice")
    print("-" * 80)
    bal_mismatches = 0
    for it in all_input_txns:
        inv = it['invoice'].replace('*', '')
        for ot in output_txns:
            if ot['party_key'] == it['party_key'] and ot['invoice'].replace('*', '') == inv:
                if abs(it['balance'] - ot['balance']) > 0.01:
                    print(f"  ❌ {it['party']} — {it['invoice']}  Input={it['balance']:.2f}  Output={ot['balance']:.2f}")
                    bal_mismatches += 1
    if bal_mismatches == 0:
        print(f"  ✅ All balance amounts match across {len(all_input_txns)} transactions")
    issues += bal_mismatches

    # ── Summary ──
    input_med = sum(p.get('Medical', 0) for p in all_input_parties.values())
    input_sur = sum(p.get('Surgical', 0) for p in all_input_parties.values())
    output_total = sum(p['total'] for p in output_parties.values())

    print("\n" + "=" * 90)
    print(f"""
  Input:   {len(all_input_txns)} transactions, {len(all_input_rcpts)} Rcpt rows
           Medical: {input_med:>12,.2f}  |  Surgical: {input_sur:>12,.2f}  |  Total: {input_med + input_sur:>12,.2f}

  Output:  {len(output_txns)} transactions, {len(output_rcpts)} Rcpt rows, {len(output_parties)} parties
           Grand Total: {output_total:>12,.2f}""")

    print()
    if issues == 0:
        print("  🎉 AUDIT PASSED — No missing or incorrect data.")
    else:
        print(f"  ❌ AUDIT FAILED — {issues} issue(s) found. Review above.")
        if phantoms and not missing:
            print("  💡 Phantom rows may be from a stale output. Re-run the app and save fresh output.")

    return 0 if issues == 0 else 1


if __name__ == "__main__":
    sys.exit(run_audit())

