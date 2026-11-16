"""CRM contact cleanup pipeline: messy export -> CRM-ready, de-duplicated list.

Steps
  1. Trim whitespace, drop junk/test rows
  2. Standardize names (fix "Last, First"), titles, companies, e-mails, phones (E.164),
     states (USPS 2-letter), ZIP codes (restore leading zeros), lead sources and dates (ISO)
  3. Flag invalid / missing e-mails (never guessed or invented)
  4. De-duplicate with matching keys (e-mail, phone, mobile, name + company),
     keep the most complete record and fill its gaps from the duplicates
  5. Write the clean CSV + an Excel delivery workbook with QA log and summary

Usage:  python crm_cleanup_pipeline.py <raw_csv> <output_dir>
"""
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

STATES = {"alabama": "AL", "alaska": "AK", "arizona": "AZ", "ariz.": "AZ", "arkansas": "AR",
          "california": "CA", "calif.": "CA", "cali": "CA", "colorado": "CO", "colo.": "CO",
          "connecticut": "CT", "conn.": "CT", "delaware": "DE", "florida": "FL", "fla.": "FL",
          "georgia": "GA", "ga.": "GA", "illinois": "IL", "ill.": "IL", "massachusetts": "MA",
          "mass.": "MA", "new jersey": "NJ", "n.j.": "NJ", "new york": "NY", "n.y.": "NY",
          "oregon": "OR", "ore.": "OR", "texas": "TX", "tex.": "TX", "washington": "WA",
          "wash.": "WA"}
SOURCE_MAP = {"website": "Website Form", "web form": "Website Form", "website form": "Website Form",
              "referral": "Referral", "referal": "Referral", "ref.": "Referral",
              "linkedin": "LinkedIn", "linked in": "LinkedIn", "trade show": "Trade Show",
              "tradeshow": "Trade Show", "expo": "Trade Show", "cold call": "Cold Outreach",
              "outbound": "Cold Outreach", "cold outreach": "Cold Outreach", "google": "Google Ads",
              "adwords": "Google Ads", "ppc": "Google Ads", "google ads": "Google Ads"}
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%Y", "%B %d, %Y"]
EMAIL_RE = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9-]+(\.[a-z0-9-]+)*\.[a-z]{2,}$")
JUNK_RE = re.compile(r"\b(test|asdf|sample|do not use|delete me)\b|^\?+$", re.I)
ACRONYMS = {"vp": "VP", "ceo": "CEO", "coo": "COO", "cfo": "CFO", "llc": "LLC"}
SMALL = {"of", "and", "the"}


def clean_ws(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def smart_title(s):
    def fix(word):
        low = word.lower()
        if low in ACRONYMS:
            return ACRONYMS[low]
        if low in SMALL:
            return low
        parts = re.split(r"([-'])", low)
        out = "".join(p.capitalize() if p not in "-'" else p for p in parts)
        if out.startswith("Mc") and len(out) > 2:
            out = "Mc" + out[2].upper() + out[3:]
        return out
    words = [fix(w) for w in s.split(" ")]
    if words:
        words[0] = words[0][0].upper() + words[0][1:] if words[0] else words[0]
    return " ".join(words)


def norm_company(raw):
    s = clean_ws(raw)
    legal = ""
    m = re.search(r",?\s+(l\.?l\.?c\.?|inc\.?|incorporated)$", s, re.I)
    if m:
        legal = "LLC" if "l" in m.group(1).lower()[:1] else "Inc."
        s = s[:m.start()]
    base = smart_title(s)
    key = re.sub(r"[^a-z0-9]", "", base.lower())
    return base, legal, key


def parse_phone(raw):
    s = clean_ws(raw)
    if not s:
        return "", "", ""
    ext = ""
    m = re.search(r"\s*(?:ext\.?|x)\s*(\d+)$", s, re.I)
    if m:
        ext, s = m.group(1), s[:m.start()]
    d = re.sub(r"\D", "", s)
    if len(d) == 11 and d.startswith("1"):
        d = d[1:]
    if len(d) != 10 or d == "0" * 10 or d[0] in "01":
        return "", "", "INVALID"
    return f"+1{d}", f"({d[:3]}) {d[3:6]}-{d[6:]}", ext


def parse_date(raw):
    s = clean_ws(raw)
    for f in DATE_FORMATS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            pass
    return None


def run(raw_csv, out_dir):
    out_dir = Path(out_dir)
    raw = pd.read_csv(raw_csv, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    log, junk, recs = [], [], []

    def change(rec_no, field, before, after, rule):
        if str(before) != str(after):
            log.append({"Record #": rec_no, "Field": field, "Before": before, "After": after, "Rule": rule})

    for _, r in raw.iterrows():
        n = int(r["Record #"])
        name = clean_ws(r["Full Name"])
        email_raw = clean_ws(r["Email"])
        contact_bits = [name, email_raw, clean_ws(r["Phone"]), clean_ws(r["Mobile"])]
        if (not name and not email_raw) or JUNK_RE.search(name) or JUNK_RE.search(email_raw.split("@")[0]):
            junk.append({"Record #": n, "Full Name": r["Full Name"], "Email": r["Email"],
                         "Reason": "Empty record" if not any(contact_bits[:2]) else "Test / placeholder entry"})
            continue
        rec = {"Record #": n}
        # --- name
        reordered = "," in name
        if reordered:
            last, first = [p.strip() for p in name.split(",", 1)]
            fixed = f"{first} {last}"
            change(n, "Full Name", r["Full Name"], fixed, "Name order fixed (Last, First)")
            name = fixed
        tname = smart_title(name)
        if tname != name:
            change(n, "Full Name", name, tname, "Name capitalization standardized")
        elif not reordered and name != r["Full Name"]:
            change(n, "Full Name", r["Full Name"], name, "Extra spaces removed")
        parts = tname.split(" ")
        rec["First Name"], rec["Last Name"] = parts[0], " ".join(parts[1:])
        # --- title
        t = clean_ws(r["Job Title"])
        tt = smart_title(t) if t else ""
        change(n, "Job Title", r["Job Title"], tt,
               "Title capitalization standardized" if tt != t else "Extra spaces removed")
        rec["Job Title"] = tt
        # --- company
        base, legal, key = norm_company(r["Company"])
        rec["_co_key"], rec["_co_base"], rec["_co_legal"] = key, base, legal
        # --- email
        e = email_raw.lower().replace(" ", "")
        if e in ("n/a", "na", "none", "-"):
            e = ""
        if not e:
            status = "Missing"
        elif EMAIL_RE.match(e):
            status = "Valid"
        else:
            status = "Invalid - verify with client"
        if status == "Valid":
            change(n, "Email", r["Email"], e, "Email lowercased / trimmed")
            rec["Email"] = e
        else:
            rec["Email"] = "" if status == "Missing" else email_raw
        rec["Email Status"] = status
        # --- phones
        for fld, col in (("Phone", "Phone"), ("Mobile", "Mobile")):
            e164, disp, ext = parse_phone(r[col])
            if ext == "INVALID":
                rec[f"{fld} (E.164)"], rec[f"{fld}"], ext = "", "", ""
                rec.setdefault("_notes", []).append(f"{fld} '{clean_ws(r[col])}' unusable")
            else:
                rec[f"{fld} (E.164)"], rec[f"{fld}"] = e164, disp
                change(n, fld, r[col], disp + (f" x{ext}" if ext else ""), "Phone formatted (E.164 + display)")
            if fld == "Phone":
                rec["Ext"] = ext
        # --- location
        city = smart_title(clean_ws(r["City"])) if clean_ws(r["City"]) else ""
        change(n, "City", r["City"], city, "City capitalization standardized")
        rec["City"] = city
        st_raw = clean_ws(r["State"])
        st = STATES.get(st_raw.lower(), st_raw.upper() if len(st_raw) == 2 else st_raw)
        change(n, "State", r["State"], st, "State converted to USPS 2-letter code")
        rec["State"] = st
        z = re.sub(r"\D", "", clean_ws(r["ZIP"]).split("-")[0])
        z5 = z.zfill(5) if 3 <= len(z) <= 5 else ""
        if z5 and z5 != r["ZIP"]:
            change(n, "ZIP", r["ZIP"], z5, "ZIP leading zero restored" if len(z) < 5 else "ZIP+4 trimmed to 5 digits")
        rec["ZIP"] = z5
        rec["Country"] = "US"
        # --- source and date
        src_raw = clean_ws(r["Lead Source"])
        src = SOURCE_MAP.get(src_raw.lower(), src_raw)
        change(n, "Lead Source", r["Lead Source"], src, "Lead source mapped to picklist")
        rec["Lead Source"] = src
        d = parse_date(r["Created Date"])
        if d:
            change(n, "Created Date", r["Created Date"], d.isoformat(), "Date converted to ISO (YYYY-MM-DD)")
        rec["Created Date"] = d
        recs.append(rec)

    # --- company canonical names: most frequent spelling per key, keep legal suffix if any
    by_key = defaultdict(Counter)
    for rec in recs:
        by_key[rec["_co_key"]][(rec["_co_base"], rec["_co_legal"])] += 1
    canon = {}
    for k, cnt in by_key.items():
        base = cnt.most_common(1)[0][0][0]
        legals = Counter(lg for (_, lg), c in cnt.items() for _ in range(c) if lg)
        canon[k] = f"{base} {legals.most_common(1)[0][0]}" if legals else base
    raw_by_no = raw.set_index("Record #")
    for rec in recs:
        rec["Company"] = canon[rec["_co_key"]]
        change(rec["Record #"], "Company", raw_by_no.loc[str(rec["Record #"]), "Company"], rec["Company"],
               "Company name standardized")

    # --- de-duplication (union-find over matching keys)
    parent = list(range(len(recs)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i
    reasons = defaultdict(set)
    seen = {}
    for i, rec in enumerate(recs):
        keys = []
        if rec["Email Status"] == "Valid":
            keys.append(("email", rec["Email"]))
        for p in ("Phone (E.164)", "Mobile (E.164)"):
            if rec[p]:
                keys.append(("phone", rec[p]))
        keys.append(("name+company", (rec["First Name"].lower(), rec["Last Name"].lower(), rec["_co_key"])))
        for k in keys:
            if k in seen:
                a, b = find(seen[k]), find(i)
                if a != b:
                    parent[b] = a
                reasons[i].add(k[0])
            else:
                seen[k] = i
    groups = defaultdict(list)
    for i in range(len(recs)):
        groups[find(i)].append(i)

    fields = ["First Name", "Last Name", "Job Title", "Company", "Email", "Email Status", "Phone",
              "Phone (E.164)", "Ext", "Mobile", "Mobile (E.164)", "City", "State", "ZIP", "Country",
              "Lead Source", "Created Date"]
    status_rank = {"Valid": 2, "Invalid - verify with client": 1, "Missing": 0}
    clean, dups = [], []
    for gi, (root, idx) in enumerate(sorted(groups.items(), key=lambda kv: min(recs[i]["Record #"] for i in kv[1]))):
        members = [recs[i] for i in idx]
        members.sort(key=lambda m: (-sum(1 for f in fields if m.get(f)), m["Record #"]))
        keep = dict(members[0])
        for other in members[1:]:
            for f in fields:
                if f == "Email Status":
                    continue
                if not keep.get(f) and other.get(f):
                    keep[f] = other[f]
            if status_rank[other["Email Status"]] > status_rank[keep["Email Status"]]:
                keep["Email"], keep["Email Status"] = other["Email"], other["Email Status"]
        dates = [m["Created Date"] for m in members if m["Created Date"]]
        keep["Created Date"] = min(dates) if dates else None
        cid = f"C-{gi + 1:04d}"
        notes = list(dict.fromkeys(n for m in members for n in m.get("_notes", [])))
        if len(members) > 1:
            notes.insert(0, "Merged records #" + ", #".join(str(m["Record #"]) for m in members))
            why = sorted(set().union(*(reasons[i] for i in idx)))
            for m in members[1:]:
                dups.append({"Kept Record #": members[0]["Record #"], "Merged Record #": m["Record #"],
                             "Contact ID": cid, "Name": f"{m['First Name']} {m['Last Name']}",
                             "Company": keep["Company"], "Matched On": " + ".join(why)})
        if keep["Email Status"] != "Valid":
            notes.append("Email needs verification" if keep["Email Status"].startswith("Invalid") else "No email on file")
        if not keep["Phone"] and not keep["Mobile"]:
            notes.append("No phone on file")
        row = {"Contact ID": cid, **{f: keep.get(f, "") for f in fields},
               "Source Records": ", ".join(str(m["Record #"]) for m in sorted(members, key=lambda m: m["Record #"])),
               "Notes": "; ".join(notes)}
        clean.append(row)

    clean_df = pd.DataFrame(clean)
    log_df = pd.DataFrame(log)
    junk_df = pd.DataFrame(junk)
    dup_df = pd.DataFrame(dups)
    csv_df = clean_df.copy()
    csv_df["Created Date"] = csv_df["Created Date"].map(lambda d: d.isoformat() if d else "")
    csv_df.to_csv(out_dir / "crm_ready_contacts.csv", index=False, encoding="utf-8-sig")
    build_workbook(raw, clean_df, log_df, junk_df, dup_df, out_dir / "CRM_Contact_Cleanup_Delivery.xlsx")
    print(f"raw={len(raw)} junk={len(junk_df)} merged_rows={len(dup_df)} clean={len(clean_df)} "
          f"changes={len(log_df)}")
    return clean_df, dup_df, junk_df, log_df


# ---------------------------------------------------------------- workbook
NAVY, TEAL, LIGHT = "1F3A5F", "2A9D8F", "EAF4F4"
HDR_FILL = PatternFill("solid", fgColor=NAVY)
HDR_FONT = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D0D7DE")


def write_table(ws, df, name, style="TableStyleMedium2", widths=None, start_row=1):
    ws.append(list(df.columns)) if start_row == 1 else None
    for row in df.itertuples(index=False):
        ws.append([("" if (v is None or (isinstance(v, float) and pd.isna(v))) else v) for v in row])
    last_col = get_column_letter(len(df.columns))
    ref = f"A{start_row}:{last_col}{start_row + max(len(df), 1)}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=True)
    ws.add_table(t)
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    for i, col in enumerate(df.columns, 1):
        w = (widths or {}).get(col) or min(max(len(str(col)) + 3, *(len(str(v)) for v in df[col].head(300))) + 2, 48)
        ws.column_dimensions[get_column_letter(i)].width = w
    return ref


def build_workbook(raw, clean_df, log_df, junk_df, dup_df, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.sheet_view.showGridLines = False

    # Clean contacts
    wc = wb.create_sheet("Clean_Contacts")
    write_table(wc, clean_df, "CleanContacts", widths={"Notes": 46, "Source Records": 16})
    n = len(clean_df) + 1
    col = {c: get_column_letter(i + 1) for i, c in enumerate(clean_df.columns)}
    for r in range(2, n + 1):
        wc[f"{col['Created Date']}{r}"].number_format = "yyyy-mm-dd"
    red = PatternFill("solid", fgColor="FDE2E1")
    amber = PatternFill("solid", fgColor="FFF4D6")
    wc.conditional_formatting.add(f"{col['Email Status']}2:{col['Email Status']}{n}",
                                  FormulaRule(formula=[f'LEFT(${col["Email Status"]}2,7)="Invalid"'], fill=red))
    wc.conditional_formatting.add(f"{col['Email Status']}2:{col['Email Status']}{n}",
                                  FormulaRule(formula=[f'${col["Email Status"]}2="Missing"'], fill=amber))
    wc.conditional_formatting.add(f"{col['Phone']}2:{col['Mobile']}{n}",
                                  FormulaRule(formula=[f'AND(${col["Phone"]}2="",${col["Mobile"]}2="")'], fill=amber))
    dv = DataValidation(type="list", formula1='"Website Form,Referral,LinkedIn,Trade Show,Cold Outreach,Google Ads"',
                        allow_blank=True)
    wc.add_data_validation(dv)
    dv.add(f"{col['Lead Source']}2:{col['Lead Source']}{n}")
    dv2 = DataValidation(type="list", formula1='"Valid,Invalid - verify with client,Missing"', allow_blank=False)
    wc.add_data_validation(dv2)
    dv2.add(f"{col['Email Status']}2:{col['Email Status']}{n}")

    wr = wb.create_sheet("Raw_Export")
    write_table(wr, raw, "RawExport", style="TableStyleLight1")
    wd = wb.create_sheet("Duplicates_Merged")
    write_table(wd, dup_df, "DuplicatesMerged", style="TableStyleMedium4")
    wj = wb.create_sheet("Junk_Removed")
    write_table(wj, junk_df, "JunkRemoved", style="TableStyleLight9")
    wl = wb.create_sheet("QA_Change_Log")
    write_table(wl, log_df, "QAChangeLog", style="TableStyleLight8",
                widths={"Before": 40, "After": 40, "Rule": 38})

    # Data dictionary
    wdd = wb.create_sheet("Data_Dictionary")
    dd = pd.DataFrame([
        ("Contact ID", "Unique ID assigned after de-duplication", "C-0001"),
        ("First Name / Last Name", "Split from Full Name; 'Last, First' order fixed; proper capitalization", "Grace | O'Brien"),
        ("Job Title", "Trimmed, capitalized; acronyms kept (VP, CEO, COO)", "VP of Operations"),
        ("Company", "Spelling unified across records; legal suffix standardized to LLC / Inc.", "Cedar Dental Partners LLC"),
        ("Email", "Lowercased and trimmed; invalid values kept as received for client review", "grace.obrien@domain.com"),
        ("Email Status", "Valid / Invalid - verify with client / Missing. Emails are never guessed", "Valid"),
        ("Phone / Mobile", "Display format (XXX) XXX-XXXX", "(312) 555-0167"),
        ("Phone (E.164) / Mobile (E.164)", "CRM/dialer import format +1XXXXXXXXXX", "+13125550167"),
        ("Ext", "Extension split from the phone field", "363"),
        ("State", "USPS 2-letter code", "IL"),
        ("ZIP", "5-digit text (leading zeros preserved)", "02134"),
        ("Lead Source", "Mapped to the CRM picklist (dropdown)", "Referral"),
        ("Created Date", "Earliest date across merged records, ISO format", "2025-10-15"),
        ("Source Records", "Raw export rows combined into this contact (audit trail)", "12, 345"),
        ("Notes", "Merges and items needing client action", "Merged records #12, #345"),
    ], columns=["Field", "Rule applied", "Example"])
    write_table(wdd, dd, "DataDictionary", style="TableStyleMedium2", widths={"Rule applied": 80, "Example": 30})

    # ---- Summary (live formulas)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 4
    ws.column_dimensions["E"].width = 40
    ws.column_dimensions["F"].width = 12
    ws["B2"] = "CRM Contact Cleanup - Delivery Summary"
    ws["B2"].font = Font(size=18, bold=True, color=NAVY)
    ws["B3"] = "All figures below are live formulas over the delivery tabs."
    ws["B3"].font = Font(italic=True, color="5B6770")

    def section(cell, text):
        ws[cell] = text
        ws[cell].font = Font(bold=True, color="FFFFFF")
        ws[cell].fill = PatternFill("solid", fgColor=TEAL)
        nxt = ws.cell(row=ws[cell].row, column=ws[cell].column + 1)
        nxt.fill = PatternFill("solid", fgColor=TEAL)

    section("B5", "Record counts")
    ec = col["Email Status"]
    pc, mc = col["Phone"], col["Mobile"]
    kpis = [
        ("Rows in client export", "=ROWS(RawExport[Full Name])"),
        ("Junk / test rows removed", "=ROWS(JunkRemoved[Reason])"),
        ("Duplicate rows merged", "=ROWS(DuplicatesMerged[Contact ID])"),
        ("Unique CRM-ready contacts", "=ROWS(CleanContacts[Contact ID])"),
        ("Check: export = removed + merged + unique", "=IF(C6=C7+C8+C9,\"OK\",\"Review\")"),
        ("Field corrections logged", "=ROWS(QAChangeLog[Rule])"),
    ]
    for i, (lab, f) in enumerate(kpis):
        ws.cell(row=6 + i, column=2, value=lab)
        ws.cell(row=6 + i, column=3, value=f)
    section("B13", "Contactability")
    cont = [
        ("Valid email", "=COUNTIF(CleanContacts[Email Status],\"Valid\")"),
        ("Invalid email (flagged for client)", "=COUNTIF(CleanContacts[Email Status],\"Invalid*\")"),
        ("No email on file", "=COUNTIF(CleanContacts[Email Status],\"Missing\")"),
        ("Has direct phone", "=COUNTIF(CleanContacts[Phone],\"?*\")"),
        ("Has mobile", "=COUNTIF(CleanContacts[Mobile],\"?*\")"),
        ("Mobile only (no direct line)", "=COUNTIFS(CleanContacts[Phone],\"\",CleanContacts[Mobile],\"?*\")"),
        ("No phone at all", "=COUNTIFS(CleanContacts[Phone],\"\",CleanContacts[Mobile],\"\")"),
        ("Valid email AND at least one phone",
         "=COUNTIFS(CleanContacts[Email Status],\"Valid\",CleanContacts[Phone],\"?*\")"
         "+COUNTIFS(CleanContacts[Email Status],\"Valid\",CleanContacts[Phone],\"\",CleanContacts[Mobile],\"?*\")"),
    ]
    for i, (lab, f) in enumerate(cont):
        ws.cell(row=14 + i, column=2, value=lab)
        ws.cell(row=14 + i, column=3, value=f)
    for r in list(range(6, 12)) + list(range(14, 22)):
        ws.cell(row=r, column=3).number_format = "#,##0"
        ws.cell(row=r, column=3).alignment = Alignment(horizontal="right")
        for c in (2, 3):
            ws.cell(row=r, column=c).border = Border(bottom=THIN)
    ws["C9"].font = Font(bold=True, size=12, color=TEAL)
    ws["C10"].font = Font(bold=True, color="2E7D32")

    section("E5", "Corrections by rule")
    rules = log_df["Rule"].value_counts().index.tolist()
    for i, rule in enumerate(rules):
        ws.cell(row=6 + i, column=5, value=rule)
        c = ws.cell(row=6 + i, column=6, value=f'=COUNTIF(QAChangeLog[Rule],E{6 + i})')
        c.number_format = "#,##0"
        for cc in (5, 6):
            ws.cell(row=6 + i, column=cc).border = Border(bottom=THIN)
    last = 5 + len(rules)
    ch = BarChart()
    ch.type = "bar"
    ch.style = 10
    ch.y_axis.title = None
    ch.legend = None
    ch.add_data(Reference(ws, min_col=6, min_row=6, max_row=last), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=5, min_row=6, max_row=last))
    ch.y_axis.majorGridlines = None
    ch.x_axis.delete = False
    ch.y_axis.delete = True
    ch.x_axis.scaling.orientation = "maxMin"
    ch.dataLabels = DataLabelList(showVal=True, showSerName=False, showCatName=False,
                                  showLegendKey=False, showPercent=False)
    ch.height, ch.width = 9.5, 17
    ch.series[0].graphicalProperties.solidFill = TEAL
    ws.add_chart(ch, "H5")

    section("B24", "Method")
    method = ["1. Trimmed spaces and removed test / empty rows.",
              "2. Standardized names, titles, companies, states, ZIPs, sources and dates.",
              "3. Phones converted to E.164 (+1XXXXXXXXXX) with extensions split out.",
              "4. Emails validated; bad ones flagged for the client - never guessed.",
              "5. Duplicates matched on email, phone, mobile and name + company;",
              "   the most complete record was kept and its gaps filled from the copies.",
              "6. Every change is logged in QA_Change_Log (before / after / rule)."]
    for i, t in enumerate(method):
        ws.cell(row=25 + i, column=2, value=t)
    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    run(sys.argv[1], sys.argv[2])
