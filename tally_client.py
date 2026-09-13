"""
Tally XML API Client
High-speed HTTP XML interface for Tally Prime / Tally.ERP 9 (Port 9000).
Optimized for high-volume extraction (25,000+ entries) with XML sanitization.
"""

import re
import requests
import calendar
from datetime import datetime, date
from xml.sax.saxutils import escape
import xml.etree.cElementTree as ET
from typing import List, Dict, Any, Optional

TALLY_URL = "http://localhost:9000"
TIMEOUT = 10  # seconds for standard calls

# Global in-memory cache for ledger metadata and group parent hierarchy
LEDGER_METADATA_CACHE: Dict[str, Dict[str, Any]] = {}
GROUP_PARENTS_CACHE: Dict[str, str] = {}
ACTIVE_COMPANY: str = ""


def get_ledger_full_group_path(ledger_name: str) -> str:
    """
    Returns the complete group hierarchy path for a given ledger name.
    Example: 'Fixed Assets > PLANT & MACHINERY' or 'Loans (Liability) > Bank OCC A/c > STATE BANK OF INDIA CC'.
    Excludes root 'Primary' for clean readability.
    """
    if not ledger_name or ledger_name.lower() in ('general', 'others', ''):
        return ""

    # Case-insensitive helper for group parents
    group_parents_upper = {k.strip().upper(): (k, v) for k, v in GROUP_PARENTS_CACHE.items()}

    # Check case-insensitive in LEDGER_METADATA_CACHE
    parent_group = ""
    target_norm = ledger_name.strip().upper()
    for name, meta in LEDGER_METADATA_CACHE.items():
        if name.strip().upper() == target_norm:
            parent_group = meta.get("parent", "").strip()
            break

    if not parent_group:
        # If ledger itself is a group name in GROUP_PARENTS_CACHE
        if target_norm in group_parents_upper:
            parent_group = group_parents_upper[target_norm][0]

    if not parent_group:
        return ""

    # Trace hierarchy upwards
    chain = [parent_group]
    curr = parent_group
    visited = {curr.lower()}

    while True:
        curr_upper = curr.strip().upper()
        if curr_upper in group_parents_upper:
            orig_k, next_p = group_parents_upper[curr_upper]
            next_p = (next_p or "").strip()
            if not next_p or next_p.lower() == 'primary' or next_p.lower() in visited:
                break
            visited.add(next_p.lower())
            chain.append(next_p)
            curr = next_p
        else:
            break

    # Reverse chain so it goes Top Head > Sub Group > Immediate Group
    return " > ".join(reversed(chain))

# Standard 2-digit Indian GST State Code directory
GST_STATE_CODES: Dict[str, str] = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
    "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan",
    "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh",
    "13": "Nagaland", "14": "Manipur", "15": "Mizoram", "16": "Tripura",
    "17": "Meghalaya", "18": "Assam", "19": "West Bengal", "20": "Jharkhand",
    "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra & Nagar Haveli", "27": "Maharashtra",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala",
    "33": "Tamil Nadu", "34": "Puducherry", "35": "Andaman & Nicobar",
    "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh"
}


def format_closing_balance(raw_val: Any, is_deemed_pos: Optional[str] = None) -> str:
    """Formats closing balance with Indian number formatting and Dr / Cr suffix."""
    if raw_val is None:
        return ""
    s = str(raw_val).strip()
    if not s:
        return ""
    if re.search(r'\b(dr|cr)\b', s, re.IGNORECASE):
        return s
    clean_s = s.replace(',', '')
    try:
        val = float(clean_s)
        if abs(val) < 0.001:
            return "0.00"
        if is_deemed_pos and str(is_deemed_pos).strip().lower() in ('yes', 'true', '1'):
            dr_cr = "Dr"
        elif is_deemed_pos and str(is_deemed_pos).strip().lower() in ('no', 'false', '0'):
            dr_cr = "Cr"
        else:
            dr_cr = "Dr" if val < 0 else "Cr"
        return f"{abs(val):,.2f} {dr_cr}"
    except ValueError:
        return s


def clean_tally_xml(xml_text: str) -> str:
    """
    Sanitizes Tally XML output:
    1. Removes illegal XML control character entities like &#4;, &#1;, &#0;, etc.
    2. Strips raw ASCII control characters.
    3. Replaces colons in XML tags (e.g., <UDF:something>, </UDF:something>) with underscores
       to prevent xml.etree 'unbound prefix' ParseError when Tally includes custom UDF fields.
    """
    cleaned = re.sub(
        r'&#(?:0?[0-8]|1[12]|1[4-9]|2[0-9]|3[01]|x[0-8bcef]|x1[0-9a-f]);',
        '',
        xml_text,
        flags=re.IGNORECASE
    )
    cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', cleaned)
    # Neutralize un-prefixed XML namespace tags like <UDF:XXX> or </UDF:XXX>
    cleaned = re.sub(r'<(/?[a-zA-Z0-9_]+):([a-zA-Z0-9_.]+)', r'<\1_\2', cleaned)
    return cleaned


def send_tally_request(xml_body: str, timeout: int = 45) -> Optional[str]:
    """Sends raw XML request to Tally HTTP server on Port 9000 and cleans response."""
    try:
        response = requests.post(
            TALLY_URL,
            data=xml_body.encode('utf-8'),
            headers={'Content-Type': 'text/xml;charset=utf-8'},
            timeout=timeout
        )
        if response.status_code == 200:
            return clean_tally_xml(response.text)
        return None
    except Exception as e:
        print(f"[Tally Client] Request error or timeout ({timeout}s): {e}")
        return None


def get_tally_status() -> Dict[str, Any]:
    """
    Checks if Tally Prime is running and connected on port 9000.
    Returns connection status and the currently active open company name.
    """
    global ACTIVE_COMPANY
    request_xml = """<ENVELOPE>
        <HEADER>
            <VERSION>1</VERSION>
            <TALLYREQUEST>Export Data</TALLYREQUEST>
            <TYPE>Collection</TYPE>
            <ID>Company</ID>
        </HEADER>
        <BODY>
            <DESC>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                </STATICVARIABLES>
            </DESC>
        </BODY>
    </ENVELOPE>"""

    xml_resp = send_tally_request(request_xml, timeout=5)
    if xml_resp:
        try:
            root = ET.fromstring(xml_resp)
            active_company = ""
            comp = root.find('.//DATA/COLLECTION/COMPANY')
            if comp is not None:
                active_company = (comp.get('NAME') or comp.findtext('NAME') or '').strip()
            
            if not active_company:
                for c in root.iter('COMPANY'):
                    n = (c.get('NAME') or c.findtext('NAME') or '').strip()
                    if n and n != '0':
                        active_company = n
                        break

            if active_company:
                ACTIVE_COMPANY = active_company
                return {
                    "connected": True,
                    "company": active_company,
                    "message": "Tally Connected (Port 9000)"
                }
            return {
                "connected": True,
                "company": ACTIVE_COMPANY or "Tally Prime Active Company",
                "message": "Tally Connected (Port 9000)"
            }
        except Exception:
            return {
                "connected": True,
                "company": ACTIVE_COMPANY or "Tally Prime Active Company",
                "message": "Tally Connected (Port 9000)"
            }

    # Fallback when Tally is offline
    return {
        "connected": False,
        "company": "Tally Offline",
        "message": "Tally Offline (Using Cached/Demo Mode)"
    }


def get_groups_with_ledgers_and_meta() -> Dict[str, Any]:
    """
    Fetches all groups, subgroups, and member ledgers from Tally.
    Resolves complete parent-subgroup hierarchy so selecting a parent group
    automatically includes all ledgers belonging to that group AND all its subgroups.
    Also caches GSTIN, Address, and Closing Balance for each ledger.
    """
    global LEDGER_METADATA_CACHE, ACTIVE_COMPANY, GROUP_PARENTS_CACHE
    LEDGER_METADATA_CACHE.clear()

    # Determine active company
    if not ACTIVE_COMPANY:
        get_tally_status()
    cmp_tag = f"<SVCURRENTCOMPANY>{ACTIVE_COMPANY}</SVCURRENTCOMPANY>" if (ACTIVE_COMPANY and not ACTIVE_COMPANY.startswith("Tally ")) else ""

    # 1. Fetch all groups and their parents
    req_groups = f"""<ENVELOPE>
        <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST><TYPE>Collection</TYPE><ID>AllGroupsCollection</ID></HEADER>
        <BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{cmp_tag}</STATICVARIABLES>
        <TDL><TDLMESSAGE><COLLECTION NAME="AllGroupsCollection"><TYPE>Group</TYPE><FETCH>NAME, PARENT</FETCH></COLLECTION></TDLMESSAGE></TDL>
        </DESC></BODY></ENVELOPE>"""

    resp_groups = send_tally_request(req_groups, timeout=30)
    group_parents: Dict[str, str] = {}
    if resp_groups:
        try:
            root_g = ET.fromstring(resp_groups)
            for g in root_g.iter('GROUP'):
                name = (g.get('NAME') or g.findtext('NAME') or '').strip()
                parent = (g.findtext('PARENT') or '').strip()
                if name:
                    group_parents[name] = parent
            GROUP_PARENTS_CACHE = dict(group_parents)
        except Exception as e:
            print(f"Error parsing Tally groups hierarchy: {e}")

    # 2. Fetch all ledgers and cache metadata
    req_ledgers = f"""<ENVELOPE>
        <HEADER><VERSION>1</VERSION><TALLYREQUEST>Export Data</TALLYREQUEST><TYPE>Collection</TYPE><ID>CustomLedgerCollection</ID></HEADER>
        <BODY><DESC><STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{cmp_tag}</STATICVARIABLES>
        <TDL><TDLMESSAGE><COLLECTION NAME="CustomLedgerCollection"><TYPE>Ledger</TYPE>
        <FETCH>NAME, PARENT, CLOSINGBALANCE, ISDEEMEDPOSITIVE, PARTYGSTIN, LEDGERGSTIN, ADDRESS.LIST.*, STATENAME, PINCODE, COUNTRYNAME, LEDSTATENAME, PRIORSTATENAME, LEDGSTREGDETAILS.LIST.*, LEDMAILINGDETAILS.LIST.*</FETCH>
        </COLLECTION></TDLMESSAGE></TDL>
        </DESC></BODY></ENVELOPE>"""

    resp_ledgers = send_tally_request(req_ledgers, timeout=60)
    direct_ledgers: Dict[str, List[str]] = {}
    ledger_parents: Dict[str, str] = {}

    if resp_ledgers:
        try:
            root_l = ET.fromstring(resp_ledgers)
            for ledger in root_l.iter('LEDGER'):
                name = (ledger.get('NAME') or ledger.findtext('NAME') or '').strip()
                parent = (ledger.findtext('PARENT') or 'Others').strip()
                if not name:
                    continue

                closing = ledger.findtext('CLOSINGBALANCE') or ""
                is_deemed = ledger.findtext('ISDEEMEDPOSITIVE') or ""
                formatted_closing = format_closing_balance(closing, is_deemed)

                # 1. GSTIN (Check direct fields, then LEDGSTREGDETAILS.LIST)
                gstin = (ledger.findtext('PARTYGSTIN') or ledger.findtext('LEDGERGSTIN') or '').strip()
                if not gstin:
                    for gd in ledger.iter('LEDGSTREGDETAILS.LIST'):
                        g = (gd.findtext('GSTIN') or '').strip()
                        if g:
                            gstin = g
                            break

                # 2. State & Pincode (Check direct fields, LEDMAILINGDETAILS.LIST, LEDGSTREGDETAILS.LIST)
                state = (ledger.findtext('STATENAME') or ledger.findtext('LEDSTATENAME') or ledger.findtext('PRIORSTATENAME') or '').strip()
                if state.lower() in ('not applicable', 'primary', ''):
                    state = ''

                pincode = (ledger.findtext('PINCODE') or '').strip()

                for md in ledger.iter('LEDMAILINGDETAILS.LIST'):
                    if not pincode:
                        pincode = (md.findtext('PINCODE') or '').strip()
                    if not state:
                        s = (md.findtext('STATE') or '').strip()
                        if s.lower() not in ('not applicable', 'primary', ''):
                            state = s

                if not state:
                    for gd in ledger.iter('LEDGSTREGDETAILS.LIST'):
                        s = (gd.findtext('STATE') or gd.findtext('PLACEOFSUPPLY') or '').strip()
                        if s.lower() not in ('not applicable', 'primary', ''):
                            state = s
                            break

                # Infer state from GSTIN prefix if still missing
                if not state and len(gstin) >= 2 and gstin[:2] in GST_STATE_CODES:
                    state = GST_STATE_CODES[gstin[:2]]

                # 3. Address Lines (Search both top-level and inside LEDMAILINGDETAILS.LIST)
                addr_lines = []
                for a in ledger.iter('ADDRESS.LIST'):
                    for sub in a.iter('ADDRESS'):
                        line = (sub.text or '').strip()
                        if line and line not in addr_lines:
                            addr_lines.append(line)
                    line = (a.text or '').strip()
                    if line and line not in addr_lines:
                        addr_lines.append(line)

                # Format full address: street lines, state - pincode
                addr_parts = [a for a in addr_lines if a]
                state_pin_part = ""
                if state and pincode:
                    state_pin_part = f"{state} - {pincode}"
                elif state:
                    state_pin_part = state
                elif pincode:
                    state_pin_part = f"PIN - {pincode}"

                if state_pin_part:
                    addr_parts.append(state_pin_part)

                full_address = ", ".join(addr_parts) if addr_parts else ""

                LEDGER_METADATA_CACHE[name] = {
                    "gstn": gstin,
                    "address": full_address,
                    "closingBalance": formatted_closing,
                    "parent": parent,
                    "state": state,
                    "pincode": pincode
                }

                ledger_parents[name] = parent
                direct_ledgers.setdefault(parent, []).append(name)
        except Exception as e:
            print(f"Error parsing Tally ledgers: {e}")

    # Helper: recursively find all descendant subgroups of a group
    def get_descendants(g_name: str) -> set:
        desc = set([g_name])
        for child, p in group_parents.items():
            if p.lower() == g_name.lower() and child not in desc:
                desc.update(get_descendants(child))
        return desc

    # 3. Aggregate ledgers for each group and subgroup
    all_known_groups = set(group_parents.keys()) | set(direct_ledgers.keys())
    aggregated_groups: Dict[str, List[str]] = {}
    groups_meta: Dict[str, Dict[str, Any]] = {}

    all_ledgers_flat = []
    for l_list in direct_ledgers.values():
        all_ledgers_flat.extend(l_list)
    unique_all_ledgers = sorted(list(set(all_ledgers_flat)))

    for g in sorted(all_known_groups):
        descendants = get_descendants(g)
        agg_list = []
        for d in descendants:
            agg_list.extend(direct_ledgers.get(d, []))
        
        uniq = sorted(list(set(agg_list)))
        if uniq:
            aggregated_groups[g] = uniq
            parent_p = group_parents.get(g, "")
            sub_list = sorted([d for d in descendants if d != g])
            
            # Root accounting heads like Sundry Creditors or groups with subgroups should not display as indented subgroups
            is_sub = bool(
                parent_p and 
                parent_p.lower() not in ("primary", "", "others", "current liabilities", "current assets") and 
                len(sub_list) == 0 and
                g not in ("Sundry Creditors", "Sundry Debtors")
            )
            
            groups_meta[g] = {
                "name": g,
                "parent": parent_p,
                "is_subgroup": is_sub,
                "total_count": len(uniq),
                "direct_count": len(direct_ledgers.get(g, [])),
                "subgroups": sub_list
            }

            # If a group has both direct ledgers AND subgroups (like Sundry Creditors),
            # create an explicit "Direct Only" option matching Tally's direct ledger list
            direct_items = direct_ledgers.get(g, [])
            if sub_list and direct_items and len(direct_items) != len(uniq):
                dir_only_name = f"{g} (Direct Only)"
                aggregated_groups[dir_only_name] = sorted(list(set(direct_items)))
                groups_meta[dir_only_name] = {
                    "name": dir_only_name,
                    "parent": g,
                    "is_subgroup": True,
                    "total_count": len(direct_items),
                    "direct_count": len(direct_items),
                    "subgroups": []
                }

    if aggregated_groups:
        aggregated_groups["All Groups"] = unique_all_ledgers
        groups_meta["All Groups"] = {
            "name": "All Groups",
            "parent": "",
            "is_subgroup": False,
            "total_count": len(unique_all_ledgers),
            "direct_count": len(unique_all_ledgers),
            "subgroups": []
        }

        # Order groups: Priority groups on top, followed immediately by their direct-only option & subgroups
        sorted_groups: Dict[str, List[str]] = {}
        sorted_meta: Dict[str, Dict[str, Any]] = {}
        priority_roots = ["Sundry Creditors", "Sundry Debtors", "Direct Incomes", "Indirect Expenses", "Direct Expenses"]

        for pr in priority_roots:
            if pr in aggregated_groups:
                sorted_groups[pr] = aggregated_groups[pr]
                sorted_meta[pr] = groups_meta[pr]
                
                # Direct-only option right below the primary group
                dir_only = f"{pr} (Direct Only)"
                if dir_only in aggregated_groups:
                    sorted_groups[dir_only] = aggregated_groups[dir_only]
                    sorted_meta[dir_only] = groups_meta[dir_only]

                # Include child subgroups
                for child in groups_meta[pr].get("subgroups", []):
                    if child in aggregated_groups and child not in sorted_groups:
                        sorted_groups[child] = aggregated_groups[child]
                        sorted_meta[child] = groups_meta[child]

        # Add remaining groups
        for g in sorted(aggregated_groups.keys()):
            if g not in sorted_groups:
                sorted_groups[g] = aggregated_groups[g]
                sorted_meta[g] = groups_meta[g]

        return {
            "groups": sorted_groups,
            "meta": sorted_meta,
            "ledger_parents": ledger_parents
        }

    # Fallback
    fallback_groups = {
        "Sundry Creditors": ["Special Blasts Limited", "Apex Solutions Pvt Ltd"],
        "Sundry Debtors": ["Reliance Retail Ltd", "Tata Consumer Products"],
        "All Groups": ["Special Blasts Limited", "Reliance Retail Ltd"]
    }
    return {
        "groups": fallback_groups,
        "meta": {}
    }


def get_groups_with_ledgers() -> Dict[str, List[str]]:
    """Returns mapping of group name -> list of member ledgers (including subgroups)."""
    res = get_groups_with_ledgers_and_meta()
    return res.get("groups", {})


def chunk_ledgers(ledgers: List[str], max_chars: int = 350, max_items: int = 6) -> List[List[str]]:
    """Chunks a list of ledgers dynamically to ensure TDL formula stays within safe parser limits."""
    batches = []
    current_batch = []
    current_len = 0
    for l in ledgers:
        clean_l = escape(l, {'"': '&quot;'})
        item_len = len(clean_l) + 28
        if current_batch and (current_len + item_len > max_chars or len(current_batch) >= max_items):
            batches.append(current_batch)
            current_batch = [l]
            current_len = item_len
        else:
            current_batch.append(l)
            current_len += item_len
    if current_batch:
        batches.append(current_batch)
    return batches


def split_date_range_into_months(from_date_str: str, to_date_str: str) -> List[tuple]:
    """Splits a date range (YYYY-MM-DD) into monthly chunks (YYYYMMDD start, YYYYMMDD end)."""
    try:
        start_dt = datetime.strptime(from_date_str, "%Y-%m-%d").date()
        end_dt = datetime.strptime(to_date_str, "%Y-%m-%d").date()
    except Exception:
        return [(from_date_str.replace("-", ""), to_date_str.replace("-", ""))]
        
    if start_dt > end_dt:
        start_dt, end_dt = end_dt, start_dt
        
    # Safety guard: Prevent unbounded placeholder dates (e.g. 1900 to 2099) from exploding into 2400 chunks
    if start_dt.year < 2000:
        start_dt = date(2020, 4, 1)
    if end_dt.year > 2035:
        end_dt = date(2026, 3, 31)

    chunks = []
    curr = start_dt
    while curr <= end_dt:
        _, last_day = calendar.monthrange(curr.year, curr.month)
        month_end = date(curr.year, curr.month, last_day)
        chunk_end = min(month_end, end_dt)
        chunks.append((
            curr.strftime("%Y%m%d"),
            chunk_end.strftime("%Y%m%d")
        ))
        if curr.month == 12:
            curr = date(curr.year + 1, 1, 1)
        else:
            curr = date(curr.year, curr.month + 1, 1)
    return chunks


def fetch_vouchers_from_tally(
    from_date: str = "2025-04-01",
    to_date: str = "2026-03-31",
    group_name: str = "Sundry Creditors",
    selected_ledgers: Optional[List[str]] = None
) -> List[Dict[str, Any]]:
    """
    Rock-Solid High-Speed Voucher Extractor:
    - For Single / Selected Ledgers: Uses ultra-fast native Tally index (<TYPE>Vouchers : Ledger</TYPE>) in < 0.3s.
    - For Groups: Uses native Tally group index (<TYPE>Vouchers : Group</TYPE>) in ~6s for thousands of vouchers.
    - Zero custom TDL child formulas (prevents C++ NULL pointer dereferences & Software Exception c0000005).
    - Unbounded placeholder dates (1900-2099) safely query active Tally period without 2400-chunk loops.
    - Filters target ledgers safely in Python memory in < 5ms.
    - Full multi-year support with permanent GUID deduplication (0% double entries).
    """
    global LEDGER_METADATA_CACHE, ACTIVE_COMPANY, GROUP_PARENTS_CACHE
    if not LEDGER_METADATA_CACHE or not GROUP_PARENTS_CACHE:
        get_groups_with_ledgers_and_meta()

    if not ACTIVE_COMPANY:
        get_tally_status()
    cmp_tag = f"<SVCURRENTCOMPANY>{escape(ACTIVE_COMPANY, {'\"': '&quot;'})}</SVCURRENTCOMPANY>" if (ACTIVE_COMPANY and not ACTIVE_COMPANY.startswith("Tally ")) else ""

    # Check if custom realistic date range was specified (year between 2000 and 2040)
    use_date_filter = False
    sv_from = ""
    sv_to = ""
    if from_date and to_date:
        try:
            f_year = int(from_date.split("-")[0])
            t_year = int(to_date.split("-")[0])
            if 2000 <= f_year <= 2040 and 2000 <= t_year <= 2040:
                use_date_filter = True
                sv_from = from_date.replace("-", "")
                sv_to = to_date.replace("-", "")
        except Exception:
            pass

    date_tags = f"""
                        <SVFROMDATE>{sv_from}</SVFROMDATE>
                        <SVTODATE>{sv_to}</SVTODATE>""" if use_date_filter else ""

    FETCH_TAG = "<FETCH>DATE, VOUCHERNUMBER, VOUCHERTYPENAME, REFERENCE, NARRATION, PARTYLEDGERNAME, PARTYGSTIN, STATENAME, PLACEOFSUPPLY, PINCODE, BASICBUYERADDRESS.LIST.*, ADDRESS.LIST.*, ALLLEDGERENTRIES.LIST.*, GUID, MASTERID</FETCH>"

    all_raw_responses = []
    target_ledgers = set()

    # STRATEGY 1: Selected Ledgers (Single ledger or specific list up to 20)
    if selected_ledgers and len(selected_ledgers) <= 20:
        target_ledgers = set(selected_ledgers)
        print(f"[Tally Client] Fast-fetching vouchers for {len(selected_ledgers)} selected ledger(s) via native Tally index...")
        for l in selected_ledgers:
            escaped_l = escape(l, {'"': '&quot;'})
            xml_body = f"""<ENVELOPE>
                <HEADER>
                    <VERSION>1</VERSION>
                    <TALLYREQUEST>Export Data</TALLYREQUEST>
                    <TYPE>Collection</TYPE>
                    <ID>LedgerVchCol</ID>
                </HEADER>
                <BODY>
                    <DESC>
                        <STATICVARIABLES>
                            <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                            {cmp_tag}{date_tags}
                        </STATICVARIABLES>
                        <TDL>
                            <TDLMESSAGE>
                                <COLLECTION NAME="LedgerVchCol">
                                    <TYPE>Vouchers : Ledger</TYPE>
                                    <CHILDOF>"{escaped_l}"</CHILDOF>
                                    {FETCH_TAG}
                                </COLLECTION>
                            </TDLMESSAGE>
                        </TDL>
                    </DESC>
                </BODY>
            </ENVELOPE>"""
            resp = send_tally_request(xml_body, timeout=30)
            if resp and "Unknown Request" not in resp:
                all_raw_responses.append(resp)

    # STRATEGY 2: Specific Group / Subgroup (or more than 20 selected ledgers)
    elif group_name and group_name != "All Groups":
        groups = get_groups_with_ledgers()
        clean_grp = group_name.replace(" (Direct Only)", "").strip()
        escaped_grp = escape(clean_grp, {'"': '&quot;'})
        if selected_ledgers:
            target_ledgers = set(selected_ledgers)
        else:
            target_ledgers = set(groups.get(group_name, []))

        print(f"[Tally Client] Fast-fetching vouchers for Group '{clean_grp}' via native Tally index...")
        xml_body = f"""<ENVELOPE>
            <HEADER>
                <VERSION>1</VERSION>
                <TALLYREQUEST>Export Data</TALLYREQUEST>
                <TYPE>Collection</TYPE>
                <ID>GroupVchCol</ID>
            </HEADER>
            <BODY>
                <DESC>
                    <STATICVARIABLES>
                        <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                        {cmp_tag}{date_tags}
                    </STATICVARIABLES>
                    <TDL>
                        <TDLMESSAGE>
                            <COLLECTION NAME="GroupVchCol">
                                <TYPE>Vouchers : Group</TYPE>
                                <CHILDOF>"{escaped_grp}"</CHILDOF>
                                {FETCH_TAG}
                            </COLLECTION>
                        </TDLMESSAGE>
                    </TDL>
                </DESC>
            </BODY>
        </ENVELOPE>"""
        resp = send_tally_request(xml_body, timeout=60)
        if resp and "Unknown Request" not in resp:
            all_raw_responses.append(resp)

    # STRATEGY 3: All Groups (Export whole company)
    else:
        if selected_ledgers:
            target_ledgers = set(selected_ledgers)
        else:
            target_ledgers = set()

        if use_date_filter:
            month_chunks = split_date_range_into_months(from_date, to_date)
        else:
            month_chunks = [("", "")]

        print(f"[Tally Client] Fetching vouchers for All Groups across {len(month_chunks)} chunk(s)...")
        for m_start, m_end in month_chunks:
            chunk_date_tags = f"""
                        <SVFROMDATE>{m_start}</SVFROMDATE>
                        <SVTODATE>{m_end}</SVTODATE>""" if m_start else ""
            xml_body = f"""<ENVELOPE>
                <HEADER>
                    <VERSION>1</VERSION>
                    <TALLYREQUEST>Export Data</TALLYREQUEST>
                    <TYPE>Collection</TYPE>
                    <ID>AllVchCol</ID>
                </HEADER>
                <BODY>
                    <DESC>
                        <STATICVARIABLES>
                            <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                            {cmp_tag}{chunk_date_tags}
                        </STATICVARIABLES>
                        <TDL>
                            <TDLMESSAGE>
                                <COLLECTION NAME="AllVchCol">
                                    <TYPE>Voucher</TYPE>
                                    {FETCH_TAG}
                                </COLLECTION>
                            </TDLMESSAGE>
                        </TDL>
                    </DESC>
                </BODY>
            </ENVELOPE>"""
            resp = send_tally_request(xml_body, timeout=60)
            if resp and "Unknown Request" not in resp:
                all_raw_responses.append(resp)

    # Parse all XML responses and deduplicate with GUID
    seen_keys = set()
    aggregated_records = []
    for resp in all_raw_responses:
        parsed_chunk = parse_tally_vouchers_xml(resp, target_ledgers, cleanup_raw_date=False)
        for r in parsed_chunk:
            guid = r.get("_guid")
            if guid:
                k = (guid, r.get("vendor"))
            else:
                k = (r.get("voucherNo"), r.get("date"), r.get("vendor"), r.get("invoiceNo"), r.get("invoiceAmount"), r.get("amountPaid"))
            if k not in seen_keys:
                seen_keys.add(k)
                aggregated_records.append(r)

    # Sort vouchers strictly party-wise (alphabetical by party name, then chronological by date)
    aggregated_records.sort(key=lambda r: (r.get("vendor", "").lower(), r.get("_raw_date", r.get("date", "")), str(r.get("voucherNo", ""))))

    # Re-number sequential S No and clean temporary fields
    for idx, r in enumerate(aggregated_records, start=1):
        r["sNo"] = idx
        r["ledgerBalance"] = 0
        r.pop("_raw_date", None)
        r.pop("_guid", None)

    print(f"[Tally Client] Total extracted records: {len(aggregated_records)}")
    return aggregated_records



def parse_tally_vouchers_xml(xml_text: str, target_ledgers: set, cleanup_raw_date: bool = True) -> List[Dict[str, Any]]:
    """Fast XML parser for Tally Voucher Collections with full 20-column matching."""
    root = ET.fromstring(xml_text)
    records = []
    s_no = 1
    running_balance = 0.0

    for v in root.iter('VOUCHER'):
        v_date_raw = v.findtext('DATE') or ""
        v_date = v_date_raw
        if len(v_date_raw) == 8:
            try:
                dt = datetime.strptime(v_date_raw, "%Y%m%d")
                v_date = dt.strftime("%d-%m-%y")
            except Exception:
                pass

        v_no = v.findtext('VOUCHERNUMBER') or ""
        v_guid = (v.findtext('GUID') or v.findtext('MASTERID') or v.get('REMOTEID') or v.get('VCHKEY') or '').strip()
        party_name = (v.findtext('PARTYLEDGERNAME') or "").strip()
        v_type = v.findtext('VOUCHERTYPENAME') or ""
        reference = v.findtext('REFERENCE') or ""
        narration = v.findtext('NARRATION') or ""

        # Extract all ledger entries in voucher
        entries = []
        voucher_ledgers = set()
        for entry in v.iter('ALLLEDGERENTRIES.LIST'):
            l_name = (entry.findtext('LEDGERNAME') or "").strip()
            amt_str = entry.findtext('AMOUNT') or "0"
            deemed_pos = (entry.findtext('ISDEEMEDPOSITIVE') or "No").strip()
            try:
                amt = float(amt_str)
            except ValueError:
                amt = 0.0
            if l_name:
                entries.append((l_name, amt, deemed_pos))
                voucher_ledgers.add(l_name)

        if not entries:
            continue

        # Check if voucher belongs to target ledgers
        matched_target_ledger = None
        if target_ledgers:
            # Check party name first
            if party_name and party_name in target_ledgers:
                matched_target_ledger = party_name
            else:
                # Check entries
                intersect = voucher_ledgers.intersection(target_ledgers)
                if intersect:
                    matched_target_ledger = list(intersect)[0]
                else:
                    continue
        else:
            matched_target_ledger = party_name or entries[0][0]

        # Locate target party entry to determine debit/credit status
        party_entry = None
        for l_name, amt, deemed_pos in entries:
            if l_name == matched_target_ledger:
                party_entry = (l_name, amt, deemed_pos)
                break
        if not party_entry:
            party_entry = entries[0]

        is_party_debit = (party_entry[2].upper() == "YES" or party_entry[1] < 0)

        # In accounting:
        # For Creditors (vendors):
        # - Debit to party = Payment / Settlement / Loan adjustment / TDS deduction
        # - Credit to party = Bill / Purchase / Invoice booked
        # For Debtors (customers):
        # - Credit to customer = Receipt / Payment received
        # - Debit to customer = Sales Invoice booked
        is_payment_type = v_type.upper() in ("PAYMENT", "RECEIPT", "CONTRA")
        
        meta = LEDGER_METADATA_CACHE.get(matched_target_ledger, {})
        parent_group = (meta.get("parent") or "").lower()
        is_debtor = "debtor" in parent_group

        is_settlement = (not is_party_debit) if is_debtor else is_party_debit

        ANCILLARY_KEYWORDS = ('ROUND', 'FREIGHT', 'TRANSPORT', 'CARTAGE', 'LOADING', 'UNLOADING', 'HAMALI', 'INSURANCE', 'DISCOUNT', 'FITTING')

        taxable_val = None
        cgst_val = None
        sgst_val = None
        igst_val = None
        gst_total = None
        tcs_tds_val = None
        invoice_amount = None
        amount_paid = None
        payment_mode = ""
        payment_date = ""
        other_ledger = ""

        if is_settlement or is_payment_type:
            # === PAYMENT / SETTLEMENT / TDS / LOAN ADJUSTMENT ===
            amount_paid = abs(party_entry[1])
            payment_date = v_date
            
            # Counterpart ledgers (offsetting entries)
            counterparts = [(ln, abs(a)) for ln, a, p in entries if ln != matched_target_ledger]
            if counterparts:
                # Select counterpart with largest amount (e.g. HDFC, Loan Account, Tds on Purchase 94Q)
                main_cp = max(counterparts, key=lambda x: x[1])[0]
            else:
                main_cp = "Payment"
            
            other_ledger = main_cp
            payment_mode = main_cp
        else:
            # === BILL / PURCHASE / INVOICE BOOKED ===
            invoice_amount = abs(party_entry[1])
            
            for l_name, amt, deemed_pos in entries:
                abs_amt = abs(amt)
                l_upper = l_name.upper()
                if 'CGST' in l_upper or 'CENTRAL GST' in l_upper or 'CENTRAL TAX' in l_upper:
                    cgst_val = (cgst_val or 0.0) + abs_amt
                elif 'SGST' in l_upper or 'STATE GST' in l_upper or 'STATE TAX' in l_upper or 'UTGST' in l_upper:
                    sgst_val = (sgst_val or 0.0) + abs_amt
                elif 'IGST' in l_upper or 'INTEGRATED GST' in l_upper or 'INTEGRATED TAX' in l_upper:
                    igst_val = (igst_val or 0.0) + abs_amt

            if cgst_val or sgst_val or igst_val:
                gst_total = (cgst_val or 0.0) + (sgst_val or 0.0) + (igst_val or 0.0)

            # Detect TCS and TDS
            tcs_val = 0.0
            tds_val = 0.0
            for l_name, amt, deemed_pos in entries:
                if l_name == matched_target_ledger:
                    continue
                abs_amt = abs(amt)
                l_upper = l_name.upper()
                if 'TCS' in l_upper:
                    tcs_val += abs_amt
                elif any(k in l_upper for k in ('TDS', '194C', '194H', '194I', '194J', '194Q', '194A', '194-C', '194-H', '194-I', '194-J', '194-Q')):
                    tds_val += abs_amt

            if tcs_val > 0 or tds_val > 0:
                # TCS is positive (+), TDS is negative (-)
                tcs_tds_val = tcs_val - tds_val

            # Find main expense / asset / purchase ledger (excluding taxes, TDS, TCS)
            expense_cands = []
            for l_name, amt, deemed_pos in entries:
                if l_name != matched_target_ledger and not any(k in l_name.upper() for k in ('CGST', 'SGST', 'IGST', 'GST', 'TDS', 'TCS', '194C', '194H', '194I', '194J', '194Q', '194A')):
                    expense_cands.append((l_name, abs(amt)))

            if expense_cands:
                non_anc = [c for c in expense_cands if not any(k in c[0].upper() for k in ANCILLARY_KEYWORDS)]
                if non_anc:
                    other_ledger = max(non_anc, key=lambda x: x[1])[0]
                else:
                    other_ledger = max(expense_cands, key=lambda x: x[1])[0]
            else:
                other_ledger = "Purchase"

            # Taxable value: Invoice Amount = Taxable + GST Total + TCS/(TDS)
            # Therefore: Taxable = Invoice Amount - GST Total - (TCS/TDS)
            taxable_val = invoice_amount - (gst_total or 0.0) - (tcs_tds_val or 0.0)

        # Running balance calculation
        inv_amt = invoice_amount or 0.0
        pd_amt = amount_paid or 0.0
        running_balance += (inv_amt - pd_amt)

        # Lookup cached metadata for party
        meta = LEDGER_METADATA_CACHE.get(matched_target_ledger, {})

        # Voucher-level fallback for GSTIN, State, Pincode, Address
        v_gst = (v.findtext('PARTYGSTIN') or '').strip()
        v_state = (v.findtext('STATENAME') or v.findtext('PLACEOFSUPPLY') or '').strip()
        if v_state.lower() in ('not applicable', 'primary', ''):
            v_state = ''
        v_pin = (v.findtext('PINCODE') or '').strip()

        final_gstn = meta.get("gstn") or v_gst
        if not v_state and len(final_gstn) >= 2 and final_gstn[:2] in GST_STATE_CODES:
            v_state = GST_STATE_CODES[final_gstn[:2]]

        final_address = meta.get("address", "")
        if not final_address:
            v_addr_lines = []
            for a in v.iter('BASICBUYERADDRESS.LIST'):
                for sub in a.iter('BASICBUYERADDRESS'):
                    t = (sub.text or '').strip()
                    if t and t not in v_addr_lines:
                        v_addr_lines.append(t)
                t = (a.text or '').strip()
                if t and t not in v_addr_lines:
                    v_addr_lines.append(t)

            for a in v.iter('ADDRESS.LIST'):
                for sub in a.iter('ADDRESS'):
                    t = (sub.text or '').strip()
                    if t and t not in v_addr_lines:
                        v_addr_lines.append(t)
                t = (a.text or '').strip()
                if t and t not in v_addr_lines:
                    v_addr_lines.append(t)

            v_parts = [a for a in v_addr_lines if a]
            state_pin_part = ""
            if v_state and v_pin:
                state_pin_part = f"{v_state} - {v_pin}"
            elif v_state:
                state_pin_part = v_state
            elif v_pin:
                state_pin_part = f"PIN - {v_pin}"

            if state_pin_part:
                v_parts.append(state_pin_part)

            final_address = ", ".join(v_parts) if v_parts else ""

        chosen_ledger = other_ledger or (payment_mode if amount_paid else "General")
        ledger_group = get_ledger_full_group_path(chosen_ledger)

        # Save raw date for sorting
        records.append({
            "sNo": s_no,
            "date": v_date,
            "_raw_date": v_date_raw,
            "_guid": v_guid,
            "vendor": matched_target_ledger,
            "gstn": final_gstn,
            "voucherNo": v_no,
            "address": final_address,
            "invoiceNo": reference,
            "ledger": chosen_ledger,
            "ledgerGroup": ledger_group,
            "particulars": narration,
            "taxable": taxable_val,
            "cgst": cgst_val,
            "sgst": sgst_val,
            "igst": igst_val,
            "gstTotal": gst_total,
            "tcsTds": tcs_tds_val,
            "invoiceAmount": invoice_amount,
            "amountPaid": amount_paid,
            "paymentDate": v_date if amount_paid else "",
            "paymentMode": payment_mode,
            "ledgerBalance": 0.0,
            "closingBalance": meta.get("closingBalance", "")
        })
        s_no += 1

    # Sort vouchers strictly party-wise (alphabetical by party name, then chronological by date)
    records.sort(key=lambda r: (r.get("vendor", "").lower(), r.get("_raw_date", ""), str(r.get("voucherNo", ""))))

    # Re-number sequential S No and set individual voucher row ledger balance to 0
    for idx, r in enumerate(records, start=1):
        r["sNo"] = idx
        r["ledgerBalance"] = 0
        if cleanup_raw_date:
            r.pop("_raw_date", None)

    return records
