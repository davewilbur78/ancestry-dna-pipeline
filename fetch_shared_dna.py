#!/usr/bin/env python3
"""
fetch_shared_dna.py -- Ancestry matchSharedDna API collector (kit-agnostic).

Collects unweighted cM, longest segment, and segment count for a list of match
GUIDs against a tester kit, using Chrome session cookies for authentication.

The tester kit is NOT hardcoded. Supply it as --tester-guid, or as --kit-url and
the GUID is derived from it. Match GUIDs come from --input-csv (a Genealogy
Assistant export; GUIDs are extracted from its URL column) or from --guids-file
(one GUID per line).

Examples:
  python fetch_shared_dna.py \\
      --kit-url "https://www.ancestry.com/dna/matches/4FB3190E-37B7-401F-A8CF-431950413661/list" \\
      --input-csv "/path/to/genealogy_assistant_export.csv" \\
      --output "/path/to/dna_api_results.csv"

  python fetch_shared_dna.py \\
      --tester-guid 4FB3190E-37B7-401F-A8CF-431950413661 \\
      --guids-file matches.txt \\
      --output results.csv
"""

import argparse
import csv
import re
import sys
import time

GUID_RE = re.compile(r"([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-"
                     r"[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})")


def derive_tester_guid(args):
    """Return the tester GUID from --tester-guid or --kit-url."""
    if args.tester_guid:
        m = GUID_RE.search(args.tester_guid)
        if not m:
            sys.exit("ERROR: --tester-guid does not look like a GUID.")
        return m.group(1).upper()
    if args.kit_url:
        m = GUID_RE.search(args.kit_url)
        if not m:
            sys.exit("ERROR: could not find a tester GUID in --kit-url.")
        return m.group(1).upper()
    sys.exit("ERROR: supply either --tester-guid or --kit-url.")


def load_match_guids(args, tester_guid):
    """Return a de-duplicated list of match GUIDs (excluding the tester)."""
    guids = []
    if args.guids_file:
        with open(args.guids_file) as f:
            for line in f:
                m = GUID_RE.search(line)
                if m:
                    guids.append(m.group(1).upper())
    elif args.input_csv:
        with open(args.input_csv, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            for row in reader:
                for cell in row:
                    for m in GUID_RE.finditer(cell):
                        guids.append(m.group(1).upper())
    else:
        sys.exit("ERROR: supply either --input-csv or --guids-file.")

    # De-duplicate, preserve order, drop the tester's own GUID
    seen = set()
    ordered = []
    for g in guids:
        if g == tester_guid or g in seen:
            continue
        seen.add(g)
        ordered.append(g)
    if not ordered:
        sys.exit("ERROR: no match GUIDs found in the input.")
    return ordered


def build_session():
    print("Loading Chrome cookies for ancestry.com...")
    import browser_cookie3
    try:
        cookies = browser_cookie3.chrome(domain_name=".ancestry.com")
    except Exception as e:
        sys.exit(f"ERROR: Could not load Chrome cookies: {e}\n"
                 "Make sure Chrome is open and logged into Ancestry.")
    import requests
    cookie_jar = {c.name: c.value for c in cookies}
    csrf_token = cookie_jar.get("_dnamatches-matchlistui-x-csrf-token", "")
    session = requests.Session()
    session.cookies.update(cookies)
    session.headers.update({
        "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0.0.0 Safari/537.36"),
        "Accept": "application/json, text/plain, */*",
        "x-csrf-token": csrf_token,
    })
    return session


def fetch_one(session, tester_guid, guid, delay):
    """Fetch a single match; retry once on failure. Returns a result dict."""
    import requests
    url = ("https://www.ancestry.com/discoveryui-matches/parents/list/api/"
           f"matchSharedDna/{tester_guid}/{guid}")
    referer = ("https://www.ancestry.com/discoveryui-matches/compare/"
               f"{tester_guid}/with/{guid}/sharedmatches")
    for attempt in (1, 2):
        try:
            resp = session.get(url, timeout=15, headers={"Referer": referer})
            resp.raise_for_status()
            data = resp.json()
            return {
                "guid": guid,
                "unweighted_cm": data.get("totalSharedCentimorgans", ""),
                "longest_segment": data.get("longestSharedSegment", ""),
                "segments": data.get("numSharedSegments", ""),
            }, None
        except requests.exceptions.HTTPError as e:
            err = f"HTTP {e.response.status_code}"
        except Exception as e:
            err = str(e)
        if attempt == 1:
            time.sleep(max(delay, 0.5))
    return {"guid": guid, "unweighted_cm": "", "longest_segment": "",
            "segments": ""}, err


def main():
    p = argparse.ArgumentParser(description="Ancestry matchSharedDna collector.")
    p.add_argument("--tester-guid", help="Tester kit GUID.")
    p.add_argument("--kit-url", help="Tester kit URL (GUID derived from it).")
    p.add_argument("--input-csv", help="Genealogy Assistant CSV export.")
    p.add_argument("--guids-file", help="Text file, one match GUID per line.")
    p.add_argument("--output", default="dna_api_results.csv",
                   help="Output CSV path (default: dna_api_results.csv).")
    p.add_argument("--delay", type=float, default=0.15,
                   help="Delay between requests in seconds (default 0.15).")
    args = p.parse_args()

    tester_guid = derive_tester_guid(args)
    guids = load_match_guids(args, tester_guid)
    total = len(guids)
    print(f"Tester GUID: {tester_guid}")
    print(f"Fetching shared DNA data for {total} match GUIDs...\n")

    session = build_session()
    results = []
    succeeded = 0
    failures = []

    for i, guid in enumerate(guids, 1):
        result, err = fetch_one(session, tester_guid, guid, args.delay)
        results.append(result)
        if err is None:
            succeeded += 1
            print(f"[{i}/{total}] OK   {guid}  -  "
                  f"{result['unweighted_cm']} cM, {result['segments']} seg(s), "
                  f"longest {result['longest_segment']}")
        else:
            failures.append(guid)
            print(f"[{i}/{total}] FAIL {guid}  -  {err}")
        if i < total:
            time.sleep(args.delay)

    with open(args.output, "w", newline="") as f:
        writer = csv.DictWriter(
            f, fieldnames=["guid", "unweighted_cm", "longest_segment", "segments"])
        writer.writeheader()
        writer.writerows(results)

    print("\n--- Summary ---")
    print(f"Total:     {total}")
    print(f"Succeeded: {succeeded}")
    print(f"Failed:    {len(failures)}")
    if failures:
        fail_rate = len(failures) / total
        print("Failed GUIDs: " + ", ".join(failures))
        if fail_rate > 0.10:
            print(f"WARNING: {fail_rate:.0%} of requests failed. "
                  "Check that Chrome is logged into Ancestry before continuing.")
    print(f"CSV saved to: {args.output}")


if __name__ == "__main__":
    main()
