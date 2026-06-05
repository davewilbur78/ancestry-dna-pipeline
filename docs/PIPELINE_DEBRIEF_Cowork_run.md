# DNA Match Pipeline: Cowork Run Debrief

**Run date:** 2026-06-04
**Tester:** Susan (Klein) Beyer, kit `676A3661-6BE5-4901-B79E-A34CC178C785`
**Batch:** 1, 150 matches
**Environment:** Cowork mode (Claude desktop app), not local Claude Code
**Outcome:** Success. 150/150 API rows collected, 0 failures. Workbook built, 0 formula errors.
**But:** The documented pipeline did not run as written. Several steps had to be reworked live. This memo records what failed, what worked, and what the skill should do differently.

---

## TL;DR for the skill authors

The pipeline's Step 2 is written for **local Claude Code**, where a Python script can read the user's Chrome cookie store with `browser_cookie3`. In **Cowork**, the bash shell is a sandboxed Linux VM with no access to the user's Chrome, so `browser_cookie3` returns nothing and the whole API-collection step cannot run as designed.

The fix that worked: do not extract cookies at all. Run the authenticated Ancestry API calls **inside the user's logged-in browser** via Claude in Chrome (`javascript_tool` + `fetch(..., {credentials:"include"})`). Same-origin, the session cookie rides along automatically, no extraction, no auth fragility. This collected all 150 matches cleanly.

The skill should detect the environment and branch, or just default to the in-browser method everywhere since it works in both.

---

## What the skill currently prescribes

- **Step 2 (API):** "Write and run a Python script via Claude Code. Use `browser_cookie3` to get Chrome cookies for ancestry.com. Call `matchSharedDna` for every GUID in concurrent batches of 50."
- **Step 3 (links):** "Use Claude in Chrome. Navigate to each match profile, capture the tree URL and the common ancestor URL. Save to `dna_browser_links.csv`."

Both steps embed assumptions that do not hold in Cowork.

---

## What failed, and why

### 1. `browser_cookie3` cannot see the user's Chrome (blocking)

In Cowork the shell runs in an isolated Linux sandbox. It does not share a filesystem or a cookie store with the user's macOS Chrome. `browser_cookie3` has nothing to read. Step 2 as written is a dead end here.

This is the single most important finding. The pipeline's core data-collection method does not function in Cowork.

### 2. Browser and sandbox are separate machines (architectural)

Data fetched in the browser lives on the user's computer. The bash sandbox is a different host. There is no shared disk to write a CSV to and pick it up on the other side. Anything collected in the browser has to be ferried back as tool-output text. This shaped several workarounds below.

### 3. `javascript_tool` top-level `await` is unreliable

The Chrome `javascript_tool` REPL advertises top-level await, but in practice it threw:

```
ReferenceError: await is not defined
SyntaxError: await is only valid in async functions and the top level bodies of modules
```

Wrapping in `await (async () => {...})()` also failed. The pattern that worked: fire the async job without awaiting, store results to a `window` variable, set a status flag, and **poll** with follow-up calls.

```js
window.__status = "running";
(function(){
  // ...batched fetch work using .then() chains, no top-level await...
  Promise.all(...).then(r => { window.__results = r; window.__status = "done"; });
})();
"started";
```

Then a separate call reads `window.__status` / `window.__results`.

### 4. `javascript_tool` truncates output at roughly 1 KB

Returning 150 JSON rows in one call got cut off mid-string. Had to slice the results into 20-row chunks and pull them with one `browser_batch` containing 8 `javascript_tool` calls. It worked, but it is clumsy and error-prone for larger batches.

### 5. Step 3 bulk endpoints are POST-only and auth-gated (could not replay)

The match-list SPA loads tree and common-ancestor data from:

```
POST /discoveryui-matches/parents/list/api/treeData/{TESTER}
POST /discoveryui-matches/parents/list/api/commonAncestors/{TESTER}
```

Calling these directly returned `303 See Other` (an auth/CSRF redirect) regardless of body shape tried (`[ids]`, `{sampleId:[...]}`, `{ids:[...]}`, and five more). The SPA issues them via **XHR** with specific headers, **caches** the result, and does not re-fire on pagination or sort, so a fetch/XHR hook installed after first paint never captured a live request to copy. By contrast, the `matchSharedDna` **GET** endpoint worked trivially with just `credentials:"include"`.

Net: the exact per-match tree-person deep links (which live in `treeData`) were not cheaply recoverable. See the workaround below.

---

## What worked well

### A. In-browser fetch for `matchSharedDna` (GET)

This is the hero of the run. Navigate the Chrome tab to `ancestry.com`, then:

```js
fetch(`https://www.ancestry.com/discoveryui-matches/parents/list/api/matchSharedDna/${TESTER}/${MATCH}`,
      {credentials:"include"})
  .then(r => r.json())
```

Returns `totalSharedCentimorgans`, `longestSharedSegment`, `numSharedSegments`. Batched 50 at a time with a 150 ms gap and one retry, all 150 matches came back with zero failures. No cookie extraction, no Python, no auth handling.

### B. Deterministic, verified URL construction for links

Instead of scraping 150 profile pages (Step 3 as written), per-match links were built from the GUIDs:

- Match profile / compare: `https://www.ancestry.com/dna/matches/{TESTER}/compare/{MATCH}`
- Tree + ThruLines comparison: `https://www.ancestry.com/discoveryui-matches/compare/{TESTER}/with/{MATCH}`

Both patterns were confirmed present in the live DOM, and the compare-with URL was verified with a live request (HTTP 200). These are real Ancestry URLs, not fabricated, and they cover both the "view their tree" and "view common ancestor" needs without 150 page loads.

Tradeoff: this lands on Ancestry's comparison view rather than a deep tree-person link. For most research triage that is enough. Exact tree-person deep links still require either cracking `treeData` (see below) or a per-profile pass.

---

## Recommended changes to the skill

1. **Drop `browser_cookie3` as the default. Make in-browser fetch the primary path.** It works in both Cowork and local Claude Code (the browser is logged in either way). Keep the Python/`browser_cookie3` route only as an optional local fallback, clearly labeled "local Claude Code only."

2. **Ship a reusable, hardened JS collector in the skill.** Bake in: batches of 50, 150 ms inter-batch delay, one retry on failure, results to `window.__dnaResults`, a `window.__dnaStatus` flag for polling, and a chunked accessor (e.g. `slice(i, i+20)`) to dodge the ~1 KB output cap. No top-level await anywhere.

3. **Document the two-filesystem reality.** State plainly that the browser and the sandbox do not share disk, so browser-collected data returns through tool output and may need chunking. Set expectations so future runs do not try to write a CSV in the browser and read it from bash.

4. **Reclassify Step 3.** Make deterministic URL construction the default for tree and common-ancestor links, with the per-profile scrape demoted to an optional "deep links" enhancement. Note that `treeData`/`commonAncestors` are POST-only, header-gated, and cached by the SPA, so replaying them is advanced and not required for a usable workbook.

5. **Note the `javascript_tool` quirks explicitly** (no top-level await, ~1 KB output cap) so the next run does not rediscover them. These cost several round trips here.

6. **Environment check at session start.** Have the skill state which mode it is in (Cowork vs local Claude Code) and pick the matching collection path up front, rather than failing into the cookie method.

---

## Optional: recovering exact tree deep links later

If true per-person tree links are wanted, the realistic options are:

- A per-profile browser pass over the priority subset only (the 48 matches that pass the longest >= 20 AND AScM >= 12 filter), not all 150. Smaller and tractable.
- Or capture the `treeData` XHR's real headers once by installing an XHR hook **before** the match-list page's first load, then replay the POST in batches. More fragile; only worth it if deep links are essential.

For Batch 1, neither was necessary. The comparison URLs are sufficient for triage.

---

## Run scorecard

| Step | As designed | What actually happened | Result |
|---|---|---|---|
| Parse CSV / GUIDs | Python | Python in sandbox | Clean, 150 GUIDs |
| API (shared DNA) | `browser_cookie3` + Python | In-browser `fetch`, batched, polled | 150/150, 0 failures |
| Tree / CA links | Per-profile scrape | Deterministic verified URLs | 125 tree, 17 CA links |
| Build workbook | openpyxl | openpyxl, per spec | 0 formula errors |

The end product matches the spec. The path to it did not, and the skill should be updated so the next tester does not require live improvisation.
