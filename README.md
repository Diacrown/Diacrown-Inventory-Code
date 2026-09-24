# BULK PI GENERATOR

Purchase Import — Bulk File Generator, for **Natural (diamond), Gemstone,
and Lab-Grown** line items in one tool. Works for any location's import
file — it isn't tied to one office.

A single-file, no-backend HTML tool. Drop in a raw purchase import export,
it corrects it in the browser, and you download the fixed `.xlsx` straight
back out. Nothing is uploaded anywhere — all parsing happens client-side via
SheetJS.

This is the same engine as the earlier gemstone-only **GS BULK PI** tool,
extended with a second logic path for diamonds (further split into Natural
and Lab-Grown for display). The underlying code logic is always automatic —
a row's `Type` decides which logic builds its Item#, never a manual switch.
What *is* a manual choice is which screen you land on: see **Pick a mode
first** below.

## What it does

On every real line-item row (rows with a `Pos` value — trailing grand-total
rows are dropped from the output):

1. **SBK** is forced to `A` — always, unconditionally, on every row.
2. **SSP and Forwarding PC are synced**, not forced in a fixed direction:
   whichever of the two columns actually has a value fills the blank one
   (the whole cell is copied so type/format match exactly). If both already
   have a value and they match, neither is touched. If both have a value
   and they **differ**, neither is touched either — the row is flagged with
   a **"SSP/Fwd PC conflict"** note (and counted in its own stat card) so it
   gets a human look, rather than guessing which side is correct. If both
   are blank, there's nothing to sync.
3. **Item # renaming**, only for rows where `IT` (Item Type) = `PP`:
   - Skipped (left as the original Item#) if `Ref #` or `Lot#` contains a
     name — that marks the line as reserved for an overseas client.
   - A bare `STOCK` or `STK` in `Ref #`/`Lot#` does **not** count as a
     client name — those rows still get renamed as ordinary stock.
   - Otherwise, the row's **`Type`** decides which of the logic paths below
     builds the new Item#.
4. **Format normalization**, **2-decimal rounding**, **manual override**,
   and **History** all work exactly as before (unchanged from GS BULK PI —
   see the last section of this document for the full detail on each).

These rules (3 and the dispatch below) run identically across all three
categories — SBK forcing, the SSP/Forwarding PC sync, and the
overseas-client skip are not gemstone-specific, they apply to the whole
file. This matches the India Diamond imports SOP, where either SSP or
Forwarding PC may be the one populated in the raw export.

### Category dispatch (automatic, per row) — three-way bifurcation

The tool reads each row's `Type` and sorts it into exactly one of three
categories:

- **Natural** — `WHD`, `BRD`, `TCD`, `FCD` (mined diamonds, color+clarity or
  treated-color graded).
- **Lab-Grown** — `LGD` only. Structurally it's built the same way as
  Natural, but it's kept as its own category everywhere in the UI (tab,
  badge, screen theme) rather than folded into Natural, so a lab-grown line
  is never mistaken for a mined one at a glance.
- **Gemstone** — everything else (`SAP`, `RUB`, `EME`, and the rest),
  unchanged from GS BULK PI.

Every row gets a **Category** badge (`natural` / `lab-grown` / `gemstone`)
in the results table, and the table itself can be filtered to just one
category at a time using the tabs above it (see below) — so a file with a
mix of all three never reads as one undifferentiated block of rows.

### Pick a mode first (landing screen)

Opening the tool shows one question before anything else: **"Which are you
working with?"** — three tiles, **Natural**, **Gemstones**, **Lab-Grown**
(plus a smaller "Not sure / mixed file" link below them). Picking one:

- Themes the whole upload screen to match (white / light green gradient /
  light purple gradient), so it's visually obvious the whole time which
  section you're in — not just after processing.
- Swaps the upload panel's copy to name that category specifically.
- Shows a small **mode chip** (e.g. "NATURAL MODE") with a **change** link
  next to it, so you can jump back to the picker at any point — this clears
  whatever file/results were loaded, but doesn't touch anything already
  downloaded.

This is a navigation choice, not a processing restriction: whichever mode
you're in, **every row still gets coded with its own correct logic** —
picking "Natural" doesn't stop a Gemstone row in the same file from getting
its Gemstone code. The mode only decides what you see by default and
which screen you're looking at, so there's no ambiguity about which
category you're currently reviewing. If a processed file turns out to
contain rows from a different category than the one you picked, a warning
banner says so explicitly (see below) rather than leaving you to notice on
your own.

"Not sure / mixed file" skips the theming and opens the tool exactly as it
worked before this landing screen existed — the tool's normal dark panel,
**All** categories together, everything auto-detected and auto-routed with
no upfront choice needed.

### Three screens, one table

After processing, the results table opens already filtered to the mode you
picked (Natural/Gemstones/Lab-Grown), or to **All** if you chose the mixed
option. Tabs above the table — **All**, **Natural**, **Gemstones**,
**Lab-Grown** — each with a live count — let you move between categories
from there. Selecting a tab re-themes the whole "Review the changes" panel
around it, matching the same palette as the landing screen:

| Tab | Panel background |
|---|---|
| Natural | Plain white |
| Gemstones | Light green gradient |
| Lab-Grown | Light purple gradient |

Switching tabs only changes what's displayed — it doesn't reprocess
anything, and edits made to a row's Item# (see **Manual override** below)
are preserved no matter which tab you're viewing when you make them or
switch away from afterward.

### Mode-mismatch warning

If you picked a mode upfront and the file you process turns out to also
contain rows from a *different* category, a banner appears above the stats
(e.g. *"You're in Natural mode, but this file also has 12 Gemstone rows...
just make sure you uploaded the right file into this section"*). Every row
is still processed correctly regardless — this is purely a heads-up so a
wrong-file upload doesn't slip through unnoticed and get mixed into the
wrong import batch. It goes away on its own once there's nothing left to
flag (e.g. a genuinely single-category file, or when you're in "Not sure /
mixed file" mode, where it never shows at all).

### Gemstone logic (unchanged)

Renamed to `{Type}{ShapeCode}{Col}{MeasurementCode}`:

- **Shape code** — looked up from the standard shape table. `OCT`
  (octagonal) is coded as **Emerald (15)**, not Octagonal (43) — a
  deliberate gemstone-industry convention, not a table lookup bug.
- **Measurement code**, parsed from `Remarks`:
  - Round shapes: first number in `Remarks`, decimal point stripped,
    padded/truncated to exactly 2 digits (`2` → `20`, `2.5` → `25`,
    `6.35` → `63`).
  - Non-round shapes: integer part of the first two numbers in `Remarks`,
    concatenated (`7 X 5 MM` → `75`).
- If the shape code or measurement can't be resolved, the row is left
  unrenamed and flagged **needs review** rather than guessed at.

### Natural & Lab-Grown (diamond) code logic — new

Built from `PP_Num_Logic_2018.xlsx` and validated against real Poland stock
codes before going into this tool. Only **Round-shape** diamonds are coded
automatically; every other shape (Princess, Marquise, Oval, etc.) is left
unrenamed and flagged **needs review**, because the fancy-shape size table
in the reference file hasn't been validated yet — this is a deliberate v1
scope limit, not an oversight. Natural and Lab-Grown share this exact same
code-building logic (LGD is structurally a diamond) — they're only split
apart for display, as described above.

For Round shapes, the new Item# is built from the `Shape`, `Size`, `Col`
(Color), and `Cla` (Clarity) columns:

- **Shape code (2 digits)** — the literal `PP_Num_Logic` Shape-table code
  (Round = `10`), looked up the same way as gemstones but **without** the
  Octagonal→Emerald override, since that override is gemstone-specific.
- **Size code (3 digits)** — the raw `Size` column in transaction files
  holds a short **CD_CODE** (e.g. `P02`, `P06`, `011`, `012`), not the
  final embedded code. The tool translates it through the `PP_Num_Logic`
  RND Sizes table (`M02→000`, `P02→002`, `P06→006`, `STM→004`, `11→008`,
  `12→010`, `ELF→013`, `14→014`, and literal 3-digit pass-through for
  larger sizes). An unrecognized size code is left unrenamed and flagged,
  not guessed.
- **WHD / BRD / FCD / LGD** (untreated / lab-grown, color+clarity graded):
  add a **Color digit** (`D/E→0`, `F/G/H/TW→1`, `I/J/WH→2`, `K/L/M→3`,
  `N/O/P/Q→4`) and a **Clarity digit** (`FL/IF/VVS1/VVS2→0`, `VS1→1`,
  `VS2→2`, `SI1→3`, `SI2→4`, `SI3→5`, `I1→6`, `I1-→7`, `I2→8`, `I3→9`).
  Final code: `ShapeCode + SizeCode + ColorDigit + ClarityDigit` (7
  digits), e.g. `1000212`.
  - **LGD only**: the code is additionally prefixed with `LG`
    (`LG1000212`). **Confirmed convention** — LGD gets the exact same
    Shape+Size+Color+Clarity code a Natural diamond would, with `LG` added
    on top, keeping codes uniform across categories rather than a separate
    numbering scheme.
- **TCD** (treated colored diamonds): no clarity digit. Instead of a
  color+clarity digit pair, the `Col` value is mapped to a fixed two-letter
  treated-color code — **`BL` for black** (covers `BL`/`BLK`/`TBLK`/`BLACK`
  source spellings — `TBLK` is data-entry drift, not a real code), plus
  `TB`/`TP`/`TY`/`TR`/`TV`/`TG`/`TD` for the other treated colors. Final
  code: `ShapeCode + SizeCode + ColorLetters`, e.g. `10006BL`.
- Any unrecognized color or clarity value leaves the row unrenamed and
  flagged, rather than guessed.

Cells the tool doesn't touch are left completely byte-identical — it edits
specific cells in place rather than rebuilding the sheet, so things like
column widths and every other column's data are untouched.

## Using it

Just open `index.html` in a browser (locally, or hosted anywhere static).
Pick **Natural**, **Gemstones**, or **Lab-Grown** on the first screen (or
"Not sure / mixed file" to skip straight to the old all-in-one view). Drop
in the raw file, hit **Process file**, review the per-row status table —
use the **All / Natural / Gemstones / Lab-Grown** tabs above it to focus on
one category at a time, and edit any Item# you want to override by hand in
any of them — then **Download corrected file**.

The `Size` and `Cla` columns are only required for files that contain
diamond rows — a gemstone-only file with no diamond `Type` values works
exactly as it did in GS BULK PI, with no changes needed to its headers.

## History — setup (optional)

Off by default. The tool works fully without it — history just won't save
or be searchable until you wire up Firebase. Same pattern as the
customs-clearance composer: Firestore and Auth's REST APIs called directly
with `fetch`, no SDK, so the tool stays a single HTML file.

**If you already set this up for GS BULK PI, reuse that same Firebase
project and paste in the same Project ID / API key** — this tool's config
constants start blank in this copy, so history won't work until you do.
Reusing means the existing history stays queryable from one place; starting
a fresh project keeps the two tools' history fully separate. Your call.

1. **Firebase console** → (existing project, or **Add project** → skip
   Analytics → Create).
2. **Build → Firestore Database** → Create database if this is a new
   project (test mode is fine to start).
3. **Build → Authentication → Get started → Sign-in method** → enable
   **Email/Password**.
4. **Authentication → Users → Add user** → `diacrownmedia@gmail.com` and a
   password of your choice. That's the one account this tool signs in as —
   the password lives only in your head and Firebase, never in the code.
5. **Firestore → Rules** — add this collection to whatever's already there
   (don't replace an existing rule block if reusing a project). This
   requires a signed-in request, not just anyone with the API key:
   ```
   rules_version = '2';
   service cloud.firestore {
     match /databases/{database}/documents {
       match /outputs/{docId} {
         allow read, write: if request.auth != null;
       }
       // keep any other existing match blocks here too
     }
   }
   ```
   Publish.
6. **Project settings → General** → copy the **Project ID**, and the
   **Web API Key** from "Your apps" (add a Web app if there isn't one yet).
7. Paste both into `index.html`, near the top of the `<script>` block:
   ```js
   const FIREBASE_PROJECT_ID = "your-project-id";
   const FIREBASE_API_KEY = "your-web-api-key";
   ```
8. Push. The History panel now shows a sign-in form — sign in with
   `diacrownmedia@gmail.com` and the password from step 4.

### What History gives you

- **Auto-save on download** — every corrected file is stored (as base64,
  inside the same Firestore doc — well within the free tier's 1&nbsp;MiB
  document limit for files this size) alongside the invoice number, a
  timestamp, and the renamed/skipped/unchanged/review/edited counts. Saving
  never blocks or delays the actual file download — if Firestore is
  unreachable or you're signed out, you still get your file, just with a
  quiet "history save failed" note.
- **Search by invoice number** — type any part of a `TSI...` number to find
  every past run for it.
- **Re-download** any past output straight from history, byte-identical,
  no reprocessing needed.
- **Export CSV** — invoice number, date, filename, and stats for every
  entry currently listed (respects an active search), for record-keeping
  outside Firestore.
- **Delete** old entries you don't need kept.
- **Duplicate-invoice heads-up** — if you process a file whose invoice
  number already has a saved run, you'll see a banner before you download,
  so a reprocess doesn't happen silently.

### How the sign-in works

Firebase's Identity Toolkit REST API, same no-SDK approach as Firestore.
Signing in exchanges the email/password for a short-lived ID token, which
gets sent as an `Authorization: Bearer` header on every Firestore request —
that's what the `request.auth != null` rule above checks. The session is
kept in `sessionStorage` (cleared when the tab closes, never written to
disk) and refreshed automatically while it's open. Nobody without that
password can read, add, or delete anything in the `outputs` collection, even
though the Project ID and API key are visible in the page's source — those
were never the actual access control.

## Deploying

**GitHub Pages**
1. Push this repo to GitHub.
2. Repo Settings → Pages → Deploy from branch → `main` / root.
3. It'll be live at `https://<user>.github.io/<repo>/`.

## Not included / out of scope (ask if you want any of these)

- **Fancy (non-Round) diamond shapes** — Princess, Marquise, Oval, Cushion,
  etc. flow to "needs review" rather than an auto-generated code. The
  `PP_Num_Logic` reference file has a separate "Other Fancy Sizes" table
  whose relationship to the "FC Sizes" table hasn't been confirmed row by
  row yet. Send a validated example set (like the RND one used to build
  the current logic) and this can be added the same way.
- **Multiple accounts / roles** — one shared sign-in
  (`diacrownmedia@gmail.com`), not per-person accounts or permission levels.
- **Separate file storage** (e.g. Firebase Storage / S3) — output files
  live inside the Firestore doc itself. Fine at this file size; would need
  revisiting if outputs ever got much larger (tens of MB).

## Notes

- Only processes the first sheet in the workbook.
- Expects the standard WSL PL Import headers (`Pos`, `IT`, `SBK`, `Type`,
  `Shape`, `Ref #`, `Remarks`, `Col`, `Item#`, `Lot#`, `SSP`,
  `Forwarding PC`) — if any are missing it'll show an error naming which
  ones, rather than guessing. `Txn #` is used for History but isn't
  required if History is off. `Size` and `Cla` are needed only for files
  containing diamond rows.
- With History off (the default), no data leaves the browser at all.
