---
name: datamask
description: Anonymise client-identifying information (company/entity names, individual names, addresses, emails, phone numbers) in Word, Excel, PowerPoint, and PDF documents - for sharing files externally, building training/sample libraries, or internal review without exposing confidential client data. Use this skill whenever the user asks to anonymise, redact, mask, de-identify, or scrub client details from a .docx/.xlsx/.pptx/.pdf file, or mentions replacing client/entity/individual names with placeholders before sharing or storing a document - even if they don't say "DataMask" by name. Also use when the user uploads a Statement of Work (SOW), engagement letter, or otherwise asks to set up a new client engagement/masking job - build the known-entities CSV from it first (see step 0), before any files are masked.
---

# DataMask

Detects and replaces client-identifying text in documents with consistent
placeholders (e.g. "Acme Corp Pte Ltd" -> "Client 1"), and produces a
mapping log for traceability. Supports .docx, .xlsx, .pptx, and .pdf.

## Requirements

Core: `python-docx`, `openpyxl`, `python-pptx`, `pymupdf` (imported as
`fitz`), `Pillow` (used for image thumbnails in the review screen and for
generating the blank placeholder used when redacting a flagged image).
Check these are installed before running anything (`python -c "import
docx, openpyxl, pptx, fitz, PIL"`) and install what's missing with pip.

Optional but recommended: `spacy` + its `en_core_web_sm` model - powers
name/address/organisation auto-detection (see step 2). Everything still
works without it, just with narrower auto-detection (regex only).

Optional: `pytesseract` (Python package) + the `tesseract-ocr` system
binary - powers local OCR for scanned/image-only PDF pages (see the OCR
section below). Everything still works without it; scanned PDF pages
just contribute no text, same as before OCR support existed. Deliberately
a LOCAL/offline OCR engine, not a cloud one - see that section for why.

## The one rule that matters most

**Never write a masked file without the user reviewing what will be
changed first.** This skill is built around a scan -> human review ->
apply workflow, not silent auto-replacement. Skipping the review step
defeats the entire purpose of the tool: a wrong or missed replacement in
a confidentiality tool is worse than doing nothing, because it creates
false confidence that the document is safe to share.

Concretely, this means: always run `scan`, always show the user what was
found before touching anything, and only run `apply` after they've
confirmed. If the user says something like "just anonymise this and send
it back" without engaging with the review step, still show them the
findings briefly (even as a fast "here's what I found and what I'll
replace it with, say stop if anything looks wrong") rather than skipping
straight to `apply`.

## Workflow

### 0. Starting a new engagement: build the known-entities CSV from an SOW (optional, if one was provided)

If the user uploads a Statement of Work, engagement letter, or similar
new-engagement document (rather than jumping straight to files that need
masking), use it to build the known-entities CSV BEFORE asking them to
upload anything else:

1. Read the SOW and identify: the client's legal name, any subsidiaries/
   related entities it mentions, and any individuals, registration
   numbers, or addresses named in it (signatories, registered office,
   etc.).
2. If the user hasn't already told you what placeholder name to use for
   the client (e.g. "Novotech" -> "NeoPath"), ask for one before building
   anything - the whole CSV hinges on it.
3. Build the CSV starting from the core bare-word substitution:
   ```
   value,category
   Novotech,=NeoPath
   ```
   Then extend it per the same rules used throughout this file: prefer
   the bare substitution over separate full-name rows unless a specific
   name would collide/garble (see the "Custom" `=Literal` convention
   below); add individuals with a `=Literal` placeholder name; add
   registration numbers as a real-shaped fake one
   (`202221075M,=104332181S`); only add explicit address rows if
   automatic address detection (see step 2) doesn't already handle them
   correctly once you test against a real document.
4. Show the user the draft CSV and ask them to confirm/edit it BEFORE
   using it against any real documents - same "review before acting"
   principle as the rest of this skill, just one step earlier in the
   process. Don't silently assume the SOW mentions every entity that
   will show up in the actual documents later; treat this as a strong
   starting point to refine once real files are scanned, not a final
   answer.
5. Once confirmed, proceed to the normal scan -> review -> apply workflow
   below using this CSV, as the user uploads the actual documents to
   mask.

If no SOW is provided, skip straight to step 1 and build the
known-entities CSV from whatever the user tells you directly instead.

### 1. Read the document and detect entities yourself (LLM detection)

Before running anything, extract the document's plain text and read it:

```bash
python scripts/datamask_cli.py extract-text INPUT_FILE
```

This is a pure read - no known-entities list needed, nothing is modified.
For multiple files, pass them all in one call; for a very large document,
add `--out-dir DIR` to write `.txt` file(s) instead of dumping everything
into the conversation at once, and read those in pieces.

Then, using your own judgement (not just pattern-matching), identify
everything that looks like client-identifying data:
- company/entity names, fund names, deal/project names
- individual names
- physical addresses
- phone numbers and email addresses (regex will also catch most of
  these in the next step as a backup, but list obvious ones anyway)
- anything else that's clearly tied to a specific client and wouldn't
  belong in a sanitised/training-library copy of the document

Turn what you find into a known-entities CSV (format below) and write it
to a file, e.g. `detected_entities.csv`. A few things matter here:
- Use the **exact spelling/casing as it appears in the document** for
  the `value` - matching is case- and whitespace-flexible from there, so
  you don't need to list every casing variant of the same entity.
- Matching is still **exact-words, in order** - it is not alias-aware.
  If a party is referred to both by full name and by a short-form alias
  or defined term (e.g. "Everstone Capital Limited" and separately
  "ECL"), those are two different literal strings and need two separate
  CSV rows if you want both masked - don't assume catching one covers
  the other.
- If the document already implies a numbering/codename convention (e.g.
  it's file N of a series for the same client), use pre-numbered
  placeholders (`Entity 1`, `Individual 2`, ...) consistently; otherwise
  a plain category (`Entity`, `Individual`, `Address`) is fine and `scan`
  will auto-number it.
- If the user already gave you their own known-entities CSV, merge your
  detected rows into it rather than replacing it - keep their rows and
  placeholders as-is for anything overlapping, and add your own findings
  for anything they didn't already list.
- Bias toward flagging borderline items rather than silently omitting
  them (a name that could be a person or could be a place, an ambiguous
  reference number). The review step in step 3 is exactly where the user
  filters these out - a missed identifier is worse than one extra item
  they untick.

This step is a detection aid, not a replacement for the review step
below - it just means the user no longer has to hand-build a CSV before
you can start.

### 1b. Check for logos and other text baked into images

None of the text-based detection in this skill (steps 1-2) can see
content that's part of an image rather than editable text - a letterhead
logo, a signature block scan, a screenshot with a company name visible in
it. Catching that needs an actual look at the image, not regex/NER, so
pull every embedded image out and view them yourself:

```bash
python scripts/datamask_cli.py extract-images INPUT_FILE --out-dir extracted_images
```

This writes each embedded image to disk plus an `images_manifest.json`
(each entry has an `id`, `path`, and `input_file`) and touches nothing in
the source document. View each extracted image. For any that contain a
logo, letterhead, signature, or other client-identifying visual content,
note its `id` - you'll add these to the review file in step 3 so they get
blanked out in step 4. Purely decorative images (icons, generic stock
photos, chart screenshots with no identifying text) don't need flagging.

If a document has no embedded images, this step is a no-op - just confirm
`images_manifest.json` came back empty (or wasn't written at all) and
move on.

### 2. Scan the document

```bash
python scripts/datamask_cli.py scan INPUT_FILE --known-entities detected_entities.csv --out candidates.json
```

- `--known-entities` here is the file you produced in step 1 (merged
  with the user's list if they had one).
- **spaCy NER runs automatically** - it's no longer an opt-in flag. By
  default it's used ONLY for address detection: a spaCy-based pattern
  (house/block number + street name + street-type word, extended through
  a trailing city/country) works the same for a Singapore, UK, US, or
  other address, as long as it uses a recognizable street-type word (see
  the format notes for the word lists and their limits). Requires
  `pip install spacy` and `python -m spacy download en_core_web_sm` -
  check these are installed first (`python -c "import spacy"`) before
  running scan; if missing, either install them (ask the user first) or
  proceed with `--no-ner` (falls back to the narrower Singapore-only
  address regex).
- **General name/organisation detection is OFF by default** (address-only
  mode) - confirmed real problem otherwise: spaCy's general PERSON/ORG
  detection mistags ordinary phrases in business documents as if they
  were distinct entities (seen directly: "Third-Party", "Vessel
  Management Services", a bare "Group" or "Time", each getting its own
  confusing "Entity N" placeholder that then needs individually
  unchecking in review). Since the known-entities list is usually already
  comprehensive for the entities that actually matter, this noise
  provided little value for the review-time cost of wading through it.
  Add `--full-ner` if a document has names/companies that AREN'T covered
  by the known-entities list and it's worth the extra review noise to
  catch them automatically.
- Detected addresses now collapse to the literal word **"Address"** (not
  a numbered "Address 1", "Location 2", etc.) unless the whole thing is
  recognizably just a bare country/region name, in which case that name
  is kept as-is - e.g. "Taiwan" stays "Taiwan", "32 Raffles Place,
  Singapore 048624" becomes just "Singapore", "12 The Esplanade, Perth WA
  6000" becomes "Australia" (a recognized state code resolves to its
  country), but a bare building or street name with nothing recognizable
  to extract a country from (e.g. "Manulife Tower" on its own) becomes
  the literal "Address". Deliberately not numbered: which specific
  document or position an address came from isn't meant to be inferable
  from the placeholder.
- Add `--no-pattern-matching` to skip the regex-based email/phone/UEN/ABN
  detection entirely and rely only on the known-entities list (and NER,
  unless also disabled). Useful if the document is full of invoice/
  reference numbers that keep getting flagged as phone numbers or UENs -
  offer this if the user complains about that kind of noise.

This produces a JSON file (`candidates.json` by default) - nothing in the
original document has been touched yet.

### 3. Present the findings - as an interactive review screen, not raw JSON

Don't dump the review JSON on the user or make them hand-edit it. Build a
clickable review screen instead, using the `visualize:show_widget` tool
(if it's available in this environment):

```bash
python scripts/build_review_widget.py --review candidates.json --out review_widget.html
```

If step 1b flagged any images, pass those through too so they show up in
the same screen, pre-checked with your suggestions:

```bash
python scripts/build_review_widget.py --review candidates.json \
  --images-manifest extracted_images/images_manifest.json \
  --flagged-images "word/media/image1.png,word/media/image2.png" \
  --out review_widget.html
```

Then read the generated file's exact contents and pass them as the
`widget_code` argument to `visualize:show_widget` (call `visualize:read_me`
with the `interactive` module first if you haven't in this conversation).
The screen shows every detected item grouped by category, with a checkbox
to include/exclude it, an editable text field with the suggested
placeholder already filled in, and (if provided) a thumbnail grid of
flagged images with their own checkboxes. A "mask all" toggle per category
saves clicking through a long list one row at a time. For **Individual**,
**UEN**, and **Address** items, the pre-filled placeholder is a realistic
scrambled value rather than a numbered one (e.g. "Mark Rober" -> "Jane
Ali", "6738219C" -> "5463933G", "32 Raffles Place, Singapore 048624" ->
"Singapore") - the field is still editable if the user wants something
else. Don't repeat the widget's contents back in text - per the tool's
own rules, the user can already see it.

When the user clicks "Send decisions to Claude", their choices arrive as
your next message, starting with `DATAMASK_REVIEW_SUBMIT:` followed by a
JSON payload: `{"items": [{"id": ..., "include": ..., "placeholder": ...}, ...], "images": [...], "notes": "..."}`.
Parse that JSON and apply it onto `candidates.json` (match on `id`, set
`include`/`placeholder` from the payload, and populate the `"images"` list
from the `images` array) - don't regenerate the review file from scratch,
edit it in place. If `notes` isn't empty, read it: it's the user typing
something the automatic detection missed - fold it in (add to the
known-entities CSV and re-scan, or edit `candidates.json` directly) before
moving on, and mention what you added.

**If `visualize:show_widget` isn't available in this environment**, fall
back to a plain-text summary instead: group items by category in a
markdown table (representative sample plus counts for a long list, not
every row), note the scrambled placeholders as above, list any flagged
images by file path, and ask the user to confirm or tell you what to
change. Edit `candidates.json` directly based on their reply the same way.

Either way, always get an explicit confirmation before the next step -
don't treat silence, or a vague "looks good", as clearance to skip
straight past a long list without at least a summary.

### 4. Apply

Only after the user has confirmed (via the widget or in chat):

```bash
python scripts/datamask_cli.py apply candidates.json --output-dir OUTPUT_DIR
```

This writes the masked document and a `mapping_log.json` (original ->
placeholder; flagged images aren't included in this log since there's no
meaningful "original text" to record for them). Tell the user the mapping
log is a traceability record for internal use - flag that it should be
kept separately and not shared externally alongside the anonymised file.

Any images listed under `"images"` in the review file are replaced with a
plain grey placeholder square - not a visual match for the original, on
purpose, since the point is to remove the identifying content rather than
disguise it. If the original image's frame had a very different aspect
ratio, the placeholder will stretch/crop to fill it; mention this as a
minor cosmetic side effect worth a glance, not a masking failure.

```bash
python scripts/datamask_cli.py apply candidates.json --output-dir OUTPUT_DIR
```

This writes the masked document and a `mapping_log.json` (original ->
placeholder; flagged images aren't included in this log since there's no
meaningful "original text" to record for them). Tell the user the mapping
log is a traceability record for internal use - flag that it should be
kept separately and not shared externally alongside the anonymised file.

Any images listed under `"images"` in the review file are replaced with a
plain grey placeholder square - not a visual match for the original, on
purpose, since the point is to remove the identifying content rather than
disguise it. If the original image's frame had a very different aspect
ratio, the placeholder will stretch/crop to fill it; mention this as a
minor cosmetic side effect worth a glance, not a masking failure.

## Known-entities CSV format

Two rows, `value,category`. Quote values containing commas.

```csv
Acme Corp Pte Ltd,Client
Acme Holdings International,Entity
John Tan,Individual
6738219C,UEN
"123 Raffles Place, Singapore",Address
```

Two conventions are supported for the category column:

- **Generic category** (e.g. `Client`, `Entity`) - gets auto-numbered per
  distinct entity found: `Client 1`, `Client 2`, etc. Use this when the
  user just wants categorized placeholders.
- **Exact pre-assigned placeholder** (e.g. `Name 1`, `Name 2`) - used
  verbatim, no additional numbering. Use this when a team already has a
  stable codename mapping they reuse across many documents for the same
  client (detected automatically: a category ending in a digit is treated
  as pre-numbered).

**`Individual`, `UEN`, and `Address` are special-cased** and don't follow
either convention above by default: instead of a numbered placeholder,
they get a deterministically-scrambled realistic fake value instead -
"Mark Rober" -> "Jane Ali", "6738219C" -> a different but same-shaped
reg. number, "32 Raffles Place, Singapore 048624" -> "Singapore" (the
country/city, with the street-level detail dropped rather than numbered).
Scrambling is stable for a given input (the same name always scrambles to
the same fake name across a batch), but is not reversible via a stored
mapping the way "Client 1" is - the mapping log still records original ->
placeholder either way, so traceability isn't lost. If the user wants
`Individual`/`Address` numbered instead (e.g. `Individual 1`) rather than
scrambled, use the pre-numbered convention above (`Individual 1` as the
literal category value) to opt out per-row.

`UEN` is also auto-detected without needing a CSV entry at all (see
below) - list it explicitly only if you want to catch a specific number
that doesn't match the regex shape (e.g. a foreign registration number).

A bare street/landmark name with no city/country attached (e.g. spaCy
tagging "Baker Street" on its own, with nothing nearby to tell it what
country it's in) shows up as category `Location`, not `Address` - it
still gets masked, just with a plain numbered placeholder (`Location 1`)
since there's no country for the scramble to extract.

Matching is case- and whitespace-flexible - "Acme Corp" in the CSV will
still match "ACME CORP" or "AcmeCorp" in the document - but the actual
document spelling is always what gets replaced, so different-cased
mentions of the same entity share one placeholder rather than being
treated as separate entities.

## Format-specific things worth knowing (mention proactively where relevant)

- **Word/PowerPoint**: masking a paragraph/text run can flatten formatting
  that varies *within* one sentence to the first run's style. Whole-cell,
  whole-paragraph formatting (bold headers, table styling) is unaffected.
- **Excel**: every sheet/tab is scanned, not just the active one - each
  finding's `sheet` field says which tab it came from. Formula cells are
  deliberately never scanned or touched (only literal text cells), since
  rewriting part of a formula would corrupt it. Worksheet **tab names**
  themselves aren't scanned or renamed - if a client name is used as a tab
  title, mention this as an unhandled case rather than silently missing it
  (renaming tabs safely would require rewriting cross-sheet formula
  references, which this skill doesn't do).
- **PDF**: works completely differently from the other formats - there's
  no editable text to rewrite, so it redacts (permanently removes) the
  matched text and redraws the placeholder at the same size/position/font
  family (approximated to the closest of 14 standard fonts). Recommend the
  user spot-check PDF output more carefully than other formats. Scanned/
  image-only PDFs have no extractable text and `scan` will say so plainly.
- **Phone number detection** is intentionally conservative (requires a
  clear separator, boundary checks against longer codes, excludes
  year-range shapes like "2025-2026") but can still occasionally flag an
  invoice/reference number that happens to be shaped like a phone number
  (e.g. "2504-0018" with nothing else around it) - that's an inherent
  ambiguity, not a bug, and is exactly what the review step is for.
- **UEN / registration number detection** runs automatically (no CSV entry
  needed) and covers the standard ACRA UEN shapes plus a looser
  digits-plus-check-letter catch-all, so it can occasionally flag an
  unrelated code that happens to end in a letter - same "review step
  catches it" logic as phone numbers.
- **ABN (Australian Business Number) detection** also runs automatically,
  by its distinctive 11-digit, 2-3-3-3-grouped shape ("34 009 200 686") -
  scrambled format-preserving the same way UENs are, kept as its own
  category in the review screen for clarity.
- **Address detection is NLP-based (spaCy), not a fixed regex.** It finds
  a generic street shape - house/block number + street name + a common
  street-type word, in either suffix ("... Street") or prefix ("Jalan
  ...") position - and extends it through spaCy's own place-name
  recognition to pick up the trailing city/country and any postal code,
  e.g. "32 Raffles Place, Singapore 048624" or "221B Baker Street,
  London" both come back as one address. Unlike the old regex, this
  isn't tied to one country's conventions - it works the same for a
  Singapore, UK, US, or other address, as long as the street name is
  recognized as a proper noun and the street-type word is one of the
  ones it knows. It's still not a full address parser: PO boxes,
  unit-only references, and street-naming conventions outside its
  word lists (an unusual regional term, or a language it isn't set up
  for) won't be auto-caught - lean on step 1's own reading and the
  known-entities CSV for those. A bare, unattached street/landmark name
  (no city/country nearby) shows up as its own "Location" category
  rather than "Address", and gets a plain numbered placeholder rather
  than the country-extraction treatment described below, since there's
  no country for it to extract. Falls back to the old Singapore-only
  regex only if spaCy/its model isn't installed at all.
- **Images/logos**: no format's automatic scan looks inside images - see
  step 1b (`extract-images`). Once flagged, a redacted image becomes a
  plain grey square by default (not a visual match for the original),
  and for Word/Excel/PowerPoint the replacement can stretch or crop
  slightly if its aspect ratio differs from the original - a cosmetic
  side effect worth a glance, not a masking failure. PDF image
  redaction instead draws directly over the image's exact position, so
  this doesn't apply there.

  A flagged image can instead be REPLACED with a specific logo/image the
  user supplies (e.g. swapping a client's real logo for the anonymized
  entity's own placeholder logo - "MMA" -> "Azure", their logo -> your
  Azure logo, everywhere it appears), rather than just blanked. Web app:
  upload a "Replacement logo" in the sidebar, then choose "Replace with
  logo" instead of "Blank" per flagged image in the review step. CLI:
  mark an image's review-JSON entry `"action": "replace"` (instead of
  the default `"blank"`) and pass `--replacement-logo PATH` to `apply`.
  The supplied image is auto-converted to match each embedding slot's
  own format, so any common input image format works regardless of what
  format the original logo happened to be saved in. If "replace" is
  requested but no replacement image was actually supplied, it silently
  falls back to a plain blank rather than erroring. Same stretch/crop
  caveat as blanking applies, and matters more here: check the actual
  output if the replacement logo's aspect ratio differs noticeably from
  the original's.

## Known limitation: pre-existing OCR text layers in scanned PDFs (confirmed, unresolved - read before relying on this for legacy/scanned agreements)

Some PDFs that look like ordinary digital documents - especially older
signed agreements that were scanned and OCR'd (by something else, e.g.
a DMS ingestion pipeline) long before they ever reached this skill -
already carry their own, often low-quality, invisible OCR text layer.
`extract_full_text_pdf` has no way to tell this apart from a genuine
digitally-typeset PDF: the page isn't empty and isn't dominated by one
big background image, so nothing here currently flags it as different
from a normal born-digital contract.

That matters because `apply_mapping_to_pdf` trusts PyMuPDF's reported
font size and position for each match (`_find_span_at_rect`) to size
and place the placeholder exactly, and trusts the `search_for`
rectangle to bound exactly what gets erased. A pre-existing OCR
layer's per-character positions and reported font size are only ever
approximate - good enough for text selection/search, not for
pixel-exact redaction.

**Confirmed real-world failure** on a scanned 2013 investment advisory
agreement: redaction rectangles consistently bled into the first
letter(s) of the following word on *every* occurrence of two different
entity names (e.g. "Everstone Capital Limited" -> "apital Limited",
with a barely-legible, undersized placeholder stamp in place of
"Everstone"). This was not an occasional miss - it was systematic for
that document, and the leftover fragments were still human-readable,
which defeats the point of anonymising in the first place.

**Until this has a real fix** (e.g. detecting anomalous font-size
uniformity or character-width variance across a page as a signal of a
pre-existing OCR layer, and falling back to coarser whole-line or
whole-paragraph redaction for those pages instead of trusting exact
character spans), treat any PDF that was originally a scan - even one
that already "has text" and isn't caught by the scanned/image-only
check - the same as low-trust OCR output:
- Spot-check every redaction visually, page by page, before the file
  leaves the building - don't rely on the mapping log alone to decide
  a PDF is safe to share externally.
- If you can tell a PDF started life as a scan (inconsistent character
  spacing, odd typos like "ll" for "II" or "lndia" for "India" already
  present in the *unmasked* extracted text), flag that to the user
  explicitly before running `apply` on it, and recommend manual review
  of the output as a mandatory step, not an optional one.

## OCR for genuinely scanned/image-only PDF pages

A page with literally no text layer at all (a true scan, as opposed to
the pre-existing-OCR-layer case above, which already has SOME text) used
to just contribute nothing - `scan` would report no extractable text and
that page's content was never seen at all. If `pytesseract` and the
`tesseract-ocr` binary are installed, such pages now get OCR'd locally
and their recognized text is included in both `extract-text` and `scan`,
per page - a mixed document (some born-digital pages, some scanned)
handles each page correctly rather than being all-or-nothing.

**Why local OCR (Tesseract), not a cloud OCR API** (Azure Document
Intelligence, Google Document AI, etc.): a cloud OCR service needs an API
key and network access to a third-party endpoint, and would send the
document's content to another vendor. Tesseract runs entirely offline,
at some cost in raw accuracy versus a commercial engine - a deliberate
trade-off given what this skill handles is, by definition, confidential.

**Redaction of OCR'd pages is coarser than born-digital text, on
purpose.** There's no real text layer to search/replace precisely, so
masking works at whole-LINE granularity: Tesseract's own line
segmentation groups recognized words into lines, each line's text is
checked against the known-entities mapping, and if anything matches, the
ENTIRE line's bounding box is redacted and redrawn with the substitution
applied. This blanks a little more than the exact matched span (whatever
else shares that line goes too), but that's the safer failure mode for
OCR'd content - a slightly-off word boundary from imprecise word-level
positioning could otherwise leave a partial, still-legible fragment of a
matched entity behind, which would be worse.

**OCR accuracy is genuinely imperfect - always spot-check scanned-page
output visually** (render the page and look, the same advice as the
pre-existing-OCR-layer section above). A misread character can cause a
real entity to not match at all (e.g. "Ltd" misread as "Lid" won't match
a CSV row written as "Ltd"), or can cause a "double-trigger-word" entity
override (see the Hard-won lessons section below) to miss its special
case and fall through to plain word-by-word substitution instead -
neither of these leaks the original text, but the output can look
cosmetically odd (e.g. a doubled placeholder word) in a way that's worth
noticing before the file goes out.

## Verify: a residual-leak check after apply

```bash
python scripts/datamask_cli.py verify OUTPUT_FILE [OUTPUT_FILE ...] --known-entities codenames.csv
```

Run this against the MASKED output (not the original), using the same
known-entities CSV you scanned with. It checks whether any of the CSV's
original values still literally appear in the output, and exits with
code 1 if anything does (0 if clean) - scriptable as a final gate before
the files actually leave the building, especially for a large batch
where you want a quick "did anything obviously fail to mask" answer
before spot-checking a sample by hand.

This intentionally does NOT re-run full auto-detection (regex/NER)
against the masked output - only checks the specific values you listed.
Re-running full detection against already-masked content produces a
flood of false positives, because the tool's own placeholder text can be
shaped exactly like the real thing it's disguising (a fake UEN is still
shaped like a UEN; a fake address's country name is still a real
country name) - `verify` sidesteps that entirely by checking only for
the literal original values, which is what actually answers the
question "did my masking work."

`verify` finding something doesn't necessarily mean the tool is broken -
per the Hard-won lessons below, it usually means a CSV row's exact
wording doesn't match how that entity appears in this particular file
(a different abbreviation, a typo, an OCR misread) and needs its own row
or a wording tweak, not a code fix.

## Batch use (multiple files, same client)

There's no persistent mapping store built in - each `scan`/`apply` run is
independent. For consistent placeholders across several files for the
same client, reuse the same known-entities CSV (with the pre-numbered
placeholder convention above) across each file's scan step, rather than
letting each file auto-number its own entities independently.

## Hard-won lessons from real multi-file batches (read before a big run)

These came from actually running this on real multi-file batches across
docx/xlsx/pptx/pdf. Most of them are about the known-entities CSV being
incomplete for a specific document's exact wording, where the fix is
"add a row", not "debug the tool" - **except the ones marked [Fixed]
below**, which were genuine bugs in the matching/redaction itself, now
fixed, documented here so the failure mode is recognizable if it's ever
seen again (e.g. via an older un-patched copy of this skill).

- **[Fixed] A street name could leak completely unmasked if it's named
  after a person and uses a street-type word this skill didn't
  recognize.** Confirmed on a real Malaysian address: "Lingkaran Syed
  Putra" (a real Kuala Lumpur road, named after a person - common in
  Malaysian street naming) went completely unmasked. Root cause:
  "lingkaran" (ring/loop road) wasn't in the recognized street-prefix
  vocabulary, so the street-address pattern never fired - and without
  that pattern claiming it first, "Syed Putra" fell back to being
  tagged as an ordinary PERSON entity, which address-only mode skips
  entirely (since enabling general name detection brings its own
  tradeoffs - see elsewhere in this file). Fixed by adding "lingkaran",
  "persiaran", and several other common Malay/Indonesian street/area
  words (lebuhraya, medan, kompleks, wisma, menara, kampung, ...) to the
  recognized vocabulary. If an address in another language/region leaks
  completely rather than partially, check first whether its street-type
  word (the local equivalent of "Street"/"Road") is in
  `_STREET_PREFIX_WORDS`/`_STREET_SUFFIX_WORDS` - a missing street-type
  word is a plausible root cause, not just "the model missed it."

  **Residual, NOT fixed**: a building name with a directional suffix
  (e.g. "Centrepoint South") went unmasked in the same real address, and
  isn't tagged as any entity by spaCy at all in that context - there's
  no signal for the existing filters to act on, unlike the street-name
  case above where the model detected SOMETHING and it was a matter of
  the categorization being wrong. Genuinely harder to fix generally
  (would need new detection logic, not a vocabulary/filter tweak) -
  flagged here rather than silently left as an assumed win. If a
  building name keeps leaking, it's worth adding literally as a
  known-entities-list row for that specific job rather than waiting on
  a general fix.

- **[Fixed] Broad geographic regions ("South East Asia", "Middle
  East/Africa", "Europe") fell through to the generic "Address"
  placeholder instead of staying visible like a country does.**
  Confirmed directly: this wasn't over-detection - spaCy correctly
  identifies these as genuine location references (tagged LOC, the same
  way a real street address is), which is exactly right. The gap was
  that only individual sovereign countries were recognized as "fine to
  leave visible" (see _KNOWN_COUNTRY_NAMES), not continents or the broad
  multi-country regions a financial "geographic market" breakdown
  commonly reports by instead of listing every country. Fixed with a
  matching _KNOWN_REGION_NAMES list (continents plus common business
  regions - APAC, EMEA, Middle East, South East Asia, etc.), plus
  explicit handling for the "/"-joined compound shorthand this kind of
  table often uses ("Australia/New Zealand", "Middle East/Africa") -
  recognized as a whole when every part on either side of the "/"
  resolves to a known country or region.

- **[Fixed] Short fiscal-period abbreviations ("FY", "Q1", "H2", ...)
  could get masked as a false "Address".** Same underlying class of
  problem as the "Vessel"/"Background" defined-term issue - a short,
  capitalized, context-light abbreviation is exactly the shape spaCy's
  NER is prone to mistagging, and financial documents are full of them
  ("in FY 2023", "Q1 revenue"). Added to the same _COMMON_DEFINED_TERMS
  blocklist used for that fix, so it's rejected regardless of which NER
  label caused it.

- **[Changed] Output filenames no longer have an "_anonymised" suffix.**
  A masked file is now named exactly the same as the input, other than
  whatever the known-entities mapping itself renamed in the name (e.g.
  "Cyan_group_listing.xlsx" -> "Azure_group_listing.xlsx", not
  "Azure_group_listing_anonymised.xlsx"). Worth knowing: if NOTHING in a
  given file's name matched anything in the mapping, the output filename
  is now IDENTICAL to the input's - saving it into the same folder as
  the original will silently overwrite it. Save outputs to a different
  folder than the originals, or rename before saving, if that matters
  for a given batch.

- **[Fixed] A short mapping key could silently mangle an unrelated word
  it happened to appear inside of - e.g. "us" (detected as the country
  abbreviation "US" elsewhere in the document) redacting the "us" INSIDE
  "customer", turning it into "c[Address]tomer".** Confirmed directly:
  neither PyMuPDF's `search_for()` (used for PDF) nor plain Python
  string replacement (used for docx/xlsx/pptx) have any word-boundary
  awareness at all - both do pure substring matching, so a short key
  matches wherever those exact letters appear, mid-word or not. This
  was a real gap specifically at APPLY time - every DETECTION path
  (the known-entities regex, the address matcher) already enforces a
  proper word boundary, so this kind of match was never presented for
  review in the first place; it only appeared once redaction ran.
  Fixed on both fronts: PDF redaction now cross-checks each
  `search_for()` match against the page's actual word-level layout
  (`page.get_text("words")`) and only accepts it if it lines up with
  real word boundaries; docx/xlsx/pptx now replace using the same
  word-boundary-respecting pattern the known-entities matching already
  used, instead of plain substring replacement. Also added "US" as a
  recognized bare-country alias (resolving to "United States" like "UK"
  already did for "United Kingdom"), so this doesn't just get suppressed
  but resolves to the right thing when it's a genuine standalone
  mention.

- **[Fixed] A common English word used as a document heading or section
  title (e.g. "Background", "Overview", "Summary") could get mistagged
  as an address.** Confirmed directly: spaCy's NER component tags a
  capitalized, isolated heading-position word as a location surprisingly
  often - capitalization and lack of surrounding sentence context are
  the same cues that indicate a real proper noun, so the model latches
  onto them even for an ordinary common noun with nothing location-like
  about its actual meaning. Fixed with a cross-check against spaCy's
  OWN, separate part-of-speech tagger: every genuine place name tested
  (Perth, Taiwan, Esplanade, Cross Street, Manulife Tower) reliably gets
  tagged as a proper noun (PROPN) by the POS tagger, while a mistagged
  heading word gets an ordinary POS tag (NOUN, VERB, ADP, ...) - so a
  GPE/LOC/FAC-tagged span is now only accepted if every substantive
  token in it is POS-tagged PROPN. This is a general filter (eliminates
  the whole class of "capitalized common word used as a heading"
  mistakes), not a one-word patch for "Background" specifically. Also
  fixed, found while testing this: the earlier "building name mistagged
  as a person" correction (e.g. "Manulife Tower") had become dead code
  once address-only mode became the default, since the code checked
  "should I skip non-address entities in this mode" BEFORE checking
  "wait, is this actually a mistagged address" - reordered so the
  address-relevant correction always runs first.

  This was originally documented (in an earlier version of this file)
  as a known, accepted limitation rather than something fixed - legal
  documents conventionally capitalize defined terms ("the Vessel", "the
  Company", "the Charterers") to signal they have a specific
  contractual meaning, and that capitalization genuinely confuses
  spaCy's POS tagger too, not just its NER, so the PROPN cross-check
  above doesn't catch it. In practice this turned out far more severe
  than "some noise to review" - confirmed directly against a real
  maritime charter: "Vessel" is used dozens of times throughout the
  document, and because every occurrence of identical text shares one
  placeholder, a SINGLE bad detection anywhere was enough to mask EVERY
  legitimate occurrence of one of the document's most common words,
  visibly degrading the whole thing rather than being an occasional
  miss. **Now actually fixed** with a curated list of common legal/
  contract/shipping/business terms (see _COMMON_DEFINED_TERMS) that are
  rejected as entities outright, regardless of NER label or address-only
  mode - covers contract-party roles (Company, Owners, Charterers,
  Buyer, Seller, ...), maritime terms (Vessel, Charter, Master, Cargo,
  ...), document-structure words (Background, Annex, Schedule, ...),
  and generic business nouns (Property, Asset, Entity, ...). If a
  similarly common word keeps showing up as a false "Entity"/"Address"
  on some other document, it likely belongs in this list too - the
  pattern is any ordinary English noun that a legal/business document
  conventionally capitalizes as a defined term.

- **[Fixed] A PDF with a broken font/ToUnicode mapping (common in
  "print to PDF" exports from a web page using a custom font - e.g. a
  government tax portal) could produce catastrophic, scattered
  corruption across the ENTIRE document, not just a missed detection.**
  Confirmed directly against a real Danish tax-portal PDF: the page's
  "native text" wasn't empty (which would have correctly triggered the
  existing OCR fallback for scanned pages) - it was PRESENT but
  genuinely meaningless: raw control-character byte sequences, because
  the PDF's glyphs render correctly on screen but don't map to real
  Unicode characters underneath. Since the old fallback logic only
  checked "is there any text at all", it treated this page as a normal
  native-text PDF and tried to search/redact based on that garbage -
  which doesn't just fail to find real matches, it can match and redact
  effectively random positions across the visually-readable page,
  while the actual sensitive text underneath was never scanned for at
  all. A symptom easy to mistake for something else: the output looks
  severely corrupted (chunks of readable text missing/replaced at
  seemingly random points), AND known-entities substitutions that
  should obviously have fired (a company name that's right there in
  the visible text) don't fire, AND the output filename doesn't get
  renamed either - all three are the SAME underlying cause (the "real"
  text was never actually readable to begin with), not three separate
  bugs. Fixed by checking whether extracted page text is actually
  USABLE, not just non-empty - a page whose text is mostly control
  characters now correctly falls through to the same local-OCR fallback
  a genuinely scanned page already used, which reads the ACTUAL VISUAL
  characters instead of the broken underlying codes. If you ever see
  scattered, seemingly-random corruption like this again (as opposed to
  a clean, complete miss), it's worth checking whether the source PDF
  has this kind of broken text layer.

- **[Added] Tesseract OCR now auto-detects and uses every installed
  language pack, not just English.** Matters for non-English documents
  hitting the OCR fallback above (or a genuinely scanned page) - a
  document in Danish, German, French, etc. OCR'd with English-only
  language data measurably loses accuracy on that language's special
  characters (e.g. Danish æ/ø/å). Install the relevant language pack
  (e.g. `apt-get install tesseract-ocr-dan` for Danish) and it's picked
  up automatically, no code or command-line change needed - check what's
  installed with `tesseract --list-langs`.

- **[Fixed] `apply` could take a very long time (minutes, sometimes much
  longer) on a large batch, specifically for PDFs.** Confirmed and
  measured directly: PyMuPDF's `search_for()` does real per-call layout/
  glyph-position work, and the PDF apply step previously called it once
  per mapping entry PER PAGE, regardless of whether that text could
  possibly be on that page. At ~1,900 mapping entries (a large but not
  extreme multi-file batch), that measured at ~10 seconds *per page* -
  which multiplies straight through by page count and is exactly what
  showed up as the tool seeming to hang indefinitely on "Applying..."
  with no further feedback. Fixed with a cheap plain-text substring
  pre-check before ever calling `search_for()`: every candidate's text
  originally came from this same PDF's own extracted text in the first
  place (that's how it was detected during scan), so if it's not present
  in a given page's plain text, it cannot have a match there either -
  skipping it is always safe, never a missed match. Measured speedup on
  a real file: ~270x (an extrapolated ~170 seconds down to 0.62 seconds).
  Also stopped needlessly re-sorting the mapping by length on every
  single page/paragraph/cell (see the "longest match first" fix further
  below) - sorting once per file instead.

- **[Changed] Placeholder conventions, by design, favor being obviously
  fake over looking plausible.** This was a deliberate choice: a
  placeholder that could pass for real data is the wrong failure mode -
  anyone reviewing the masked document should recognize a placeholder
  at a glance, not have to double-check whether it's real.
  - **Phone numbers, UENs, and ABNs** are hashtagged, not scrambled to a
    different-but-plausible-looking number: every letter/digit becomes
    "#", spacing and punctuation left as-is (e.g. "+65 8123-4567" ->
    "+65 ####-####", "201912345A" -> "##########").
  - **Names** always come out as a "J___ D___" pattern (John Doe, Julian
    Des, Jasmine Davis, ...) - deterministic per original name (same
    input always maps to the same fake name), but constrained to always
    start with those two initials specifically, so a "J" first name
    paired with a "D" last name is recognizable on sight as a
    placeholder, not just "some fake name that happens to look real."
  - **Addresses** collapse to the literal word "Address" - styled with
    an actual grey visual treatment where the file format supports it
    (a solid grey redaction bar in PDF, a grey highlight in docx, grey
    text color in pptx, a grey cell fill in xlsx) - UNLESS the whole
    thing is recognizably just a country/major-territory name, in which
    case that name is kept as normal, unstyled text (see the full
    country list note below).
  - **Generic entities** (a company/organisation with no scrambled-value
    rule and no explicit known-entities-list row) get a letter suffix
    instead of a number - "Entity A", "Entity B", "Client A" - not
    "Entity 1", "Entity 2". Different entities still get different
    placeholders (collapsing them all to one identical word would make
    a document describing relationships BETWEEN entities unreadable),
    just without a sequence of numbers that could look like it means
    something (an ordering, a count) beyond "these are different".
  - A known-entities-list row using the `=Literal` convention always
    wins over all of the above, verbatim, exactly as before - if a CSV
    already has an explicit override for a specific UEN or address (from
    before this change), it keeps that exact value rather than
    switching to the new hashtag/grey treatment. Only fresh,
    auto-detected matches use the new conventions.

- **[Expanded] The recognized bare-country-name list now covers every
  UN member/observer state** (about 195 countries) plus a few
  commonly-referenced territories (Hong Kong, Taiwan) - previously only
  ~17 countries were recognized, so e.g. "Russia" or "France" mentioned
  bare would fall through to the generic "Address" label instead of
  staying as the country name. Also now handles a leading "the" ("the
  United Kingdom", "the Netherlands") and a few long-form aliases
  ("United States of America" -> "United States", "UK" -> "United
  Kingdom").

- **[Changed] spaCy's general PERSON/ORG detection is now off by
  default - address detection only.** Confirmed directly, repeatedly:
  general NER mistags ordinary business phrases as if they were distinct
  organisations - "Third-Party", "Vessel Management Services", a bare
  "Group" or "Time", each getting its own confusing numbered "Entity N"
  placeholder that then has to be individually unchecked in review. On a
  document already covered by a solid known-entities list, this noise
  provided little detection value for a real review-time cost. Pass
  `--full-ner` (CLI) or check "Also detect names & organisations" (web
  app) to restore the old broader behavior if a document has names or
  companies that genuinely aren't in the known-entities list yet.
  Relatedly, detected addresses now collapse to the literal, non-numbered
  word "Address" (instead of "Location 1", "Location 2", etc.) whenever
  they're not resolvable to just a bare country/region name - see the
  scan step above for the full rule.

- **[Fixed] A shorter match could silently block a longer, more complete
  match elsewhere in the same document, leaking a trailing postcode.**
  Two candidates can be correctly detected as separate, non-overlapping
  matches at SCAN time - e.g. bare "Singapore" from one part of a
  document, and "Singapore 048424" from a completely different part.
  But at APPLY time, each mapping entry is searched for independently
  across the WHOLE document/page - "Singapore" (the shorter key) also
  matches the "Singapore" substring INSIDE "Singapore 048424" elsewhere.
  If the shorter entry happened to be processed first, it claimed that
  spot and blocked the longer, more complete match from ever running
  there - leaving the postcode exposed, unmasked. Fixed by always
  processing the longest original text first, across every format
  (docx/xlsx/pptx/pdf, plus filename sanitization) - the same principle
  already used at scan time, now extended to apply time too.

- **[Fixed] A redaction "already claimed" check could falsely block
  unrelated text on the very next line.** PDF redaction tracks which
  page regions have already been claimed by a higher-priority match, to
  prevent double-stamping (see the double-masking entry below). Two
  words on CONSECUTIVE lines can have bounding boxes that touch or
  overlap by a hairline (a fraction of a point) purely from font
  line-height, even sharing no actual content or position - confirmed
  directly: "Owners" on one line and "MMA" on the very next line
  overlapped by 0.06pt, which was enough to make "MMA" get silently
  skipped entirely (never masked). Fixed by shrinking a rect slightly
  before checking for overlap, so a hairline touch no longer registers
  as a genuine collision while real overlaps still do.

- **[Fixed] A match spanning multiple visual lines could get its
  placeholder text stamped twice.** If an original value (e.g. a UEN or
  ABN wrapped across a line break in the source PDF, "34 009\n200 686")
  is searched for via PyMuPDF's search_for(), it returns ONE RECT PER
  LINE for what's logically a single match - confirmed directly.
  Previously the full placeholder text was drawn at every one of those
  rects, producing a visible duplicate (e.g. "327 648 350 327 648 350"
  stacked at two different heights). Fixed by only drawing the
  placeholder once, at the first line, while still fully erasing every
  line the match spanned.

- **[Fixed] Address detection had several real gaps, found against an
  actual BIMCO BARECON charter party.** All now fixed: (1) a determiner
  between the house number and street name ("12 **The** Esplanade")
  broke the match entirely; (2) "esplanade" (and several other common
  street-type words) wasn't in the recognized word list at all; (3) a
  building name right after a street match (e.g. "Manulife Tower")
  occasionally got mistagged by spaCy as a PERSON, giving it a
  randomized fake HUMAN name instead of being treated as part of the
  address - actively confusing output, not just a miss; (4) a
  street-address match that spaCy didn't extend through to a
  recognized city (e.g. "Perth" not tagged GPE in that sentence) now
  has a text-pattern fallback requiring confirmatory evidence (a state
  code or postcode right after) before trusting a capitalized word as a
  city - avoids the earlier, riskier version of this fix which ended up
  swallowing a building name as if it were a city; (5) a bare country
  mention with a trailing postcode on its own line (e.g. "Singapore
  048424" with no preceding street context) now gets the postcode
  folded in too, rather than leaving it exposed; (6) bare country names
  other than Singapore (Australia, New Zealand, Denmark, etc.) now stay
  recognizable as themselves rather than falling through to a generic
  "Location" label, consistent with how Singapore was already handled.

- **[Fixed] spaCy could tag an entire dense-tabular chunk as one giant
  "entity", destroying financial figures in the process.** In content
  like a vessel specification table or a rate card - cells concatenated
  into running text with no normal sentence structure - spaCy would
  occasionally tag a whole run of unrelated fields (a vessel name, its
  code, a location, a day-rate figure) as a single ORG, e.g. "MMA
  Monarch S09 Malaysia 4,500 McDermott (65 day firm)...". Masking that
  whole span wipes out numbers that were never supposed to be touched,
  while STILL not correctly masking just "the entity" the user actually
  wanted. Fixed by rejecting any NER-tagged span that contains a literal
  newline, is implausibly long (>60 characters), or contains 3+ separate
  digit-groups - all strong, low-false-positive signals that multiple
  unrelated fields got glued together rather than this being one real
  entity name. The genuine entity within it (e.g. "MMA") still gets
  caught correctly, either via its own properly-scoped NER tag elsewhere
  or via a known-entities-list bare-word match.

- **[Added] ABN (Australian Business Number) detection.** Previously
  only Singapore-shaped UENs were auto-detected; an ABN like
  "34 009 200 686" either went unnoticed or got swallowed into the
  tabular-garbage bug above. Now detected on its own (by its distinctive
  11-digit, 2-3-3-3-grouped shape) and scrambled format-preserving, the
  same way UENs are.

- **[Fixed] Combining a known-entities CSV with spaCy NER could
  double-mask overlapping text, badly corrupting output - especially
  PDFs.** If a known-entities row matches part of a longer phrase that
  spaCy ALSO independently tags as its own entity (e.g. a bare
  trigger-word row "MMA" matches inside "MMA Clean Energy Co., Ltd.",
  which spaCy separately tags whole as an ORG), both used to get treated
  as independent candidates. At apply time each is redacted and redrawn
  separately - and for PDFs, two overlapping replacement texts get
  stamped on top of each other, producing garbled, unreadable output.
  A related variant: PyMuPDF's `search_for()` is case-insensitive by
  default, so even a NON-overlapping NER fragment detected elsewhere
  (e.g. bare "LTD" spotted once in a comparables table) would get
  globally re-applied everywhere that text occurs case-insensitively -
  including inside an unrelated, already-correctly-masked company name
  elsewhere in the same document (turning "Azure Clean Energy Co., Ltd."
  into the nonsensical "Azure Clean Energy Co., Entity 19"). Both are
  fixed as of this version: known-entities-list matches now take
  absolute priority and claim their exact position (longest match
  first), any other source's candidate that overlaps a claimed span is
  dropped entirely, PDF redaction additionally tracks already-claimed
  page regions as a second line of defense, and bare corporate-suffix
  fragments ("LTD", "CO", "INC", "PTE", etc. alone) are no longer
  accepted as standalone NER candidates at all, since they're a known
  NER mis-segmentation artifact, not a real independent entity. If
  you're ever running an older un-patched version of this skill, this
  bug is very hard to miss - the output is visibly garbled with
  overlapping/doubled text, not a subtle miss - and `verify` (see
  below) also won't catch it, since the original text genuinely IS
  gone, just replaced with a corrupted mess instead of one clean
  placeholder. If you see this: get the current version of the skill.

- **Legal boilerplate can still produce some NER noise, though much
  less than before.** A charter party or other heavily templated legal
  document mentions many countries in the abstract (e.g. a standard
  war-risks clause listing example countries like "Russia" or "the
  United Kingdom") - these get masked too (safely - nothing leaks), so
  expect a slightly noisier review list on this kind of document than
  on a plain business letter. Generic contract-role words like "Owners"
  or "Charterers" are now filtered out directly (see the
  _COMMON_DEFINED_TERMS fix above) rather than needing to be unchecked
  each time, but if some OTHER heavily-capitalized defined term specific
  to a document's subject matter still shows up as a false positive,
  that list is the place to extend.

- **Matching requires the same words, not just the same meaning.**
  Case/whitespace-flexible, not synonym-flexible: a CSV row for
  "Foo Holdings Pty Limited" will NOT match "FOO HOLDINGS PTY LTD" in a
  document that abbreviates "Limited" as "Ltd" - that's a different word,
  not just different case. Watch for this especially in legal agreements,
  which often write out entity names differently (full legal form) than
  an internal entity-listing spreadsheet does (abbreviated form). If the
  same entity shows up under two different exact wordings across your
  batch, it needs two CSV rows, both pointing at the same placeholder.

- **An entity name containing two separate trigger words needs its own
  explicit row, or you'll get double-garbled output.** If your codebook
  has bare "Foo,=Placeholder" and "Bar,=Placeholder" rows (a common
  pattern when a group has a country/color-codename scheme), and some
  entity is literally named "Foo Bar Holdings Ltd" (containing BOTH
  trigger words), each will match and get replaced independently -
  producing "Placeholder Placeholder Holdings Ltd". Give any such entity
  its own explicit row that overrides both bare substitutions at once.

- **A short-form or split-across-cells address needs its own row - it
  won't inherit from a longer address row that contains it.** Spreadsheets
  often repeat just the city or street name alone in a separate "Office"
  or "Location" column (e.g. bare "Perth" or "Kuala Lumpur"), distinct
  from the fuller "12-14 The Esplanade, Perth, WA 6000, Australia" used
  in a formal entity listing. These are different literal strings and
  each needs its own row if you want both masked. A quick way to catch
  these: after a first pass, grep the extracted text for country/city
  names your CSV already maps and see if any bare, unmapped occurrences
  turn up.

- **Multiple adjacent CSV-mapped fragments in one address can produce
  repeated placeholder text** (e.g. "Singapore Singapore, Singapore" when
  a street line, a building name, and a postal-code line each
  independently map to "Singapore"). This is a cosmetic readability quirk,
  not a data leak - every fragment IS masked, just not merged into one
  clean replacement. Worth a mention to the user, not worth blocking
  delivery over.

- **PDF redaction can leave a visual overlap artifact** where the
  replacement text partially overlaps the edge of the redaction box,
  especially when the placeholder is a different length than the
  original. This doesn't leak the original text (redaction genuinely
  removes it), but always spot-check PDF output visually (render a page
  and look, don't just extract text) rather than trusting extracted text
  layout - PyMuPDF's text extraction order for redacted+redrawn content
  doesn't always match the original reading order, even when the visual
  result is correct.
