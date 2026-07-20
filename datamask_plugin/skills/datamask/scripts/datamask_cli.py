#!/usr/bin/env python3
"""
CLI driver for the DataMask skill.

Wraps datamask_core.py (the same tested detection/masking logic used by the
interactive Streamlit review app) into two commands, so Claude can run the
scan -> human review -> apply workflow conversationally instead of through
a web UI:

  scan  - reads one or more documents, detects candidate sensitive items,
          writes them to a JSON review file. Does NOT modify anything yet.

  apply - reads a (possibly hand-edited) review file and produces the
          masked output file(s) + a combined mapping log. This is the only
          step that writes anonymised documents.

The split exists on purpose: DataMask's whole design principle is that
nothing gets replaced without a human seeing and approving it first. When
running as a skill, Claude is expected to run `scan`, present the results
in chat (accepted / edited / excluded), and only run `apply` once the user
has confirmed - never both steps back-to-back without that checkpoint.

Multiple files: `scan` accepts more than one input file, and placeholder
numbering is shared across all of them in one pass - so if "Acme Corp"
appears in three files being scanned together, it gets the same "Client 1"
in all three, rather than each file numbering its entities independently.
This is what makes batch use actually useful; scanning files one at a time
in separate commands would not give you this consistency.
"""
import argparse
import csv
import io
import json
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datamask_core import (
    open_document,
    extract_text_for_format,
    extract_full_text_xlsx,
    find_sheet_for_position,
    find_candidates,
    build_context_snippet,
    apply_mapping_for_format,
    save_document,
    sanitize_filename_with_mapping,
    load_known_entities_rows,
    try_load_ner,
    assign_placeholders,
    extract_images_for_format,
    redact_images_in_saved_bytes,
    redact_images_in_pdf,
    _build_flexible_entity_pattern,
    should_grey_address,
)

SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf"}


def _load_known_entities(path):
    if not path:
        return []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        rows = [r for r in reader if r and r[0].strip()]
    # A header row (e.g. "value,category") is common when the CSV was
    # built/exported from a spreadsheet. Without this check it would be
    # treated as a literal entity to match - "value" is a real word that
    # shows up constantly in ordinary business documents ("contract
    # value", "transaction value"), so silently masking it would corrupt
    # the output. Detected by the first cell being exactly "value"
    # (case-insensitive) - the one word a real entity name is never going
    # to be by coincidence.
    if rows and rows[0][0].strip().lower() == "value":
        rows = rows[1:]
    pairs = [(r[0], r[1] if len(r) > 1 else "Entity") for r in rows]
    return load_known_entities_rows(pairs)


def cmd_extract_text(args):
    """Dump each document's plain text with no detection applied yet.

    This exists so Claude can read a document's actual content directly
    and do LLM-based entity spotting (company/individual names, addresses,
    context-dependent client data that regex and even NER can miss) before
    calling `scan`. It is a pure read - nothing about the file's content
    is altered, and no known-entities list is required to run it.

    Output is written to stdout by default so it lands in the current
    conversation; use --out-dir if you want it saved to text file(s)
    instead (useful for a very large document you'd rather page through
    than dump in one go).
    """
    for input_path in args.inputs:
        ext = os.path.splitext(input_path)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            print(f"Skipped {input_path}: unsupported file type {ext}", file=sys.stderr)
            continue

        with open(input_path, "rb") as f:
            doc_bytes = f.read()
        doc = open_document(doc_bytes, ext)
        full_text = extract_text_for_format(doc, ext)

        if not full_text.strip():
            print(f"--- {input_path}: no extractable text found (scanned/image-only PDF?) ---", file=sys.stderr)
            continue

        if args.out_dir:
            os.makedirs(args.out_dir, exist_ok=True)
            base = os.path.splitext(os.path.basename(input_path))[0]
            out_path = os.path.join(args.out_dir, f"{base}.txt")
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(full_text)
            print(f"{input_path} -> {out_path} ({len(full_text)} chars)")
        else:
            print(f"===== {input_path} ({len(full_text)} chars) =====")
            print(full_text)


def cmd_extract_images(args):
    """Pull every embedded image out of one or more documents so Claude (or
    the user) can look at them directly - this is the only reliable way to
    catch a logo or other identifying content baked into a picture, since
    none of the text-based detection can see inside an image. Nothing in
    the source document is modified; images are only copied out.
    """
    manifest = []
    for input_path in args.inputs:
        ext = os.path.splitext(input_path)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            print(f"Skipped {input_path}: unsupported file type {ext}", file=sys.stderr)
            continue

        with open(input_path, "rb") as f:
            file_bytes = f.read()

        sub_dir = os.path.join(args.out_dir, os.path.splitext(os.path.basename(input_path))[0])
        images = extract_images_for_format(file_bytes, ext, sub_dir)
        for img in images:
            img["input_file"] = os.path.abspath(input_path)
            manifest.append(img)
        print(f"{input_path}: extracted {len(images)} image(s) to {sub_dir}")

    if not manifest:
        print("No embedded images found in any input file.")
        return

    manifest_path = os.path.join(args.out_dir, "images_manifest.json")
    os.makedirs(args.out_dir, exist_ok=True)
    with open(manifest_path, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print(f"Manifest written to {manifest_path}")
    print("View each extracted image and decide which contain logos/entity-identifying")
    print("content. Add flagged ones to the review file's \"images\" list (with their")
    print("\"id\" and \"file_index\") before running apply - they will be blanked out.")


def cmd_scan(args):
    known_entities = _load_known_entities(args.known_entities)
    nlp = None if args.no_ner else try_load_ner()
    if not args.no_ner and nlp is None:
        print(
            "Note: spaCy (or its 'en_core_web_sm' model) isn't installed - continuing with "
            "regex-only detection. Address detection in particular is far more limited without "
            "it (Singapore-shaped addresses only). Install with:\n"
            "  pip install spacy && python -m spacy download en_core_web_sm",
            file=sys.stderr,
        )

    files_meta = []
    skipped = []
    per_file = []  # (file_index, input_path, full_text, sheet_locations, raw_candidates)

    for file_index, input_path in enumerate(args.inputs):
        ext = os.path.splitext(input_path)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            skipped.append((input_path, f"unsupported file type {ext}"))
            continue

        with open(input_path, "rb") as f:
            doc_bytes = f.read()
        doc = open_document(doc_bytes, ext)

        sheet_locations = None
        if ext == ".xlsx":
            full_text, sheet_locations = extract_full_text_xlsx(doc, with_locations=True)
        else:
            full_text = extract_text_for_format(doc, ext)

        files_meta.append({
            "file_index": file_index,
            "input_file": os.path.abspath(input_path),
            "file_name": os.path.basename(input_path),
            "extension": ext,
        })

        if not full_text.strip():
            skipped.append((input_path, "no extractable text found"))
            continue

        raw_candidates = find_candidates(
            full_text, known_entities, nlp,
            use_regex=not args.no_pattern_matching,
            address_only=not args.full_ner,
        )
        per_file.append((file_index, input_path, full_text, sheet_locations, raw_candidates))

    # Placeholders are assigned across the WHOLE batch at once (not per
    # file) so the same entity gets the same placeholder in every file it
    # appears in - see assign_placeholders in datamask_core.
    flat_candidates = [cand for _, _, _, _, raw in per_file for cand in raw]
    assigned = assign_placeholders(flat_candidates)

    all_items = []
    item_id = 0
    pos = 0
    for file_index, input_path, full_text, sheet_locations, raw_candidates in per_file:
        for _ in raw_candidates:
            cand = assigned[pos]
            pos += 1
            prefix, target, suffix = build_context_snippet(full_text, cand["text"])
            item = {
                "id": item_id,
                "file_index": file_index,
                "file_name": os.path.basename(input_path),
                "text": cand["text"],
                "category": cand["category"],
                "source": cand["source"],
                "placeholder": cand["placeholder"],
                "context": f"{prefix}[[{target}]]{suffix}",
                "include": True,
            }
            if sheet_locations is not None:
                idx = full_text.find(cand["text"])
                item["sheet"] = find_sheet_for_position(sheet_locations, idx)
            all_items.append(item)
            item_id += 1

    review = {"files": files_meta, "items": all_items, "images": []}
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(review, f, indent=2, ensure_ascii=False)

    print(f"Scanned {len(files_meta)} file(s), found {len(all_items)} candidate item(s) total.")
    for path, reason in skipped:
        print(f"  Skipped {path}: {reason}", file=sys.stderr)
    print(f"Review file written to {args.out}")
    print("Do not apply yet - show these to the user for approval/edits first.")
    print("Run `extract-images` too if the document has embedded pictures/logos -")
    print("that's a separate check the text scan above can't do.")


def cmd_apply(args):
    with open(args.review_file, encoding="utf-8") as f:
        review = json.load(f)

    replacement_logo_bytes = None
    if args.replacement_logo:
        with open(args.replacement_logo, "rb") as f:
            replacement_logo_bytes = f.read()

    files_meta = review["files"]
    items = review["items"]
    image_flags = review.get("images", [])

    os.makedirs(args.output_dir, exist_ok=True)
    combined_mapping = {}  # original -> placeholder, deduplicated across the whole batch
    total_masked = 0
    total_images_masked = 0
    any_pdf = False

    for meta in files_meta:
        file_index = meta["file_index"]
        file_items = [it for it in items if it["file_index"] == file_index and it.get("include", True)]
        file_images = [img for img in image_flags if img.get("file_index") == file_index and img.get("include", True)]
        # "action" is "blank" (the default - see is_replace below, for
        # review JSON written before this existed) or "replace"; a
        # "replace" flag with no --replacement-logo actually given falls
        # back to "blank" rather than erroring, since there's nothing
        # else sensible to insert.
        blank_ids = [
            img["id"] for img in file_images
            if img.get("action", "blank") != "replace" or not replacement_logo_bytes
        ]
        replace_ids = [
            img["id"] for img in file_images
            if img.get("action", "blank") == "replace" and replacement_logo_bytes
        ]
        if not file_items and not blank_ids and not replace_ids:
            print(f"Nothing to mask in {meta['file_name']} - skipping.")
            continue

        ext = meta["extension"]
        any_pdf = any_pdf or (ext == ".pdf")
        with open(meta["input_file"], "rb") as f:
            doc_bytes = f.read()
        doc = open_document(doc_bytes, ext)

        mapping = {it["text"]: it["placeholder"] for it in file_items}
        grey_texts = {it["text"] for it in file_items if should_grey_address(it["category"], it["placeholder"])}
        if mapping:
            apply_mapping_for_format(doc, mapping, ext, grey_texts)

        if ext == ".pdf":
            if blank_ids:
                total_images_masked += redact_images_in_pdf(doc, blank_ids)
            if replace_ids:
                total_images_masked += redact_images_in_pdf(doc, replace_ids, replacement_image_bytes=replacement_logo_bytes)
            output_buffer = save_document(doc, ext)
        else:
            output_buffer = save_document(doc, ext)
            if blank_ids:
                new_bytes, n_images = redact_images_in_saved_bytes(output_buffer.read(), ext, blank_ids)
                output_buffer = io.BytesIO(new_bytes)
                total_images_masked += n_images
            if replace_ids:
                new_bytes, n_images = redact_images_in_saved_bytes(
                    output_buffer.read(), ext, replace_ids, replacement_image_bytes=replacement_logo_bytes,
                )
                output_buffer = io.BytesIO(new_bytes)
                total_images_masked += n_images

        base = os.path.splitext(meta["file_name"])[0]
        base = sanitize_filename_with_mapping(base, mapping)
        out_path = os.path.join(args.output_dir, f"{base}{ext}")
        with open(out_path, "wb") as f:
            f.write(output_buffer.read())

        file_image_ids = blank_ids + replace_ids
        msg = f"Masked {len(mapping)} text item(s)"
        if file_image_ids:
            msg += f" and {len(file_image_ids)} image(s)"
        print(f"{msg} in {meta['file_name']} -> {out_path}")
        total_masked += len(mapping)
        combined_mapping.update(mapping)

    log_path = os.path.join(args.output_dir, "mapping_log.json")
    mapping_log = [{"original": k, "placeholder": v} for k, v in combined_mapping.items()]
    with open(log_path, "w", encoding="utf-8") as f:
        json.dump(mapping_log, f, indent=2, ensure_ascii=False)

    print(f"Done - {total_masked} text item(s) and {total_images_masked} image(s) masked "
          f"across {len(files_meta)} file(s).")
    print(f"Combined mapping log: {log_path}  (internal traceability record - do not share externally)")
    if any_pdf:
        print("Note: PDF masking works by redaction + redraw, not text editing - spot-check the output.")
    if total_images_masked:
        if replacement_logo_bytes:
            print("Note: redacted images are replaced with either a plain grey placeholder square or your")
            print("supplied replacement logo (per each image's \"action\"), which may stretch/crop to fit the")
            print("original image's frame - spot-check those spots too.")
        else:
            print("Note: redacted images are replaced with a plain grey placeholder square, which may")
            print("stretch/crop to fit the original image's frame - spot-check those spots too.")


def cmd_verify(args):
    """Residual scan / QA check for already-masked output. Checks whether
    any of the ORIGINAL values from a known-entities CSV still literally
    appear in the given (already-masked) file(s) - i.e. did anything on
    the list fail to get replaced.

    Deliberately does NOT re-run full auto-detection (regex/NER) against
    the masked output - only the known-entities list, and only their
    exact known-entities-list categories (never plain regex-guessed
    original text is fine since none is available for
    already-scrambled UEN/name/address/email placeholders, which are
    generated, not looked up). Re-running full detection against masked
    output produces a flood of false positives, since the tool's own
    placeholder text can look exactly like the real thing it's
    disguising (a fake UEN is still shaped like a UEN, a fake address's
    country name is still a country name) - checking only for the
    SPECIFIC original values you listed is what actually answers "did
    my masking work", without that noise.
    """
    known_entities = _load_known_entities(args.known_entities)

    any_leak = False
    for input_path in args.inputs:
        ext = os.path.splitext(input_path)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            print(f"Skipped {input_path}: unsupported file type {ext}", file=sys.stderr)
            continue

        with open(input_path, "rb") as f:
            doc_bytes = f.read()
        doc = open_document(doc_bytes, ext)
        if ext == ".xlsx":
            full_text, _ = extract_full_text_xlsx(doc, with_locations=True)
        else:
            full_text = extract_text_for_format(doc, ext)

        leaks = []
        for original, _category in known_entities:
            pattern = _build_flexible_entity_pattern(original)
            if pattern.search(full_text):
                leaks.append(original)

        if leaks:
            any_leak = True
            print(f"LEAK in {input_path}: {len(leaks)} original value(s) still present:")
            for original in leaks:
                print(f"  {original!r}")
        else:
            print(f"Clean: {input_path} - no known original values found.")

    if any_leak:
        print("\nAt least one file still contains an original value from the known-entities list.")
        print("This means either a row's exact wording doesn't match how it appears in this file")
        print("(see SKILL.md's 'exact wording' note), or the file wasn't actually run through apply.")
        sys.exit(1)
    else:
        print("\nAll clear - no known original values detected in any file checked.")


def main():
    parser = argparse.ArgumentParser(description="DataMask CLI - scan then apply, with a human review step in between")
    sub = parser.add_subparsers(dest="command", required=True)

    p_extract = sub.add_parser("extract-text", help="Dump plain text from one or more files, no detection - for reading/LLM entity-spotting before scan")
    p_extract.add_argument("inputs", nargs="+", help="Path(s) to .docx/.xlsx/.pptx/.pdf file(s)")
    p_extract.add_argument("--out-dir", default=None, help="Write one .txt file per input here instead of printing to stdout")
    p_extract.set_defaults(func=cmd_extract_text)

    p_extract_images = sub.add_parser("extract-images", help="Extract embedded images from one or more files for visual (logo/entity) review - for reading before scan/apply")
    p_extract_images.add_argument("inputs", nargs="+", help="Path(s) to .docx/.xlsx/.pptx/.pdf file(s)")
    p_extract_images.add_argument("--out-dir", default="extracted_images", help="Directory to write extracted image files + manifest.json")
    p_extract_images.set_defaults(func=cmd_extract_images)

    p_scan = sub.add_parser("scan", help="Detect candidate sensitive items across one or more files and write a review file")
    p_scan.add_argument("inputs", nargs="+", help="Path(s) to .docx/.xlsx/.pptx/.pdf file(s) to scan")
    p_scan.add_argument("--known-entities", default=None, help="Optional CSV of known names/entities (value,category)")
    p_scan.add_argument("--out", default="candidates.json", help="Where to write the review JSON")
    p_scan.add_argument("--no-pattern-matching", action="store_true", help="Skip email/phone regex detection")
    p_scan.add_argument("--no-ner", action="store_true", help="Skip spaCy NER (disables name/address auto-detection beyond the regex fallbacks; use if spaCy is unavailable or too slow on a huge batch)")
    p_scan.add_argument("--full-ner", action="store_true", help="Also use spaCy for general name/organisation detection, not just addresses (default is address-only, since general detection is a confirmed source of noise on real business documents - see SKILL.md)")
    p_scan.set_defaults(func=cmd_scan)

    p_apply = sub.add_parser("apply", help="Apply an approved/edited review file and produce the masked output(s)")
    p_apply.add_argument("review_file", help="The (possibly edited) JSON review file from `scan`")
    p_apply.add_argument("--output-dir", default=".", help="Directory to write masked file(s) + combined mapping log")
    p_apply.add_argument(
        "--replacement-logo", default=None,
        help="Path to an image file to use for any image entry marked \"action\": \"replace\" in the review "
             "JSON (as opposed to the default \"blank\", a plain grey square) - e.g. swapping a client's real "
             "logo for your own placeholder logo, everywhere it appears. Auto-converted to match each "
             "embedding slot's format, so any common image format works as input regardless of the original.",
    )
    p_apply.set_defaults(func=cmd_apply)

    p_verify = sub.add_parser("verify", help="QA check: confirm none of a known-entities CSV's original values remain in already-masked output file(s)")
    p_verify.add_argument("inputs", nargs="+", help="Path(s) to the MASKED/anonymised output file(s) to check")
    p_verify.add_argument("--known-entities", required=True, help="The same known-entities CSV used to produce this output")
    p_verify.set_defaults(func=cmd_verify)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
