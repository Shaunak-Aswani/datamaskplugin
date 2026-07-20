"""
datamask_webapp_logic.py - the non-UI logic behind datamask_webapp.py,
split into its own module so it can be imported and tested without
executing a Streamlit page (importing datamask_webapp.py directly would
re-run the whole page, since Streamlit scripts execute top to bottom).

Everything here is a thin wrapper over datamask_core.py that works with
in-memory uploaded bytes instead of file paths, mirroring what
datamask_cli.py's cmd_scan/cmd_apply do for the CLI. Keeping the actual
detection/masking logic itself only in datamask_core.py (not duplicated
here) is what keeps CLI, skill, and web-app behavior in sync.
"""
import csv
import io
import os
import tempfile

from datamask_core import (
    open_document,
    extract_text_for_format,
    extract_full_text_xlsx,
    find_sheet_for_position,
    find_candidates,
    build_context_snippet,
    assign_placeholders,
    apply_mapping_for_format,
    save_document,
    sanitize_filename_with_mapping,
    load_known_entities_rows,
    extract_images_for_format,
    redact_images_in_saved_bytes,
    redact_images_in_pdf,
    should_grey_address,
)

SUPPORTED_EXTENSIONS = {".docx", ".xlsx", ".pptx", ".pdf"}


def parse_known_entities_csv(raw_bytes):
    """raw_bytes: the uploaded CSV's raw content, or None."""
    if raw_bytes is None:
        return []
    text = raw_bytes.decode("utf-8-sig")
    reader = csv.reader(text.splitlines())
    rows = [r for r in reader if r and r[0].strip()]
    # Skip a header row (e.g. "value,category") rather than treating it as
    # a literal entity to match - see datamask_cli.py's _load_known_entities
    # for why this matters (masking the literal word "value" everywhere).
    if rows and rows[0][0].strip().lower() == "value":
        rows = rows[1:]
    pairs = [(r[0], r[1] if len(r) > 1 else "Entity") for r in rows]
    return load_known_entities_rows(pairs)


def run_scan(uploaded_files, known_entities, nlp, use_regex, address_only=True):
    """uploaded_files: list of (file_name, raw_bytes) tuples.
    Returns (review_dict, doc_bytes_by_index, skipped_list)."""
    files_meta = []
    per_file = []
    skipped = []
    doc_bytes_by_index = {}

    for file_index, (file_name, doc_bytes) in enumerate(uploaded_files):
        ext = os.path.splitext(file_name)[1].lower()
        if ext not in SUPPORTED_EXTENSIONS:
            skipped.append((file_name, f"unsupported file type {ext}"))
            continue

        doc_bytes_by_index[file_index] = doc_bytes
        doc = open_document(doc_bytes, ext)

        sheet_locations = None
        if ext == ".xlsx":
            full_text, sheet_locations = extract_full_text_xlsx(doc, with_locations=True)
        else:
            full_text = extract_text_for_format(doc, ext)

        files_meta.append({"file_index": file_index, "file_name": file_name, "extension": ext})

        if not full_text.strip():
            skipped.append((file_name, "no extractable text found"))
            continue

        raw_candidates = find_candidates(full_text, known_entities, nlp, use_regex=use_regex, address_only=address_only)
        per_file.append((file_index, file_name, full_text, sheet_locations, raw_candidates))

    flat_candidates = [cand for _, _, _, _, raw in per_file for cand in raw]
    assigned = assign_placeholders(flat_candidates)

    all_items = []
    item_id = 0
    pos = 0
    for file_index, file_name, full_text, sheet_locations, raw_candidates in per_file:
        for _ in raw_candidates:
            cand = assigned[pos]
            pos += 1
            prefix, target, suffix = build_context_snippet(full_text, cand["text"])
            item = {
                "id": item_id,
                "file_index": file_index,
                "file_name": file_name,
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
    return review, doc_bytes_by_index, skipped


def run_extract_images(doc_bytes_by_index, files_meta):
    """Extract every embedded image from each uploaded file to a temp
    directory and return a flat list of image dicts with file_index/
    file_name attached."""
    all_images = []
    tmp_root = tempfile.mkdtemp(prefix="datamask_images_")
    for meta in files_meta:
        file_index = meta["file_index"]
        doc_bytes = doc_bytes_by_index[file_index]
        out_dir = os.path.join(tmp_root, str(file_index))
        images = extract_images_for_format(doc_bytes, meta["extension"], out_dir)
        for img in images:
            img["file_index"] = file_index
            img["file_name"] = meta["file_name"]
            all_images.append(img)
    return all_images


def run_apply(review, doc_bytes_by_index, replacement_logo_bytes=None):
    """Returns (list of (out_filename, bytes) tuples, mapping_log list).

    replacement_logo_bytes: optional bytes of a user-supplied replacement
    image (e.g. a client's real logo swapped for a masked-entity's own
    placeholder logo). Each flagged image in review["images"] carries an
    "action" of "blank" (the default, for backward compatibility with
    review JSON that predates this - see is_replace below) or "replace" -
    only images flagged "replace" use replacement_logo_bytes; "blank"
    ones still get the plain grey square as before, and if
    replacement_logo_bytes wasn't actually provided, a "replace" flag
    silently falls back to "blank" rather than erroring, since there's
    nothing else sensible to insert.
    """
    files_meta = review["files"]
    items = review["items"]
    image_flags = review.get("images", [])

    outputs = []
    combined_mapping = {}

    for meta in files_meta:
        file_index = meta["file_index"]
        file_items = [it for it in items if it["file_index"] == file_index and it.get("include", True)]
        file_images = [img for img in image_flags if img.get("file_index") == file_index and img.get("include", True)]
        blank_ids = [
            img["id"] for img in file_images
            if img.get("action", "blank") != "replace" or not replacement_logo_bytes
        ]
        replace_ids = [
            img["id"] for img in file_images
            if img.get("action", "blank") == "replace" and replacement_logo_bytes
        ]
        if not file_items and not blank_ids and not replace_ids:
            continue

        ext = meta["extension"]
        doc_bytes = doc_bytes_by_index[file_index]
        doc = open_document(doc_bytes, ext)

        mapping = {it["text"]: it["placeholder"] for it in file_items}
        grey_texts = {it["text"] for it in file_items if should_grey_address(it["category"], it["placeholder"])}
        if mapping:
            apply_mapping_for_format(doc, mapping, ext, grey_texts)

        if ext == ".pdf":
            if blank_ids:
                redact_images_in_pdf(doc, blank_ids)
            if replace_ids:
                redact_images_in_pdf(doc, replace_ids, replacement_image_bytes=replacement_logo_bytes)
            output_buffer = save_document(doc, ext)
        else:
            output_buffer = save_document(doc, ext)
            if blank_ids:
                new_bytes, _ = redact_images_in_saved_bytes(output_buffer.read(), ext, blank_ids)
                output_buffer = io.BytesIO(new_bytes)
            if replace_ids:
                new_bytes, _ = redact_images_in_saved_bytes(
                    output_buffer.read(), ext, replace_ids, replacement_image_bytes=replacement_logo_bytes,
                )
                output_buffer = io.BytesIO(new_bytes)

        base = os.path.splitext(meta["file_name"])[0]
        base = sanitize_filename_with_mapping(base, mapping)
        out_name = f"{base}{ext}"
        outputs.append((out_name, output_buffer.read()))
        combined_mapping.update(mapping)

    mapping_log = [{"original": k, "placeholder": v} for k, v in combined_mapping.items()]
    return outputs, mapping_log
