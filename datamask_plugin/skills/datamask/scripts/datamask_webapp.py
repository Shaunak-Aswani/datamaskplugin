#!/usr/bin/env python3
"""
datamask_webapp.py - Standalone local web app for DataMask.

Run with:
    streamlit run datamask_webapp.py

This opens a page in your browser (usually http://localhost:8501) where
you can upload a document, review what will be masked (an editable table
per category, plus a deduped image gallery), and download the masked
result - all without a JSON file to hand-edit.

All the actual detection/masking logic lives in datamask_core.py and
datamask_webapp_logic.py (a thin in-memory wrapper over it, shared with
nothing Streamlit-specific so it can be unit tested on its own) - this
file is just the page/UI on top.

PERFORMANCE NOTE (read before changing the review section): an earlier
version rendered one st.checkbox + one st.text_input PER detected item -
for a real batch (300+ items across several files) that's 600+ individual
Streamlit widgets, and Streamlit reruns the ENTIRE script top-to-bottom on
every single widget interaction. That combination made the page
noticeably sluggish on every click, even though each click only touched
one row. Two changes fixed this:
  1. One st.data_editor (a spreadsheet-like table) per category instead
     of two widgets per row - a handful of components instead of hundreds,
     which Streamlit's frontend renders far more efficiently than an
     equivalent number of separate widgets.
  2. The whole review section lives inside an st.form, so edits (checking
     a box, changing replacement text) don't trigger a rerun AT ALL until
     "Apply masking" is clicked - only then does the script re-execute.
Together these are the difference between "a rerun of the whole page per
click" and "no reruns while reviewing, one rerun on submit."
"""
import io
import json
import os
import sys
import zipfile

import pandas as pd
import streamlit as st

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from datamask_core import try_load_ner
from datamask_webapp_logic import (
    parse_known_entities_csv,
    run_scan,
    run_extract_images,
    run_apply,
)

st.set_page_config(page_title="DataMask", page_icon="🔒", layout="wide")


@st.cache_resource(show_spinner=False)
def _cached_nlp():
    """Load spaCy once per server process, not once per rerun - Streamlit
    reruns this whole script on every click, and reloading a spaCy model
    each time would make every checkbox toggle take several seconds."""
    return try_load_ner()


CATEGORY_ORDER = ["Client", "Entity", "Individual", "Email", "Phone", "UEN", "ABN", "Address", "Location"]


def _cat_sort_key(cat):
    try:
        return (0, CATEGORY_ORDER.index(cat))
    except ValueError:
        return (1, cat)


# ---------------------------------------------------------------------------
# Page state
# ---------------------------------------------------------------------------
if "review" not in st.session_state:
    st.session_state.review = None
    st.session_state.doc_bytes_by_index = {}
    st.session_state.images = []
    st.session_state.outputs = None
    st.session_state.mapping_log = None

st.title("DataMask")
st.caption("Mask client-identifying information in Word, Excel, PowerPoint, and PDF files before sharing them.")

with st.sidebar:
    st.header("1. Upload")
    uploaded_docs = st.file_uploader(
        "Documents to mask", type=["docx", "xlsx", "pptx", "pdf"], accept_multiple_files=True,
    )
    known_csv = st.file_uploader(
        "Your own list of names/companies to catch (optional)", type=["csv"],
        help="A spreadsheet with two columns: value,category - e.g. 'Acme Corp,Client'. Use this for "
             "anything specific to this job that the automatic scan might miss.",
    )
    replacement_logo = st.file_uploader(
        "Replacement logo (optional)", type=["png", "jpg", "jpeg"],
        help="If provided, any image you flag as \"Replace with logo\" in the review step (rather than "
             "just blanking it) will be swapped for this image instead - e.g. replacing a client's real "
             "logo with your own placeholder logo, everywhere it appears.",
    )
    nlp_available = _cached_nlp() is not None
    use_ner = st.checkbox(
        "Detect addresses automatically", value=nlp_available, disabled=not nlp_available,
        help="Finds street addresses written in plain text (not already in your known-entities list). "
             "This needs a one-time extra setup step - ask your technical contact to run: "
             "pip install spacy && python -m spacy download en_core_web_sm" if not nlp_available else
             "Recommended - finds street addresses written in plain text.",
    )
    full_ner = st.checkbox(
        "Also detect people's and company names", value=False, disabled=not use_ner,
        help="Off by default - this can mistake ordinary words for names or companies (e.g. flagging "
             "\"Vessel\" or \"the Company\" as if they were real entities), which means more items for you "
             "to double-check. Turn it on if a document has names or companies that aren't already covered "
             "by your known-entities list, and you're willing to review some extra false matches to catch them.",
    )
    use_regex = not st.checkbox("Skip automatic detection of emails, phone numbers, and ID numbers (UEN/ABN)", value=False)
    scan_clicked = st.button("Scan for sensitive info", type="primary", disabled=not uploaded_docs)

# Computed once here (not inside the images-review block below) so it's
# always defined even when a scan finds no images to review at all -
# that block wouldn't otherwise run, and referencing an undefined
# variable at the "if submitted" apply step further down would crash.
replacement_logo_bytes = replacement_logo.getvalue() if replacement_logo else None

if scan_clicked and uploaded_docs:
    known_entities = parse_known_entities_csv(known_csv.getvalue() if known_csv else None)
    nlp = _cached_nlp() if use_ner else None
    uploaded_pairs = [(f.name, f.getvalue()) for f in uploaded_docs]
    with st.spinner("Scanning for sensitive information..."):
        review, doc_bytes_by_index, skipped = run_scan(
            uploaded_pairs, known_entities, nlp, use_regex, address_only=not full_ner,
        )
        images = run_extract_images(doc_bytes_by_index, review["files"])
    st.session_state.review = review
    st.session_state.doc_bytes_by_index = doc_bytes_by_index
    st.session_state.images = images
    st.session_state.outputs = None
    st.session_state.mapping_log = None
    for name, reason in skipped:
        st.warning(f"Skipped {name}: {reason}")

review = st.session_state.review

if review is None:
    st.info("Upload one or more documents and click **Scan for sensitive info** to get started.")
    st.stop()

if not review["items"] and not st.session_state.images:
    st.success("No sensitive information detected.")
    st.stop()

st.header("2. Review")
st.caption(
    "Uncheck anything you don't want masked, and edit any suggested replacement text. "
    "Names, registration numbers, and addresses get a realistic scrambled value by default "
    "(e.g. \"Mark Rober\" -> \"Jane Ali\") rather than a generic label - still editable below. "
    "Edits here don't take effect until you click **Apply masking** at the bottom."
)

by_category = {}
for it in review["items"]:
    by_category.setdefault(it["category"], []).append(it)

with st.form("review_form"):
    decisions = {}  # item id -> (include, placeholder)

    for cat in sorted(by_category.keys(), key=_cat_sort_key):
        cat_items = by_category[cat]
        with st.expander(f"{cat} ({len(cat_items)})", expanded=len(by_category) <= 4):
            df = pd.DataFrame([
                {
                    "id": it["id"],
                    "Mask": it.get("include", True),
                    "Context": it["context"],
                    "Replace with": it["placeholder"],
                }
                for it in cat_items
            ])
            edited = st.data_editor(
                df,
                key=f"editor_{cat}",
                hide_index=True,
                width="stretch",
                disabled=["id", "Context"],
                column_config={
                    "id": None,  # hidden - internal join key only
                    "Mask": st.column_config.CheckboxColumn("Mask", width="small"),
                    "Context": st.column_config.TextColumn("Match (in context)", width="large"),
                    "Replace with": st.column_config.TextColumn("Replace with", width="medium"),
                },
            )
            for _, row in edited.iterrows():
                decisions[row["id"]] = (bool(row["Mask"]), row["Replace with"])

    image_decisions = {}
    if st.session_state.images:
        st.subheader("Images")
        st.caption(
            "Check any that contain a logo, letterhead, or signature. \"Blank\" replaces it with a plain "
            "grey square; \"Replace with logo\" swaps in the logo you uploaded in the sidebar instead (only "
            "available if you uploaded one). The same image embedded multiple times (e.g. a logo repeated "
            "on every page) is shown once, with a count - your choice applies to every occurrence."
        )
        img_cols = st.columns(4)
        for i, img in enumerate(st.session_state.images):
            with img_cols[i % 4]:
                st.image(img["path"], width="stretch")
                count = img.get("count", 1)
                suffix = "" if count == 1 else f" ({count}x)"
                options = ["Leave as is", f"Blank{suffix}"]
                if replacement_logo_bytes:
                    options.append(f"Replace with logo{suffix}")
                action = st.radio(
                    "Action", options, key=f"img_{img['id']}_{img['file_index']}",
                    label_visibility="collapsed",
                )
                if action == "Leave as is":
                    resolved_action = None
                elif action.startswith("Blank"):
                    resolved_action = "blank"
                else:
                    resolved_action = "replace"
                image_decisions[(img["file_index"], tuple(img.get("all_ids", [img["id"]])))] = resolved_action

    st.text_area("Anything the scan missed? (optional, for your own reference)")

    n_total = len(review["items"])
    st.caption(f"{n_total} text item(s) and {len(st.session_state.images)} distinct image(s) detected.")
    submitted = st.form_submit_button("Apply masking", type="primary")

if submitted:
    for it in review["items"]:
        include, placeholder = decisions[it["id"]]
        it["include"] = include
        it["placeholder"] = placeholder
    # Expand each flagged image entry back out to every underlying
    # duplicate id (all_ids) - the review screen only shows one
    # thumbnail per unique image, but every occurrence needs its own
    # redaction entry so run_apply actually blanks/replaces all of them.
    # "action" travels with each entry ("blank" or "replace") so
    # run_apply knows which images from a "replace" flag to swap in
    # replacement_logo_bytes for, versus a "blank" flag getting the
    # plain grey square as before.
    flat_image_flags = []
    for (file_index, all_ids), action in image_decisions.items():
        if action is None:
            continue
        for img_id in all_ids:
            flat_image_flags.append({"file_index": file_index, "id": img_id, "include": True, "action": action})
    review["images"] = flat_image_flags

    with st.spinner("Applying..."):
        outputs, mapping_log = run_apply(
            review, st.session_state.doc_bytes_by_index, replacement_logo_bytes=replacement_logo_bytes,
        )
    st.session_state.outputs = outputs
    st.session_state.mapping_log = mapping_log

if st.session_state.outputs is not None:
    st.success(f"Done - {len(st.session_state.outputs)} file(s) masked.")

    # Download everything as one .zip - both what was actually asked for
    # (fewer clicks than downloading N files one at a time) and a
    # practical mitigation for a real, observed Safari issue: some
    # browser/Streamlit combinations misinterpret internal background
    # requests (image thumbnails, component data) as file downloads,
    # littering the Downloads folder with small hash-named junk files
    # alongside the real ones. That's a Safari-side quirk, not
    # corruption of the actual masked documents (the real output files
    # download fine, just with extra unrelated junk next to them) - but
    # a single zip download is fewer total download-triggering
    # interactions on the page, which reduces how often that has a
    # chance to happen, and is more reliable anyway.
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        seen_names = {}
        for name, data in st.session_state.outputs:
            # De-duplicate names WITHIN the zip itself - two output
            # files can legitimately end up with the same name (nothing
            # in either one matched the known-entities mapping, so
            # neither got renamed - see the "_anonymised suffix removed"
            # change). Two zip entries with the identical name is
            # technically valid but ambiguous to extract; append a
            # counter instead so both are always retrievable.
            if name in seen_names:
                seen_names[name] += 1
                base, ext = os.path.splitext(name)
                arcname = f"{base} ({seen_names[name]}){ext}"
            else:
                seen_names[name] = 0
                arcname = name
            zf.writestr(arcname, data)
        zf.writestr("mapping_log.json", json.dumps(st.session_state.mapping_log, indent=2, ensure_ascii=False))
    st.download_button(
        "⬇ Download all files (.zip)",
        data=zip_buffer.getvalue(),
        file_name="datamask_output.zip",
        mime="application/zip",
        key="download_all_zip",
        type="primary",
    )

    with st.expander("Or download files individually"):
        # key=... is required here, not optional: Streamlit auto-generates a
        # widget ID from the element type + its parameters, and two
        # download_button calls with the SAME label + file_name (e.g. two
        # output files that ended up with an identical name because nothing
        # in either one matched the known-entities mapping - see the
        # "_anonymised suffix removed" change, which made this collision
        # more likely) get treated as the SAME widget. Confirmed real bug,
        # not just a cosmetic duplicate-ID crash: Streamlit can end up
        # serving the wrong widget's data for a given click when IDs
        # collide, which is what "downloads a file, but it's not found /
        # won't open" looks like from the outside - the browser wrote
        # SOMETHING, just not reliably the file you actually clicked for.
        # The loop index is unique regardless of whether names collide, so
        # it's used here rather than the filename itself.
        for i, (name, data) in enumerate(st.session_state.outputs):
            st.download_button(f"Download {name}", data=data, file_name=name, key=f"download_output_{i}")
        st.download_button(
            "Download the record of what was replaced (keep this private - don't share it with the masked files)",
            data=json.dumps(st.session_state.mapping_log, indent=2, ensure_ascii=False),
            file_name="mapping_log.json",
            key="download_mapping_log",
    )
