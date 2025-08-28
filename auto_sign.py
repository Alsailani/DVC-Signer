# auto_sign_app.py
# Run:  python -m streamlit run auto_sign_app.py

import io
import uuid
import zipfile
import random
import datetime
import fitz  # PyMuPDF
import streamlit as st

st.set_page_config(page_title="DVC Auto-Signer", page_icon="✅", layout="centered")

DEFAULT_NAME = "Mohammed Alsailani"

# ---------- helpers ----------
def text_width(txt, fontsize=11):
    try:
        return fitz.get_text_length(txt, fontname="helv", fontsize=fontsize)
    except Exception:
        return len(txt) * fontsize * 0.5

def find_ieq_sections(doc):
    """
    Return dict {page_index: ieq_heading_rect} for pages that contain the
    'Indoor Environmental Quality' heading. Use rect's Y to:
      - skip marking 'Yes' BELOW this Y (keep items above)
      - place 'NA' in the first 'Notes:' AFTER this Y (even on next page)
    """
    sections = {}
    for i, p in enumerate(doc):
        rects = p.search_for("Indoor Environmental Quality")
        if rects:
            sections[i] = rects[0]
    return sections

def mark_yes_boxes_red(
    doc,
    x_offset=9,     # smaller base offset so result sits more to the RIGHT
    box=10,
    width=1.2,
    skip_after_y_by_page=None,
):
    """
    Draw a red 'X' to the left of 'Yes' *checkbox labels* only (always humanized).

    Rules:
      - Require a 'No' on the SAME LINE to the RIGHT (30–250pt)
      - Skip 'Yes' preceded by numbers/% (e.g., '100% Yes')
      - Respect IEQ skip: skip 'Yes' BELOW the IEQ title Y on that page
      - Always add small random jitter to make marks look hand-drawn
    """
    skip_after_y_by_page = skip_after_y_by_page or {}

    for idx, page in enumerate(doc):
        words = page.get_text("words")
        if not words:
            continue

        # group words by approx line (via y-mid rounding)
        line_buckets = {}
        for w in words:
            x0, y0, x1, y1, txt, *_ = w
            y_mid = (y0 + y1) / 2
            key = round(y_mid * 2) / 2.0
            line_buckets.setdefault(key, []).append(w)

        y_threshold_rect = skip_after_y_by_page.get(idx, None)
        y_threshold = y_threshold_rect.y0 if y_threshold_rect else None

        for key, line_words in line_buckets.items():
            line_words.sort(key=lambda w: w[0])
            yes_cands = [w for w in line_words if w[4] == "Yes"]
            has_no = any(w[4] == "No" for w in line_words)

            for w in yes_cands:
                x0, y0, x1, y1, _txt, *_ = w

                # IEQ skip: anything below/at the IEQ title
                if y_threshold is not None and y0 >= y_threshold - 2:
                    continue

                # Must have a plausible "No" to the right on the same line
                if not has_no:
                    continue
                right_no = None
                min_dx = 1e9
                for ww in line_words:
                    if ww[4] == "No" and ww[0] > x1:
                        dx = ww[0] - x1
                        if 30 <= dx <= 250 and dx < min_dx:
                            min_dx = dx
                            right_no = ww
                if right_no is None:
                    # likely inline 'Yes' (not a checkbox label)
                    continue

                # Ignore "100% Yes": immediate left neighbor numeric or contains '%'
                left_neighbor = None
                for ww in line_words:
                    if ww[2] <= x0 and (x0 - ww[2]) < 25:
                        if (left_neighbor is None) or (ww[2] > left_neighbor[2]):
                            left_neighbor = ww
                if left_neighbor:
                    t = left_neighbor[4]
                    if any(ch.isdigit() for ch in t) or "%" in t:
                        continue

                # --- Humanized X geometry (always on) ---
                cx = x0 - (x_offset + random.uniform(-1.5, 1.5))  # shifted right baseline
                cy = y0 + (y1 - y0) / 2 + random.uniform(-1.2, 1.2)
                b = max(7.5, box + random.uniform(-1.0, 1.2))
                wline = max(0.9, width + random.uniform(-0.2, 0.25))

                xA, yA = cx - b/2, cy - b/2
                xB, yB = cx + b/2, cy + b/2
                xC, yC = cx - b/2, cy + b/2
                xD, yD = cx + b/2, cy - b/2

                page.draw_line((xA, yA), (xB, yB), color=(1, 0, 0), width=wline)
                page.draw_line((xC, yC), (xD, yD), color=(1, 0, 0), width=wline)

def insert_ieq_notes_na(doc, ieq_sections):
    """
    For each IEQ section, find the first 'Notes:' AFTER the section title.
    If none on same page, search next pages. Insert 'NA' near that label.
    """
    for start_page, ieq_rect in sorted(ieq_sections.items()):
        page = doc[start_page]
        # Notes on same page but below IEQ heading?
        candidates = [r for r in page.search_for("Notes:") if r.y0 > ieq_rect.y1]
        target = candidates[0] if candidates else None

        # else search forward
        if target is None:
            for pi in range(start_page + 1, len(doc)):
                nxt = doc[pi].search_for("Notes:")
                if nxt:
                    page = doc[pi]
                    target = nxt[0]
                    break

        if target is not None:
            page.insert_text(
                (target.x1 + 10, target.y1 + 8),
                "NA",
                fontsize=11,
                fontname="helv",
                fill=(1, 0, 0)
            )

def find_signature_page(doc):
    """Find page with signature section; robust across layouts."""
    for i, p in enumerate(doc):
        if p.search_for("Signature & Stamp of Verifying Licensed Professional"):
            return i
    # fallback: page with a 'Signature' nearest to a 'Date'
    best, best_gap = None, 1e9
    for i, p in enumerate(doc):
        sigs, dates = p.search_for("Signature"), p.search_for("Date")
        for s in sigs:
            for d in dates:
                gap = abs((s.y0+s.y1)/2 - (d.y0+d.y1)/2)
                if gap < best_gap:
                    best_gap, best = gap, i
    return best if best is not None else max(0, len(doc)-2)

def fill_signature_section_red(doc, signer_name: str, bottom_date: str, tol=60):
    """
    Fill (Name)->signer_name and (Date)->NA in the paragraph if present.
    Then place signature+date next to 'Signature'/'Date' labels paired by vertical proximity.
    Everything in red.
    """
    page = doc[find_signature_page(doc)]

    # Paragraph tokens
    name_marks = page.search_for("(Name)")
    date_marks_para = page.search_for("(Date)")
    if name_marks:
        r = name_marks[0]
        w = text_width(signer_name, 11)
        page.insert_text((r.x0 - 6 - w, r.y1 - 2), signer_name,
                         fontsize=11, fontname="helv", fill=(1, 0, 0))
    if date_marks_para:
        chosen = None
        if name_marks:
            ny = name_marks[0].y0
            for d in date_marks_para:
                if d.y0 >= ny - 5:
                    chosen = d; break
        if chosen is None:
            chosen = date_marks_para[0]
        w = text_width("NA", 11)
        page.insert_text((chosen.x0 - 6 - w, chosen.y1 - 2), "NA",
                         fontsize=11, fontname="helv", fill=(1, 0, 0))

    # Pair proper Signature/Date by vertical proximity
    sig_rects = page.search_for("Signature")
    date_rects = page.search_for("Date")
    best_pair, best_gap = None, 1e9
    for s in sig_rects:
        for d in date_rects:
            gap = abs((s.y0+s.y1)/2 - (d.y0+d.y1)/2)
            if gap <= tol and gap < best_gap:
                best_gap, best_pair = gap, (s, d)
    if not best_pair and sig_rects and date_rects:
        s = min(sig_rects, key=lambda r: r.y0)  # upper signature label
        d = min(date_rects, key=lambda r: abs((r.y0+r.y1)/2 - (s.y0+s.y1)/2))
        best_pair = (s, d)

    if best_pair:
        s, d = best_pair
        page.insert_text((s.x1 + 12, s.y1 - 2), signer_name,
                         fontsize=16, fontname="helv", fill=(1, 0, 0))
        page.insert_text((d.x1 + 12, d.y1 - 2), bottom_date,
                         fontsize=12, fontname="helv", fill=(1, 0, 0))

# ---------- UI ----------
st.title("ENERGY STAR DVC – Auto Fill & Sign")

row1 = st.columns(3)
with row1[0]:
    opt_yes = st.checkbox("Mark all 'Yes' with red X", value=True)
with row1[1]:
    opt_sign = st.checkbox("Fill paragraph + sign/date", value=True)
with row1[2]:
    opt_ieq = st.checkbox("Leave IEQ blank + 'NA' in Notes", value=True)

name = st.text_input("Signer name", value=DEFAULT_NAME, disabled=not opt_sign)
date_str = st.text_input("Bottom date (YYYY-MM-DD)",
                         value=datetime.date.today().strftime("%Y-%m-%d"),
                         disabled=not opt_sign)

uploads = st.file_uploader("Upload one or more PDFs", type=["pdf"], accept_multiple_files=True)

# Persist results across reruns
if "results" not in st.session_state:
    st.session_state["results"] = []

if st.button("Process PDFs"):
    st.session_state["results"] = []  # clear previous run
    if not uploads:
        st.warning("Please upload at least one PDF.")
    elif not opt_yes and not opt_sign and not opt_ieq:
        st.warning("Select at least one action.")
    else:
        with st.spinner("Processing..."):
            for uploaded in uploads:
                try:
                    file_bytes = uploaded.read()
                    if not file_bytes:
                        st.error(f"{uploaded.name}: empty upload (try reselecting the file).")
                        continue

                    doc = fitz.open(stream=file_bytes, filetype="pdf")

                    if getattr(doc, "needs_pass", False):
                        st.error(f"{uploaded.name}: PDF is password-protected; cannot process.")
                        continue

                    st.caption(f"Loaded {uploaded.name} — {len(doc)} page(s), {len(file_bytes)/1024:.1f} KB")

                    # IEQ sections
                    ieq_sections = find_ieq_sections(doc) if (opt_ieq or opt_yes) else {}

                    # 1) Mark 'Yes' everywhere EXCEPT below the IEQ heading on that page
                    if opt_yes:
                        mark_yes_boxes_red(
                            doc,
                            skip_after_y_by_page=ieq_sections,
                        )

                    # 2) IEQ Notes: write 'NA' below IEQ (may be next page)
                    if opt_ieq and ieq_sections:
                        insert_ieq_notes_na(doc, ieq_sections)

                    # 3) Fill/sign
                    if opt_sign:
                        fill_signature_section_red(doc, name.strip(), date_str.strip())

                    # Save to memory
                    out = io.BytesIO()
                    doc.save(out)
                    out.seek(0)

                    stem = uploaded.name.rsplit(".", 1)[0]
                    out_name = f"{stem}_signed.pdf"
                    st.session_state["results"].append({
                        "filename": out_name,
                        "bytes": out.getvalue()
                    })
                except Exception as e:
                    st.exception(e)

# ---- Results (persistent) ----
if st.session_state["results"]:
    st.write("### Results")
    for item in st.session_state["results"]:
        st.download_button(
            label=f"Download {item['filename']}",
            data=item["bytes"],
            file_name=item["filename"],
            mime="application/pdf",
            key=f"{item['filename']}-{uuid.uuid4()}"
        )

    # Download ALL as one ZIP
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for item in st.session_state["results"]:
            zf.writestr(item["filename"], item["bytes"])
    zip_buffer.seek(0)

    st.download_button(
        label="⬇️ Download ALL as ZIP",
        data=zip_buffer.getvalue(),
        file_name=f"DVC_outputs_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip",
        key=f"zip-{uuid.uuid4()}"
    )
