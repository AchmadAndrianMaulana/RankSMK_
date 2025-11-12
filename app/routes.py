# Standard library
import os
from datetime import datetime
import numbers  # ← dipakai di _df_to_xml_rows

# Third-party
import pandas as pd
from flask import (
    Blueprint, render_template, request, redirect,
    url_for, flash, current_app, send_file,
)
from flask_login import login_required, current_user
from werkzeug.utils import secure_filename

# ---- Helpers untuk HTML float formatting (fungsi, bukan string) ----
def _fmt6(x: float) -> str:
    """Format float 6 decimal; kembalikan string selalu."""
    try:
        if pd.isna(x):
            return ""
        return f"{float(x):.6f}"
    except Exception:
        # fallback: tetap kembalikan string agar tidak None
        return str(x)

# ---- Minimal Excel 2003 XML writer (no external engine) ----
def _df_to_xml_rows(df: pd.DataFrame) -> str:
    """Konversi DataFrame ke baris XML (format Excel 2003)."""
    def esc(s: object) -> str:
        return (
            str(s)
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    rows: list[str] = []

    # Header
    header_cells = "".join(
        f"<Cell><Data ss:Type='String'>{esc(col)}</Data></Cell>"
        for col in df.columns
    )
    rows.append(f"<Row>{header_cells}</Row>")

    # Data
    for _, row in df.iterrows():
        cells: list[str] = []
        for val in row:
            if pd.isna(val):
                # kosong → string kosong
                cells.append("<Cell><Data ss:Type='String'></Data></Cell>")
            # Angka integer (bukan bool)
            elif isinstance(val, numbers.Integral) and not isinstance(val, bool):
                cells.append(
                    f"<Cell><Data ss:Type='Number'>{int(val)}</Data></Cell>"
                )
            # Angka real (float dsb, bukan bool)
            elif isinstance(val, numbers.Real) and not isinstance(val, bool):
                cells.append(
                    f"<Cell><Data ss:Type='Number'>{float(val)}</Data></Cell>"
                )
            else:
                # Selain itu anggap string
                cells.append(
                    f"<Cell><Data ss:Type='String'>{esc(val)}</Data></Cell>"
                )
        rows.append(f"<Row>{''.join(cells)}</Row>")

    return "\n".join(rows)

def make_excel_xml_sheets(sheets: dict) -> str:
    # sheets: {name: pandas.DataFrame}
    ws_xml = []
    for name, df in sheets.items():
        ws_xml.append(
            f"<Worksheet ss:Name='{name}'>"
            f"<Table>{_df_to_xml_rows(df)}</Table>"
            f"</Worksheet>"
        )
    xml = (
        "<?xml version='1.0'?>"
        "<?mso-application progid='Excel.Sheet'?>"
        "<Workbook xmlns='urn:schemas-microsoft-com:office:spreadsheet' "
        " xmlns:o='urn:schemas-microsoft-com:office:office' "
        " xmlns:x='urn:schemas-microsoft-com:office:excel' "
        " xmlns:ss='urn:schemas-microsoft-com:office:spreadsheet' "
        " xmlns:html='http://www.w3.org/TR/REC-html40'>"
        + "".join(ws_xml) +
        "</Workbook>"
    )
    return xml

# ---- Minimal PDF writer (no external library) ----
def _escape_pdf_text(s: str) -> str:
    return str(s).replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")

def write_pdf_table(path: str, title: str, df: pd.DataFrame, max_rows_per_page: int = 45):
    # --- Siapkan isi tabel sebagai teks monospace (ASCII table) ---
    # Header & data
    headers = [str(c) for c in df.columns.tolist()]
    data_rows = []
    for _, r in df.iterrows():
        row_vals = []
        for v in r.tolist():
            row_vals.append("" if pd.isna(v) else str(v))
        data_rows.append(row_vals)

    # Hitung lebar kolom (max len header / isi)
    col_widths = [len(h) for h in headers]
    for row in data_rows:
        for i, val in enumerate(row):
            if len(val) > col_widths[i]:
                col_widths[i] = len(val)

    # (Opsional) batasi lebar kolom agar tidak terlalu lebar di PDF
    max_col_width = 25  # boleh kamu atur 20/30 sesuai kebutuhan
    col_widths = [min(w, max_col_width) for w in col_widths]

    def _clip(val: str, width: int) -> str:
        s = val
        if len(s) > width:
            # potong dan tambahkan … jika terlalu panjang
            return s[: max(0, width - 1)] + "…" if width > 1 else s[:width]
        return s

    def _make_row(cells):
        padded = [
            _clip(str(c), col_widths[i]).ljust(col_widths[i])
            for i, c in enumerate(cells)
        ]
        return "| " + " | ".join(padded) + " |"

    # Garis pemisah
    sep_line = "+-" + "-+-".join("-" * w for w in col_widths) + "-+"

    lines = []
    lines.append(sep_line)
    lines.append(_make_row(headers))
    lines.append(sep_line)
    for row in data_rows:
        lines.append(_make_row(row))
    lines.append(sep_line)

    # --- Paging: bagi per halaman berdasarkan jumlah baris ---
    pages_lines, tmp = [], []
    for line in lines:
        tmp.append(line)
        if len(tmp) >= max_rows_per_page:
            pages_lines.append(tmp)
            tmp = []
    if tmp:
        pages_lines.append(tmp)

    # --- Bangun stream konten PDF tiap halaman ---
    contents = []
    for lines_page in pages_lines:
        parts = []
        parts.append("BT")
        # Title di atas
        parts.append("/F1 14 Tf")
        parts.append("40 800 Td")
        parts.append(f"({_escape_pdf_text(title)}) Tj")
        parts.append("T*")
        parts.append("14 TL")
        parts.append("T*")
        # Tabel
        parts.append("/F1 9 Tf")   # font tabel
        parts.append("11 TL")      # jarak antar baris
        for ln in lines_page:
            parts.append(f"({_escape_pdf_text(ln)}) Tj")
            parts.append("T*")
        parts.append("ET")

        data = "\n".join(parts).encode("latin-1", "ignore")
        contents.append(
            b"<< /Length " + str(len(data)).encode() + b" >>\nstream\n" +
            data +
            b"\nendstream\n"
        )

    # --- Struktur PDF (objek, pages, xref) ---
    catalog = b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
    # PENTING: pakai font monospace Courier
    font    = b"3 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj\n"
    res     = b"<< /Font << /F1 3 0 R >> >>"

    objs = [catalog, b"", font]  # pages placeholder di index 1

    kids = []
    objnum = 4
    for stream in contents:
        # objek content stream
        objs.append(f"{objnum} 0 obj\n".encode() + stream + b"endobj\n")
        cnum = objnum
        objnum += 1

        # objek page
        page = (
            f"{objnum} 0 obj\n".encode() +
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources " +
            res +
            b" /Contents " + f"{cnum} 0 R".encode() + b" >>\nendobj\n"
        )
        objs.append(page)
        kids.append(f"{objnum} 0 R")
        objnum += 1

    kids_arr = "[ " + " ".join(kids) + " ]"
    pages = (
        b"2 0 obj\n<< /Type /Pages /Count " + str(len(kids)).encode() +
        b" /Kids " + kids_arr.encode() + b" >>\nendobj\n"
    )
    objs[1] = pages

    pdf = b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n"
    xref = []
    off = len(pdf)
    for o in objs:
        xref.append(off)
        pdf += o
        off = len(pdf)
    xstart = len(pdf)
    pdf += f"xref\n0 {len(objs)+1}\n".encode()
    pdf += b"0000000000 65535 f \n"
    for pos in xref:
        pdf += f"{pos:010d} 00000 n \n".encode()
    pdf += b"trailer\n<< /Size " + str(len(objs)+1).encode() + b" /Root 1 0 R >>\n"
    pdf += b"startxref\n" + str(xstart).encode() + b"\n%%EOF"

    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as f:
        f.write(pdf)


# Local packages
from . import db
from .models import Upload, Criteria, Subcriteria
from .forms import UploadForm, CriteriaForm, SubcriteriaForm
from .utils import (
    preprocess_dataframe, roc_weights, aras_compute, aras_score_and_rank, aras_contributions
)

# (Dihilangkan) Optional dependency reportlab — tidak diperlukan.

main_bp = Blueprint("main", __name__)

ALLOWED_EXT = {".xlsx", ".xls"}

def _latest_upload():
    return Upload.query.order_by(Upload.created_at.desc()).first()

@main_bp.route("/")
def index():
    return redirect(url_for("main.dashboard"))

@main_bp.route("/dashboard")
@login_required
def dashboard():
    up = _latest_upload()
    return render_template("dashboard.html", upload=up)

# ---------------- Upload data ----------------
@main_bp.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    form = UploadForm()
    if form.validate_on_submit():
        f = form.file.data
        ext = os.path.splitext(f.filename)[1].lower()
        if ext not in ALLOWED_EXT:
            flash("Format harus Excel (.xlsx / .xls).", "warning")
            return redirect(url_for("main.upload"))

        safe = secure_filename(f.filename)
        save_path = os.path.join(current_app.config["UPLOAD_FOLDER"], safe)
        os.makedirs(current_app.config["UPLOAD_FOLDER"], exist_ok=True)
        f.save(save_path)

        try:
            df = pd.read_excel(save_path)
        except Exception as e:
            flash(f"Gagal membaca Excel: {e}", "danger")
            return redirect(url_for("main.upload"))

        df = preprocess_dataframe(df)
        processed_path = save_path + ".processed.csv"
        df.to_csv(processed_path, index=False)

        # ---- FIX Pylance (hindari kwargs pada model) ----
        up = Upload()
        up.user_id = current_user.id
        up.original_filename = safe
        up.saved_path = save_path
        up.processed_path = processed_path
        up.n_rows = len(df)

        db.session.add(up)
        db.session.commit()
        flash("Upload & proses berhasil.", "success")
        return redirect(url_for("main.upload"))

    uploads = Upload.query.order_by(Upload.created_at.desc()).all()
    return render_template("upload.html", form=form, uploads=uploads)

@main_bp.route("/upload/<int:uid>/view")
@login_required
def upload_view(uid):
    up = Upload.query.get_or_404(uid)

    raw_html = "<em>File tidak ditemukan.</em>"
    proc_html = "<em>Belum ada data olah.</em>"

    if up.saved_path and os.path.exists(up.saved_path):
        try:
            dfr = pd.read_excel(up.saved_path).head(100)
            raw_html = dfr.to_html(classes="table table-sm table-striped", index=False)
        except Exception as e:
            raw_html = f"<em>Gagal baca raw: {e}</em>"

    if up.processed_path and os.path.exists(up.processed_path):
        try:
            dfp = pd.read_csv(up.processed_path).head(100)
            proc_html = dfp.to_html(classes="table table-sm table-striped", index=False)
        except Exception as e:
            proc_html = f"<em>Gagal baca processed: {e}</em>"

    return render_template(
        "upload_view.html",
        up=up,
        raw_html=raw_html,
        proc_html=proc_html,
    )

@main_bp.route("/upload/<int:uid>/reprocess", methods=["POST"])
@login_required
def upload_reprocess(uid):
    up = Upload.query.get_or_404(uid)
    if not up.saved_path or not os.path.exists(up.saved_path):
        flash("File mentah tidak ada. Silakan upload ulang.", "warning")
        return redirect(url_for("main.upload"))

    try:
        df = pd.read_excel(up.saved_path)
        df = preprocess_dataframe(df)
        processed_path = up.saved_path + ".processed.csv"
        df.to_csv(processed_path, index=False)
        up.processed_path = processed_path
        up.n_rows = len(df)
        db.session.commit()
        flash("Reproses berhasil.", "success")
    except Exception as e:
        flash(f"Gagal memproses ulang: {e}", "danger")
    return redirect(url_for("main.upload_view", uid=uid))

@main_bp.route("/upload/<int:uid>/delete", methods=["POST"])
@login_required
def upload_delete(uid):
    up = Upload.query.get_or_404(uid)
    for p in [up.saved_path, up.processed_path]:
        if p and os.path.exists(p):
            try:
                os.remove(p)
            except Exception:
                pass
    db.session.delete(up)
    db.session.commit()
    flash("Upload dihapus.", "info")
    return redirect(url_for("main.upload"))

@main_bp.route("/uploads/<path:filename>")
@login_required
def download_raw(filename):
    path = os.path.join(current_app.config["UPLOAD_FOLDER"], filename)
    if not os.path.exists(path):
        flash("File tidak ditemukan.", "warning")
        return redirect(url_for("main.upload"))
    return send_file(path, as_attachment=True)

# ----------------- Criteria CRUD + ROC -----------------
@main_bp.route("/criteria", methods=["GET", "POST"])
@login_required
def criteria_list():
    form = CriteriaForm()
    if form.validate_on_submit():
        # ---- FIX Pylance ----
        c = Criteria()
        c.name = (form.name.data or "").strip()
        c.ctype = form.ctype.data
        c.display_order = form.display_order.data

        db.session.add(c)
        db.session.commit()
        flash("Kriteria ditambahkan.", "success")
        return redirect(url_for("main.criteria_list"))
    items = Criteria.query.order_by(Criteria.display_order.asc()).all()
    n = len(items)
    weights = {}
    if n > 0:
        roc = roc_weights(n)
        for i, c in enumerate(items):
            weights[c.name] = roc[i]
    return render_template("criteria.html", form=form, items=items, weights=weights)

@main_bp.route("/criteria/<int:cid>/edit", methods=["GET", "POST"])
@login_required
def criteria_edit(cid):
    c = Criteria.query.get_or_404(cid)
    form = CriteriaForm(obj=c)
    if form.validate_on_submit():
        c.name = (form.name.data or "").strip()
        c.ctype = form.ctype.data
        c.display_order = form.display_order.data
        db.session.commit()
        flash("Kriteria diupdate.", "success")
        return redirect(url_for("main.criteria_list"))
    return render_template("criteria_edit.html", form=form, c=c)

@main_bp.route("/criteria/<int:cid>/delete", methods=["POST"])
@login_required
def criteria_delete(cid):
    c = Criteria.query.get_or_404(cid)
    db.session.delete(c)
    db.session.commit()
    flash("Kriteria dihapus.", "info")
    return redirect(url_for("main.criteria_list"))

# ----------------- Subcriteria CRUD -----------------
@main_bp.route("/subcriteria", methods=["GET", "POST"])
@login_required
def subcriteria_list():
    form = SubcriteriaForm()
    form.criteria_id.choices = [
        (c.id, c.name) for c in Criteria.query.order_by(Criteria.display_order.asc()).all()
    ]

    if form.validate_on_submit():
        # ---- FIX Pylance ----
        s = Subcriteria()
        s.criteria_id = form.criteria_id.data
        s.name = (form.name.data or "").strip()
        s.min_val = form.min_val.data

        db.session.add(s)
        db.session.commit()
        flash("Subkriteria ditambahkan.", "success")
        return redirect(url_for("main.subcriteria_list"))

    items = Subcriteria.query.all()
    criteria_lookup = {c.id: c.name for c in Criteria.query.order_by(Criteria.display_order.asc()).all()}
    return render_template("subcriteria.html", form=form, items=items, criteria_lookup=criteria_lookup)

@main_bp.route("/subcriteria/<int:sid>/edit", methods=["GET", "POST"])
@login_required
def subcriteria_edit(sid):
    s = Subcriteria.query.get_or_404(sid)

    form = SubcriteriaForm(obj=s)
    form.criteria_id.choices = [
        (c.id, c.name) for c in Criteria.query.order_by(Criteria.display_order.asc()).all()
    ]

    if request.method == "GET":
        form.criteria_id.data = s.criteria_id

    if form.validate_on_submit():
        s.criteria_id = form.criteria_id.data
        s.name = (form.name.data or "").strip()
        s.min_val = form.min_val.data
        db.session.commit()
        flash("Subkriteria diupdate.", "success")
        return redirect(url_for("main.subcriteria_list"))

    return render_template("subcriteria_edit.html", form=form, s=s)

@main_bp.route("/subcriteria/<int:sid>/delete", methods=["POST"])
@login_required
def subcriteria_delete(sid):
    s = Subcriteria.query.get_or_404(sid)
    db.session.delete(s)
    db.session.commit()
    flash("Subkriteria dihapus.", "info")
    return redirect(url_for("main.subcriteria_list"))

@main_bp.route("/aras/matrix")
@login_required
def aras_matrix():
    up = _latest_upload()
    if not up or not up.processed_path or not os.path.exists(up.processed_path):
        flash("Belum ada data terproses. Silakan upload data.", "warning")
        return redirect(url_for("main.upload"))

    df = pd.read_csv(up.processed_path)

    criteria_items = Criteria.query.order_by(Criteria.display_order.asc()).all()
    if not criteria_items:
        flash("Belum ada kriteria. Tambahkan dulu di menu Kriteria.", "warning")
        return redirect(url_for("main.criteria_list"))

    criteria = [c.name for c in criteria_items]
    types = [c.ctype for c in criteria_items]

    roc = roc_weights(len(criteria))
    weights = pd.Series(roc, index=criteria)

    # Matriks A & normalisasi (termasuk A0)
    df_all, df_norm = aras_compute(df, criteria, types)

    # ==== Pastikan 'Alternatif' pakai NAMA, bukan angka ====
    name_keywords = ("alternatif", "nama", "siswa", "peserta", "calon")
    text_cols = [c for c in df.columns if df[c].dtype == object]
    candidates_kw = [c for c in text_cols if any(k in c.lower() for k in name_keywords)]
    candidates_uniq = [c for c in text_cols if df[c].nunique(dropna=True) >= max(5, int(0.5 * df.shape[0]))]

    name_col_detected = None
    for c in candidates_kw + candidates_uniq:
        if c != "rid":
            name_col_detected = c
            break

    name_map: pd.Series | None = None
    if name_col_detected and "rid" in df.columns:
        name_map = (
            df[["rid", name_col_detected]]
            .drop_duplicates(subset=["rid"])
            .set_index("rid")[name_col_detected]
            .astype(str)
        )

    def _apply_name_map(frame: pd.DataFrame) -> pd.DataFrame:
        if "Alternatif" not in frame.columns:
            return frame
        out = frame.copy()
        if name_map is not None and "rid" in out.columns:
            out["Alternatif"] = out["rid"].map(name_map).fillna(out["Alternatif"]).astype(str)
            return out
        if name_map is not None:
            alt_numeric = pd.to_numeric(out["Alternatif"], errors="coerce")
            mapped = alt_numeric.map(name_map)
            out["Alternatif"] = mapped.fillna(out["Alternatif"]).astype(str)
            return out
        if name_col_detected and out["Alternatif"].dtype != object:
            out["Alternatif"] = out["Alternatif"].astype(str)
        return out

    df_all = _apply_name_map(df_all)
    df_norm = _apply_name_map(df_norm)

    # Tambahan guard kedua (jika ada 'rid' dan nama yang jelas)
    name_candidates = ["Alternatif", "Nama", "Nama Siswa", "nama", "nama_siswa"]
    name_col_explicit = next((c for c in name_candidates if c in df.columns), None)
    if "rid" in df.columns and name_col_explicit:
        name_map2 = (
            df[["rid", name_col_explicit]]
            .drop_duplicates(subset=["rid"])
            .set_index("rid")[name_col_explicit]
            .astype(str)
        )
        if "rid" in df_all.columns and "Alternatif" in df_all.columns:
            df_all["Alternatif"] = df_all["rid"].map(name_map2).fillna(df_all["Alternatif"]).astype(str)
        if "rid" in df_norm.columns and "Alternatif" in df_norm.columns:
            df_norm["Alternatif"] = df_norm["rid"].map(name_map2).fillna(df_norm["Alternatif"]).astype(str)

    # Matriks ternormalisasi * bobot
    base_cols = ["Alternatif"] + [c for c in criteria if c in df_norm.columns]
    df_weighted = df_norm[base_cols].copy()
    for col in criteria:
        if col in df_weighted.columns:
            df_weighted[col] = df_weighted[col] * weights[col]

    # Hitung S_i & U_i
    df_norm_scored, _, _ = aras_score_and_rank(df_all, df_norm, weights, criteria, types)

    if "Alternatif" in df_norm_scored.columns and name_map is not None:
        if "rid" in df_norm_scored.columns:
            df_norm_scored["Alternatif"] = (
                df_norm_scored["rid"].map(name_map).fillna(df_norm_scored["Alternatif"]).astype(str)
            )
        else:
            alt_numeric = pd.to_numeric(df_norm_scored["Alternatif"], errors="coerce")
            df_norm_scored["Alternatif"] = (
                alt_numeric.map(name_map).fillna(df_norm_scored["Alternatif"]).astype(str)
            )

    # Tabel bobot
    weights_df = pd.DataFrame({
        "Kriteria": criteria,
        "Tipe": types,
        "Bobot ROC": [weights.get(c, float("nan")) for c in criteria],
    })

    return render_template(
        "aras_matrix.html",
        columns=df_all.columns.tolist(),
        sample_all=df_all.head(20).to_html(classes="table table-sm table-striped", index=False),
        sample_norm=df_norm.head(20).to_html(classes="table table-sm table-striped", index=False),
        weights_html=weights_df.to_html(classes="table table-sm table-striped", index=False, float_format=_fmt6),
        weighted_html=df_weighted.head(20).to_html(classes="table table-sm table-striped", index=False, float_format=_fmt6),
        si_ui_html=df_norm_scored[['Alternatif', 'S_i', 'U_i']].head(20).to_html(classes="table table-sm table-striped", index=False, float_format=_fmt6),
    )

@main_bp.route("/aras/ranking", methods=["GET", "POST"])
@login_required
def aras_ranking():
    up = _latest_upload()
    if not up or not up.processed_path or not os.path.exists(up.processed_path):
        flash("Belum ada data terproses. Silakan upload data.", "warning")
        return redirect(url_for("main.upload"))

    df = pd.read_csv(up.processed_path)

    # ===== Helper kolom fleksibel
    def _norm(s: str) -> str:
        return str(s).strip().lower().replace(" ", "").replace("_", "")

    def _find_col(frame, wanted):
        norm_map = {_norm(c): c for c in frame.columns}
        for w in wanted:
            if w in norm_map:
                return norm_map[w]
        return None

    col_jur = _find_col(df, ["jurusan"])
    col_per = _find_col(df, ["periode"])
    col_jp = _find_col(df, [
        "jurusanperiode", "jurusanperiodegabung", "jurusanperiodecombined",
        "jurusan&periode", "jurusandanperiode", "jurusan-periode"
    ])
    use_combined = col_jp is not None and (col_jur is None and col_per is None)

    # ==== Inisialisasi default untuk menghindari "possibly unbound" ====
    jurusanperiode_values: list[str] = []
    selected_jurusanperiode: str = ""
    jurusan_values: list[str] = []
    periode_values: list[str] = []
    selected_jurusan: str = ""
    selected_periode: str = ""

    # ==== Dropdown values & selected ====
    if use_combined:
        jurusanperiode_values = sorted(df[col_jp].dropna().astype(str).unique().tolist())
        selected_jurusanperiode = request.form.get("jurusanperiode", "") if request.method == "POST" else ""
    else:
        if col_jur:
            jurusan_values = sorted(df[col_jur].dropna().astype(str).unique().tolist())
        if col_per:
            periode_values = sorted(df[col_per].dropna().astype(str).unique().tolist())
        selected_jurusan = request.form.get("jurusan", "") if request.method == "POST" else ""
        selected_periode = request.form.get("periode", "") if request.method == "POST" else ""

    # ==== Filtering ====
    dff = df.copy()
    if use_combined:
        if selected_jurusanperiode:
            dff = dff[dff[col_jp].astype(str) == selected_jurusanperiode]
    else:
        if col_jur and selected_jurusan:
            dff = dff[dff[col_jur].astype(str) == selected_jurusan]
        if col_per and selected_periode:
            dff = dff[dff[col_per].astype(str) == selected_periode]

    # ==== Hitung ROC–ARAS ====
    criteria_items = Criteria.query.order_by(Criteria.display_order.asc()).all()
    if not criteria_items:
        flash("Belum ada kriteria. Tambahkan dulu di menu Kriteria.", "warning")
        return redirect(url_for("main.criteria_list"))

    criteria = [c.name for c in criteria_items]
    types = [c.ctype for c in criteria_items]
    roc = roc_weights(len(criteria))
    weights = pd.Series(roc, index=criteria)

    df_all, df_norm = aras_compute(dff, criteria, types)
    df_norm, hasil_siswa, _ = aras_score_and_rank(df_all, df_norm, weights, criteria, types)

    # ==== Pastikan 'Alternatif' pakai NAMA ====
    if "rid" in hasil_siswa.columns and "rid" in dff.columns and "Alternatif" in dff.columns:
        name_map = (
            dff[["rid", "Alternatif"]]
            .drop_duplicates(subset=["rid"])
            .set_index("rid")["Alternatif"]
            .astype(str)
        )
        if "Alternatif" in hasil_siswa.columns:
            hasil_siswa["Alternatif"] = (
                hasil_siswa["rid"].map(name_map).fillna(hasil_siswa["Alternatif"])
            ).astype(str)

    # # ====== Susun DataFrame rapi untuk output ======
    # result_df = hasil_siswa.copy()

    # extra_cols = []
    # for c in [col_jur, col_per, col_jp]:
    #     if c and c in dff.columns:
    #         extra_cols.append(c)
    # if extra_cols:
    #     extras = dff[["rid"] + extra_cols].drop_duplicates(subset=["rid"])
    #     result_df = result_df.merge(extras, on="rid", how="left")

    # preferred = ["Ranking", "Alternatif", "S_i", "U_i"]
    # if use_combined and col_jp:
    #     preferred.insert(2, col_jp)
    # else:
    #     if col_jur:
    #         preferred.insert(2, col_jur)
    #     if col_per:
    #         preferred.insert(3, col_per)

    # criteria_cols = [c for c in criteria if c in result_df.columns]
    # other_cols = [c for c in result_df.columns if c not in (preferred + criteria_cols + ["rid"])]
    # result_df = result_df[[c for c in preferred if c in result_df.columns] + criteria_cols + other_cols]

    # for c in ["S_i", "U_i"]:
    #     if c in result_df.columns:
    #         result_df[c] = result_df[c].astype(float).round(6)

        # ====== Susun DataFrame rapi untuk output ======
    result_df = hasil_siswa.copy()

    # Tetap boleh dibulatkan dulu
    for c in ["S_i", "U_i"]:
        if c in result_df.columns:
            result_df[c] = result_df[c].astype(float).round(6)

    # >>> HANYA ambil kolom yang diinginkan
    cols_keep = [c for c in ["Ranking", "Alternatif", "S_i", "U_i"] if c in result_df.columns]
    slim_df = result_df[cols_keep].copy()


    # ====== Nama file dinamis (ikut filter) ======
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    parts = ["ranking"]
    if use_combined:
        if selected_jurusanperiode:
            parts.append(selected_jurusanperiode)
    else:
        if col_jur and selected_jurusan:
            parts.append(selected_jurusan)
        if col_per and selected_periode:
            parts.append(selected_periode)

    top = hasil_siswa.nlargest(10, "U_i")[["Alternatif", "U_i"]]
    chart_labels = top["Alternatif"].tolist()
    chart_values = top["U_i"].round(4).tolist()

    base = secure_filename("_".join(parts)) or "ranking"

    # ====== Tulis PDF (tanpa library)
    os.makedirs(current_app.config["DOWNLOAD_FOLDER"], exist_ok=True)
    pdf_filename = f"{base}_{ts}.pdf"
    pdf_path = os.path.join(current_app.config["DOWNLOAD_FOLDER"], pdf_filename)
    write_pdf_table(pdf_path, "Hasil Perangkingan ROC–ARAS", slim_df)
    download_pdf_url = url_for("main.download_file", filename=pdf_filename)

    # ====== Tulis Excel XML (XLS)
    sheets = {"Ranking": slim_df.copy()}
    info_rows = [["Dibuat pada", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]]
    if use_combined:
        info_rows.append(["Mode Filter", "JurusanPeriode"])
        # >>> Perbaikan di sini: gunakan operator Python `or`
        info_rows.append(["Filter JurusanPeriode", selected_jurusanperiode or "(Semua)"])
    else:
        info_rows.append(["Mode Filter", "Jurusan + Periode"])
        info_rows.append(["Filter Jurusan", selected_jurusan or "(Semua)"])
        info_rows.append(["Filter Periode", selected_periode or "(Semua)"])
    sheets["Info"] = pd.DataFrame(info_rows, columns=["Kunci", "Nilai"])
    if len(criteria) > 0:
        sheets["Bobot ROC"] = pd.DataFrame({
            "Kriteria": criteria,
            "Tipe": types,
            "Bobot ROC": [weights[c] for c in criteria],
        })
    excel_filename = f"{base}_{ts}.xls"
    excel_path = os.path.join(current_app.config["DOWNLOAD_FOLDER"], excel_filename)
    with open(excel_path, "w", encoding="utf-8") as f:
        f.write(make_excel_xml_sheets(sheets))

    # # ==== Tabel HTML dengan tombol Lihat Detail ====
    # hasil_view = hasil_siswa.copy()
    # hasil_view["Lihat Detail"] = hasil_view["Alternatif"].apply(
    #     lambda x: f'<a class="btn btn-sm btn-outline-primary" '
    #               f'href="{url_for("main.hasil_detail", alt=x)}" '
    #               f'title="Lihat kontribusi tiap kriteria">Lihat Detail</a>'
    # )
    # cols = [c for c in hasil_view.columns if c != "Lihat Detail"] + ["Lihat Detail"]
    # table_html = hasil_view[cols].to_html(
    #     classes="table table-sm table-hover",
    #     index=False,
    #     escape=False
    # )

        # ==== Tabel HTML tanpa kolom Lihat Detail ====
        table_html = slim_df.to_html(
        classes="table table-sm table-hover",
        index=False,
        float_format=_fmt6,
        escape=True
    )


    if use_combined:
        return render_template(
            "ranking.html",
            use_combined=True,
            jurusanperiode_values=jurusanperiode_values,
            selected_jurusanperiode=selected_jurusanperiode,
            table_html=table_html,
            download_url=download_pdf_url,
            download_excel_url=url_for("main.download_file", filename=excel_filename),
            chart_labels=chart_labels, chart_values=chart_values,
        )
    else:
        return render_template(
            "ranking.html",
            use_combined=False,
            jurusan_values=jurusan_values if col_jur else [],
            periode_values=periode_values if col_per else [],
            selected_jurusan=selected_jurusan if col_jur else "",
            selected_periode=selected_periode if col_per else "",
            table_html=table_html,
            download_url=download_pdf_url,
            download_excel_url=url_for("main.download_file", filename=excel_filename),
            chart_labels=chart_labels, chart_values=chart_values,
        )

# @main_bp.route("/hasil/<alt>")
# @login_required
# def hasil_detail(alt):
#     last_upload = Upload.query.order_by(Upload.created_at.desc()).first()
#     if not last_upload or not last_upload.processed_path or not os.path.exists(last_upload.processed_path):
#         flash("Belum ada data terproses.", "warning")
#         return redirect(url_for("main.upload"))

#     df = pd.read_csv(last_upload.processed_path)

#     criteria_objs = Criteria.query.order_by(Criteria.display_order.asc()).all()
#     if not criteria_objs:
#         flash("Belum ada kriteria.", "warning")
#         return redirect(url_for("main.criteria_list"))

#     criteria = [c.name for c in criteria_objs]
#     types = [c.ctype for c in criteria_objs]
#     weights_ = roc_weights(len(criteria))

#     df_all, df_norm = aras_compute(df, criteria, types)
#     df_norm_scored, hasil_siswa, _ = aras_score_and_rank(
#         df_all, df_norm, pd.Series(weights_, index=criteria), criteria, types
#     )

#     contrib = aras_contributions(df_norm, weights_, criteria, alt)
#     if contrib is None or contrib.empty:
#         flash("Alternatif tidak ditemukan.", "warning")
#         return redirect(url_for("main.aras_ranking"))

#     table_html = contrib.to_html(classes="table table-sm table-hover", index=False, escape=False)
#     return render_template("hasil_detail.html", alt=alt, table_html=table_html)

@main_bp.route("/help")
@login_required
def help_page():
    return render_template("help.html")

@main_bp.route("/hasil/export/pdf")
@login_required
def hasil_export_pdf():
    up = _latest_upload()
    if not up or not up.processed_path or not os.path.exists(up.processed_path):
        flash("Belum ada data terproses. Silakan upload data.", "warning")
        return redirect(url_for("main.upload"))

    df = pd.read_csv(up.processed_path)

    criteria_items = Criteria.query.order_by(Criteria.display_order.asc()).all()
    if not criteria_items:
        flash("Belum ada kriteria.", "warning")
        return redirect(url_for("main.criteria_list"))

    criteria = [c.name for c in criteria_items]
    types = [c.ctype for c in criteria_items]
    weights = pd.Series(roc_weights(len(criteria)), index=criteria)

    df_all, df_norm = aras_compute(df, criteria, types)
    df_norm_scored, hasil_siswa, _ = aras_score_and_rank(df_all, df_norm, weights, criteria, types)

    out = hasil_siswa[['Ranking', 'Alternatif', 'S_i', 'U_i']].copy()
    out['S_i'] = out['S_i'].astype(float).round(6)
    out['U_i'] = out['U_i'].astype(float).round(6)
    out = out.head(30)

    os.makedirs(current_app.config["DOWNLOAD_FOLDER"], exist_ok=True)
    path = os.path.join(current_app.config["DOWNLOAD_FOLDER"], "ranking.pdf")
    write_pdf_table(path, "Hasil Perangkingan (Top 30) — ROC–ARAS", out)

    return send_file(path, as_attachment=True)

@main_bp.route("/downloads/<path:filename>")
@login_required
def download_file(filename):
    path = os.path.join(current_app.config["DOWNLOAD_FOLDER"], filename)
    if not os.path.exists(path):
        flash("File tidak ditemukan.", "warning")
        return redirect(url_for("main.dashboard"))
    return send_file(path, as_attachment=True)
