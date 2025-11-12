import pandas as pd

# ---------- Konversi & Helpers ----------

def convert_sikap(value):
    """
    Mengubah string:
    'Sangat Baik (x) & Perlu Bimbingan (y)'
    → skor 1–5
    TANPA regex, hanya pakai operasi string biasa.
    """
    # Tangani nilai kosong / NaN
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0

    s = str(value).strip()

    # rapikan spasi ganda
    while "  " in s:
        s = s.replace("  ", " ")

    # cek frasa minimal
    if "Sangat Baik" not in s or "Perlu Bimbingan" not in s:
        return 0

    try:
        # Ambil angka setelah "Sangat Baik (" atau "Sangat Baik("
        if "Sangat Baik(" in s:
            sb_part = s.split("Sangat Baik(", 1)[1]
        else:
            sb_part = s.split("Sangat Baik (", 1)[1]
        sb_str = sb_part.split(")", 1)[0].strip()
        sb = int(sb_str)

        # Ambil angka setelah "Perlu Bimbingan (" atau "Perlu Bimbingan("
        if "Perlu Bimbingan(" in s:
            pb_part = s.split("Perlu Bimbingan(", 1)[1]
        else:
            pb_part = s.split("Perlu Bimbingan (", 1)[1]
        pb_str = pb_part.split(")", 1)[0].strip()
        pb = int(pb_str)
    except Exception:
        # kalau parsing gagal, kembalikan 0
        return 0

    # Konversi sb & pb ke skor 1–5
    if sb == 8 and pb == 0:
        return 5
    elif 5 <= sb <= 7 and 1 <= pb <= 3:
        return 4
    elif 3 <= sb <= 4 and 4 <= pb <= 5:
        return 3
    elif 1 <= sb <= 2 and 6 <= pb <= 7:
        return 2
    elif sb == 0 and pb == 8:
        return 1
    return 0


def convert_descriptive(value):
    if pd.isnull(value):
        return 0
    value_clean = str(value).strip().lower().replace('\xa0', ' ')
    mapping = {
        "sangat baik": 5,
        "baik": 4,
        "cukup": 3,
        "kurang": 2,
        "tidak ikut": 1,
    }
    return mapping.get(value_clean, 0)


def preprocess_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    # Terapkan konversi jika kolom tersedia
    if 'Nilai Sikap' in df.columns:
        df['Nilai Sikap'] = df['Nilai Sikap'].apply(convert_sikap)
    if 'Nilai Estrakulikuler' in df.columns:
        df['Nilai Estrakulikuler'] = df['Nilai Estrakulikuler'].apply(convert_descriptive)
    if 'Nilai Prakerin' in df.columns:
        df['Nilai Prakerin'] = df['Nilai Prakerin'].apply(convert_descriptive)
    if 'Nilai Prestasi' in df.columns:
        df['Nilai Prestasi'] = df['Nilai Prestasi'].apply(convert_descriptive)
    if 'Nilai Ketidakhadiran' in df.columns:
        df['Nilai Ketidakhadiran'] = pd.to_numeric(
            df['Nilai Ketidakhadiran'], errors='coerce'
        ).fillna(0)
        df['Nilai Ketidakhadiran'] = df['Nilai Ketidakhadiran'].clip(lower=1)

    # Bersihkan nama siswa
    if 'Nama Siswa' in df.columns:
        df['Nama Siswa'] = df['Nama Siswa'].fillna('Nama Tidak Diketahui')
        df = df[df['Nama Siswa'] != 'A0'].copy()

    # Hapus kolom duplikat
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # Tambah rid jika belum ada
    if 'rid' not in df.columns:
        df.insert(0, 'rid', range(1, len(df) + 1))

    return df

# ---------- ROC ----------

def roc_weights(n: int):
    return [sum(1 / (k + 1) for k in range(i, n)) / n for i in range(n)]

# ---------- ARAS ----------

def aras_compute(df_raw: pd.DataFrame, criteria: list, types: list):
    df = df_raw.copy()
    # Pastikan semua kriteria ada
    for c in criteria:
        if c not in df.columns:
            raise ValueError(f"Kolom kriteria '{c}' tidak ditemukan di data.")

    # A0 (solusi ideal)
    ideal = []
    for i, ctype in enumerate(types):
        col = criteria[i]
        if ctype == "benefit":
            ideal_value = df[col].max()
        else:
            # Ambil minimum non-0; kalau semua 0 → 0
            nonzero = df[col].replace(0, pd.NA)
            nonzero_min = nonzero.min(skipna=True)
            ideal_value = 0 if pd.isna(nonzero_min) else nonzero_min
        ideal.append(ideal_value)

    df_ideal = pd.DataFrame([[-1, 'A0'] + ideal], columns=['rid', 'Nama Siswa'] + criteria)
    if 'Nama Siswa' not in df.columns:
        df['Nama Siswa'] = df['rid'].astype(str)

    df_all = pd.concat(
        [df_ideal, df[['rid', 'Nama Siswa'] + criteria]],
        ignore_index=True
    )
    df_all['Alternatif'] = df_all['Nama Siswa']

    # Normalisasi ARAS
    df_norm = df_all.copy()
    for i, ctype in enumerate(types):
        col = criteria[i]
        if ctype == 'benefit':
            total = df_all[col].sum()
            df_norm[col] = df_all[col] / total if total != 0 else 0
        else:
            df_norm[col] = df_all[col].apply(lambda x: 0 if x == 0 else 1 / x)

    return df_all, df_norm


def aras_contributions(df_norm, weights, criteria, alt_name):
    """Kembalikan DataFrame kontribusi tiap kriteria untuk satu alternatif."""
    row = df_norm[df_norm['Alternatif'] == alt_name]
    if row.empty:
        return None
    contrib = []
    for c, w in zip(criteria, weights):
        val_norm = float(row.iloc[0][c])
        contrib_val = val_norm * w
        contrib.append({
            "Kriteria": c,
            "Bobot": w,
            "Nilai_Norm": val_norm,
            "Kontribusi": contrib_val,
        })
    out = pd.DataFrame(contrib).sort_values(
        "Kontribusi", ascending=False
    ).reset_index(drop=True)
    out["%_Terhadap_Skor"] = out["Kontribusi"] / out["Kontribusi"].sum() * 100.0
    return out


def aras_score_and_rank(
    df_all: pd.DataFrame,
    df_norm: pd.DataFrame,
    weights: pd.Series,
    criteria: list,
    types: list,
):
    # Hitung S_i
    df_norm = df_norm.copy()
    df_norm['S_i'] = 0.0
    for col in criteria:
        w = weights.loc[col] if col in weights.index else 0
        df_norm['S_i'] += df_norm[col] * w

    S0_series = df_norm.loc[df_norm['Alternatif'] == 'A0', 'S_i']
    if S0_series.empty:
        raise ValueError("Baris A0 tidak ditemukan.")
    S0 = S0_series.iloc[0]
    df_norm['U_i'] = df_norm['S_i'] / S0 if S0 != 0 else 0

    # Hasil gabungan
    df_kriteria_asli = df_all[['rid', 'Alternatif'] + criteria].copy()
    hasil = df_norm[['rid', 'Alternatif', 'S_i', 'U_i']].merge(
        df_kriteria_asli, on=['rid', 'Alternatif'], how='left'
    )

    sort_cols = ['U_i'] + criteria
    ascending_flags = [False] + [False if t == 'benefit' else True for t in types]

    hasil_all_sorted = hasil.sort_values(
        by=sort_cols,
        ascending=ascending_flags
    ).reset_index(drop=True)
    hasil_all_sorted['Ranking_dengan_A0'] = hasil_all_sorted.index + 1

    hasil_siswa = hasil[hasil['Alternatif'] != 'A0'].copy()
    hasil_siswa = hasil_siswa.sort_values(
        by=sort_cols,
        ascending=ascending_flags
    ).reset_index(drop=True)
    hasil_siswa['Ranking'] = hasil_siswa.index + 1

    return df_norm, hasil_siswa, hasil_all_sorted
