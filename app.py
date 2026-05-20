"""
Backend API - Sistem Rekomendasi Properti
Menggunakan Flask + AHP + Profile Matching
+ Penyesuaian Wilayah: Provinsi → Kota → Kecamatan
+ Validasi Input
+ Endpoint Evaluasi Massal (1000 konsumen sintetik untuk Bab 4)
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import numpy as np
import os
import random

app = Flask(__name__)
CORS(app)

# ===================== MAPPING PROVINSI =====================
# Ketentuan khusus: batasi cakupan hanya untuk Pulau Jawa — tiga provinsi berikut saja
PROVINSI_KOTA_MAP = {
    "Jawa Barat": [
        "Bandung", "Kota Bandung", "Kabupaten Bandung", "Kabupaten Bandung Barat",
        "Bekasi", "Kota Bekasi", "Kabupaten Bekasi",
        "Bogor", "Kota Bogor", "Kabupaten Bogor",
        "Cimahi", "Kota Cimahi",
        "Cianjur", "Kabupaten Cianjur",
        "Cirebon", "Kota Cirebon", "Kabupaten Cirebon",
        "Depok", "Kota Depok",
        "Garut", "Kabupaten Garut",
        "Indramayu", "Kabupaten Indramayu",
        "Karawang", "Kabupaten Karawang",
        "Kuningan", "Kabupaten Kuningan",
        "Majalengka", "Kabupaten Majalengka",
        "Pangandaran", "Kabupaten Pangandaran",
        "Purwakarta", "Kabupaten Purwakarta",
        "Subang", "Kabupaten Subang",
        "Sukabumi", "Kota Sukabumi", "Kabupaten Sukabumi",
        "Sumedang", "Kabupaten Sumedang",
        "Tasikmalaya", "Kota Tasikmalaya", "Kabupaten Tasikmalaya",
    ],
    "Jawa Tengah": [
        "Banjarnegara", "Kabupaten Banjarnegara",
        "Banyumas", "Kabupaten Banyumas",
        "Batang", "Kabupaten Batang",
        "Blora", "Kabupaten Blora",
        "Boyolali", "Kabupaten Boyolali",
        "Brebes", "Kabupaten Brebes",
        "Cilacap", "Kabupaten Cilacap",
        "Demak", "Kabupaten Demak",
        "Grobogan", "Kabupaten Grobogan",
        "Jepara", "Kabupaten Jepara",
        "Karanganyar", "Kabupaten Karanganyar",
        "Kebumen", "Kabupaten Kebumen",
        "Kendal", "Kabupaten Kendal",
        "Klaten", "Kabupaten Klaten",
        "Kudus", "Kabupaten Kudus",
        "Magelang", "Kota Magelang", "Kabupaten Magelang",
        "Pati", "Kabupaten Pati",
        "Pekalongan", "Kota Pekalongan", "Kabupaten Pekalongan",
        "Pemalang", "Kabupaten Pemalang",
        "Purbalingga", "Kabupaten Purbalingga",
        "Purworejo", "Kabupaten Purworejo",
        "Rembang", "Kabupaten Rembang",
        "Salatiga", "Kota Salatiga",
        "Semarang", "Kota Semarang", "Kabupaten Semarang",
        "Sragen", "Kabupaten Sragen",
        "Sukoharjo", "Kabupaten Sukoharjo",
        "Surakarta", "Kota Surakarta", "Solo", "Kota Solo",
        "Tegal", "Kota Tegal", "Kabupaten Tegal",
        "Temanggung", "Kabupaten Temanggung",
        "Wonogiri", "Kabupaten Wonogiri",
        "Wonosobo", "Kabupaten Wonosobo",
    ],
    "Jawa Timur": [
        "Bangkalan", "Kabupaten Bangkalan",
        "Banyuwangi", "Kabupaten Banyuwangi",
        "Blitar", "Kota Blitar", "Kabupaten Blitar",
        "Bojonegoro", "Kabupaten Bojonegoro",
        "Bondowoso", "Kabupaten Bondowoso",
        "Gresik", "Kabupaten Gresik",
        "Jember", "Kabupaten Jember",
        "Jombang", "Kabupaten Jombang",
        "Kediri", "Kota Kediri", "Kabupaten Kediri",
        "Lamongan", "Kabupaten Lamongan",
        "Lumajang", "Kabupaten Lumajang",
        "Madiun", "Kota Madiun", "Kabupaten Madiun",
        "Magetan", "Kabupaten Magetan",
        "Malang", "Kota Malang", "Kabupaten Malang",
        "Mojokerto", "Kota Mojokerto", "Kabupaten Mojokerto",
        "Nganjuk", "Kabupaten Nganjuk",
        "Ngawi", "Kabupaten Ngawi",
        "Pacitan", "Kabupaten Pacitan",
        "Pamekasan", "Kabupaten Pamekasan",
        "Pasuruan", "Kota Pasuruan", "Kabupaten Pasuruan",
        "Ponorogo", "Kabupaten Ponorogo",
        "Probolinggo", "Kota Probolinggo", "Kabupaten Probolinggo",
        "Sampang", "Kabupaten Sampang",
        "Sidoarjo", "Kabupaten Sidoarjo",
        "Situbondo", "Kabupaten Situbondo",
        "Sumenep", "Kabupaten Sumenep",
        "Surabaya", "Kota Surabaya",
        "Trenggalek", "Kabupaten Trenggalek",
        "Tuban", "Kabupaten Tuban",
        "Tulungagung", "Kabupaten Tulungagung",
    ],
}

# Buat lookup terbalik: kota → provinsi
KOTA_TO_PROVINSI = {}
for provinsi, kota_list in PROVINSI_KOTA_MAP.items():
    for kota in kota_list:
        KOTA_TO_PROVINSI[kota.lower()] = provinsi

def get_provinsi(kota_kab):
    if pd.isna(kota_kab):
        return "Lainnya"
    return KOTA_TO_PROVINSI.get(str(kota_kab).strip().lower(), "Lainnya")

# ===================== LOAD DATASET =====================
DATA_PATH = "data_original.xlsx"

def load_data():
    if not os.path.exists(DATA_PATH):
        return generate_sample_data()
    df = pd.read_excel(DATA_PATH)
    df = preprocess(df)
    return df

def generate_sample_data():
    np.random.seed(42)
    n = 100
    data_wilayah = [
        ("Menteng", "Jakarta Pusat"),
        ("Kebayoran Baru", "Jakarta Selatan"),
        ("Kelapa Gading", "Jakarta Utara"),
        ("Cibubur", "Jakarta Timur"),
        ("Serpong", "Tangerang Selatan"),
        ("Ciputat", "Tangerang Selatan"),
        ("Margonda", "Depok"),
        ("Cimanggis", "Depok"),
        ("Bekasi Timur", "Bekasi"),
        ("Bekasi Barat", "Bekasi"),
        ("Bogor Tengah", "Bogor"),
        ("Cibinong", "Bogor"),
        ("Cicendo", "Bandung"),
        ("Coblong", "Bandung"),
        ("Lowokwaru", "Malang"),
        ("Klojen", "Malang"),
        ("Gubeng", "Surabaya"),
        ("Rungkut", "Surabaya"),
        ("Tembalang", "Semarang"),
        ("Banyumanik", "Semarang"),
        ("Banjarsari", "Surakarta"),
        ("Laweyan", "Surakarta"),
        ("Depok", "Sleman"),
        ("Gamping", "Sleman"),
    ]
    rows = []
    for _ in range(n):
        kec, kota = random.choice(data_wilayah)
        rows.append({
            'Price_Clean': np.random.randint(300_000_000, 5_000_000_000),
            'Luas_bangunan': np.random.randint(36, 500),
            'Luas_tanah': np.random.randint(60, 600),
            'Kamar_tidur_clean': np.random.choice([1, 2, 3, 4, 5]),
            'Kamar_mandi': np.random.choice([1, 2, 3, 4]),
            'Jenis_Properti': np.random.choice(['Rumah', 'Apartemen', 'Townhouse']),
            'Kecamatan': kec,
            'Kota_Kab': kota,
        })
    return pd.DataFrame(rows)

def preprocess(df):
    cols_needed = ['Price_Clean', 'Luas_bangunan', 'Kamar_tidur_clean',
                   'Jenis_Properti', 'Kecamatan', 'Kota_Kab']
    for col in cols_needed:
        if col not in df.columns:
            df[col] = 0
    df = df.dropna(subset=['Price_Clean', 'Luas_bangunan', 'Kamar_tidur_clean'])
    if 'Provinsi' not in df.columns:
        df['Provinsi'] = df['Kota_Kab'].apply(get_provinsi)
    else:
        mask_kosong = df['Provinsi'].isna() | (df['Provinsi'] == '')
        df.loc[mask_kosong, 'Provinsi'] = df.loc[mask_kosong, 'Kota_Kab'].apply(get_provinsi)
    return df

# Load data saat startup
df_properti = load_data()
# Terapkan batasan geografis: hanya properti yang berada di Pulau Jawa (Jawa Barat, Jawa Tengah, Jawa Timur)
allowed_provinsi = set(PROVINSI_KOTA_MAP.keys())
if 'Provinsi' in df_properti.columns:
    # Normalisasi nilai provinsi untuk menghindari perbedaan case/spasi
    df_properti['Provinsi'] = df_properti['Provinsi'].astype(str).str.strip()
    df_properti = df_properti[df_properti['Provinsi'].isin(allowed_provinsi)].reset_index(drop=True)
else:
    # Jika kolom Provinsi tidak ada setelah preprocessing, pastikan tetap kosong
    df_properti = df_properti.iloc[0:0]

# ===================== AHP FUNCTIONS =====================
def ahp_weights(criteria_matrix):
    matrix = np.array(criteria_matrix, dtype=float)
    n = matrix.shape[0]
    geo_means = np.prod(matrix, axis=1) ** (1.0 / n)
    weights = geo_means / geo_means.sum()
    weighted_sum = matrix @ weights
    lambda_max = np.mean(weighted_sum / weights)
    ci = (lambda_max - n) / (n - 1) if n > 1 else 0
    ri_table = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24, 7: 1.32}
    ri = ri_table.get(n, 1.49)
    cr = ci / ri if ri > 0 else 0
    return {
        'weights': weights.tolist(),
        'consistency_ratio': round(cr, 4),
        'is_consistent': cr < 0.10
    }

def get_weights_by_pekerjaan(pekerjaan):
    """Mengembalikan bobot AHP berdasarkan profesi konsumen."""
    pekerjaan_lower = pekerjaan.strip().lower()
    if pekerjaan_lower == 'buruh':
        return [0.45, 0.10, 0.10, 0.05, 0.30]
    elif pekerjaan_lower == 'asn':
        return [0.25, 0.20, 0.15, 0.10, 0.30]
    elif pekerjaan_lower == 'pengusaha':
        return [0.10, 0.35, 0.25, 0.15, 0.15]
    else:  # Umum / default
        return [0.35, 0.20, 0.15, 0.10, 0.20]

# ===================== VALIDASI INPUT =====================
def validate_preferences(data):
    """
    Validasi input preferensi dari user.
    Mengembalikan (is_valid: bool, error_message: str)
    """
    errors = []

    max_price = data.get('max_price', 0)
    min_luas = data.get('min_luas', 0)
    kamar_tidur = data.get('kamar_tidur', 0)
    top_k = data.get('top_k', 10)

    # Validasi harga
    try:
        max_price = float(max_price)
        if max_price < 0:
            errors.append("Budget maksimal tidak boleh negatif.")
        elif max_price > 0 and max_price < 100_000_000:
            errors.append("Budget maksimal terlalu kecil (min Rp 100.000.000).")
    except (ValueError, TypeError):
        errors.append("Budget maksimal harus berupa angka.")

    # Validasi luas
    try:
        min_luas = float(min_luas)
        if min_luas < 0:
            errors.append("Luas bangunan tidak boleh negatif.")
        elif min_luas > 10000:
            errors.append("Luas bangunan tidak realistis (maks 10.000 m²).")
    except (ValueError, TypeError):
        errors.append("Luas bangunan harus berupa angka.")

    # Validasi kamar tidur
    try:
        kamar_tidur = int(kamar_tidur)
        if kamar_tidur < 0 or kamar_tidur > 10:
            errors.append("Jumlah kamar tidur harus antara 0 - 10.")
    except (ValueError, TypeError):
        errors.append("Jumlah kamar tidur harus berupa angka bulat.")

    # Validasi top_k
    try:
        top_k = int(top_k)
        if top_k < 1 or top_k > 50:
            errors.append("Jumlah rekomendasi (K) harus antara 1 - 50.")
    except (ValueError, TypeError):
        errors.append("Jumlah rekomendasi (K) harus berupa angka bulat.")

    # Validasi pekerjaan
    valid_pekerjaan = ['buruh', 'asn', 'pengusaha', 'umum']
    pekerjaan = str(data.get('pekerjaan', 'umum')).strip().lower()
    if pekerjaan not in valid_pekerjaan:
        errors.append(f"Pekerjaan tidak valid. Pilihan: {', '.join(valid_pekerjaan)}.")

    if errors:
        return False, errors
    return True, []

# ===================== PROFILE MATCHING =====================
def profile_matching(df, preferences, weights):
    results = []
    for idx, row in df.iterrows():
        scores = {}

        # GAP Harga
        if preferences.get('max_price', 0) > 0:
            price_ratio = row['Price_Clean'] / preferences['max_price']
            if price_ratio <= 1:
                scores['harga'] = 5 - (1 - price_ratio) * 4
            else:
                scores['harga'] = max(1, 5 - (price_ratio - 1) * 5)
        else:
            scores['harga'] = 3

        # GAP Luas Bangunan
        if preferences.get('min_luas', 0) > 0:
            luas_ratio = row['Luas_bangunan'] / preferences['min_luas']
            scores['luas'] = min(5, max(1, luas_ratio * 3))
        else:
            scores['luas'] = 3

        # GAP Kamar Tidur
        pref_kamar = preferences.get('kamar_tidur', 0)
        if pref_kamar > 0:
            gap_kamar = int(row['Kamar_tidur_clean']) - pref_kamar
            gap_score_map = {0: 5, 1: 4.5, -1: 4, 2: 3.5, -2: 3}
            scores['kamar'] = gap_score_map.get(gap_kamar, max(1, 3 - abs(gap_kamar)))
        else:
            scores['kamar'] = 3

        # GAP Jenis Properti
        if preferences.get('jenis_properti'):
            scores['jenis'] = 5 if row['Jenis_Properti'] == preferences['jenis_properti'] else 2
        else:
            scores['jenis'] = 3

        # GAP Lokasi — Bertingkat: Kecamatan > Kota > Provinsi
        loc_score = 1
        row_provinsi = str(row.get('Provinsi', '')).strip().lower()
        row_kota = str(row.get('Kota_Kab', '')).strip().lower()
        row_kec = str(row.get('Kecamatan', '')).strip().lower()

        pref_provinsi = preferences.get('provinsi', '').strip().lower()
        pref_kota = preferences.get('kota', '').strip().lower()
        pref_kec = preferences.get('kecamatan', '').strip().lower()

        if pref_kec and row_kec == pref_kec:
            loc_score = 5
        elif pref_kota and row_kota == pref_kota:
            loc_score = 4
        elif pref_provinsi and row_provinsi == pref_provinsi:
            loc_score = 3
        elif not pref_provinsi and not pref_kota and not pref_kec:
            loc_score = 3
        scores['lokasi'] = loc_score

        # Skor akhir dengan bobot AHP
        criteria_keys = ['harga', 'luas', 'kamar', 'jenis', 'lokasi']
        final_score = sum(scores.get(k, 3) * weights[i] for i, k in enumerate(criteria_keys))

        results.append({
            'index': int(idx),
            'scores': scores,
            'final_score': round(final_score, 4)
        })

    return results

# ===================== EVALUATION METRICS =====================
def precision_at_k(recommended, relevant, k):
    rec_k = recommended[:k]
    return len(set(rec_k) & set(relevant)) / k if k > 0 else 0

def recall_at_k(recommended, relevant, k):
    rec_k = recommended[:k]
    return len(set(rec_k) & set(relevant)) / len(relevant) if relevant else 0

def f1_at_k(prec, rec):
    return 2 * prec * rec / (prec + rec) if (prec + rec) > 0 else 0

def ndcg_at_k(recommended, relevant, k):
    dcg = sum(1.0 / np.log2(i + 2) for i, idx in enumerate(recommended[:k]) if idx in relevant)
    idcg = sum(1.0 / np.log2(i + 2) for i in range(min(len(relevant), k)))
    return dcg / idcg if idcg > 0 else 0

def mrr(recommended, relevant):
    for i, idx in enumerate(recommended):
        if idx in relevant:
            return 1.0 / (i + 1)
    return 0

# ===================== API ROUTES =====================
@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/api/data-info', methods=['GET'])
def data_info():
    """Info dataset + struktur wilayah bertingkat: provinsi → kota → kecamatan"""
    wilayah = {}
    for _, row in df_properti[['Provinsi', 'Kota_Kab', 'Kecamatan']].dropna().iterrows():
        prov = str(row['Provinsi']).strip()
        kota = str(row['Kota_Kab']).strip()
        kec = str(row['Kecamatan']).strip()
        if not prov or prov == 'nan':
            continue
        wilayah.setdefault(prov, {})
        wilayah[prov].setdefault(kota, set())
        wilayah[prov][kota].add(kec)

    wilayah_sorted = {
        prov: {
            kota: sorted(kec_set)
            for kota, kec_set in sorted(kota_dict.items())
        }
        for prov, kota_dict in sorted(wilayah.items())
    }

    return jsonify({
        'total_properti': len(df_properti),
        'wilayah': wilayah_sorted,
        'provinsi_list': sorted(wilayah_sorted.keys()),
        'kota_list': sorted(df_properti['Kota_Kab'].dropna().unique().tolist()),
        'kecamatan_list': sorted(df_properti['Kecamatan'].dropna().unique().tolist()),
        'jenis_list': sorted(df_properti['Jenis_Properti'].dropna().unique().tolist()),
        'price_range': {
            'min': int(df_properti['Price_Clean'].min()),
            'max': int(df_properti['Price_Clean'].max())
        },
        'luas_range': {
            'min': int(df_properti['Luas_bangunan'].min()),
            'max': int(df_properti['Luas_bangunan'].max())
        }
    })

@app.route('/api/rekomendasi', methods=['POST'])
def get_rekomendasi():
    """Endpoint utama: rekomendasi dengan filter Provinsi → Kota → Kecamatan"""
    data = request.json

    # ── Validasi Input ───────────────────────────────────────
    is_valid, errors = validate_preferences(data)
    if not is_valid:
        return jsonify({'error': 'Input tidak valid.', 'detail': errors}), 400

    preferences = {
        'max_price': float(data.get('max_price', 0)),
        'min_luas': float(data.get('min_luas', 0)),
        'kamar_tidur': int(data.get('kamar_tidur', 0)),
        'jenis_properti': data.get('jenis_properti', ''),
        'provinsi': data.get('provinsi', ''),
        'kota': data.get('kota', ''),
        'kecamatan': data.get('kecamatan', ''),
        'pekerjaan': data.get('pekerjaan', 'Umum'),
    }

    top_k = int(data.get('top_k', 10))

    # ── Bobot AHP berdasarkan profesi ────────────────────────
    weights = get_weights_by_pekerjaan(preferences['pekerjaan'])
    ahp_result = {
        'kategori_konsumen': preferences['pekerjaan'],
        'weights': weights,
        'consistency_ratio': 0.0,
        'is_consistent': True
    }

    # ── HARD FILTER WILAYAH ──────────────────────────────────
    df_scoring = df_properti.copy()
    if preferences['provinsi'] and preferences['provinsi'].lower() not in ('', 'semua'):
        df_scoring = df_scoring[
            df_scoring['Provinsi'].str.lower() == preferences['provinsi'].lower()
        ]
    if preferences['kota'] and preferences['kota'].lower() not in ('', 'semua'):
        df_scoring = df_scoring[
            df_scoring['Kota_Kab'].str.lower() == preferences['kota'].lower()
        ]
    if preferences['kecamatan'] and preferences['kecamatan'].lower() not in ('', 'semua'):
        df_scoring = df_scoring[
            df_scoring['Kecamatan'].str.lower() == preferences['kecamatan'].lower()
        ]

    if df_scoring.empty:
        return jsonify({
            'recommendations': [],
            'ahp': ahp_result,
            'metrics': {
                'precision_at_k': 0, 'recall_at_k': 0,
                'f1_at_k': 0, 'ndcg_at_k': 0,
                'mrr': 0, 'k': top_k
            },
            'total_scored': 0,
            'pesan': 'Tidak ada properti ditemukan untuk wilayah yang dipilih.'
        })

    # ── Profile Matching ─────────────────────────────────────
    scores = profile_matching(df_scoring, preferences, weights)

    # ── Ranking ──────────────────────────────────────────────
    scores_sorted = sorted(scores, key=lambda x: x['final_score'], reverse=True)
    top_results = scores_sorted[:top_k]

    # ── Gabungkan dengan data properti ───────────────────────
    recommendations = []
    for item in top_results:
        row = df_properti.loc[item['index']]
        recommendations.append({
            'rank': len(recommendations) + 1,
            'harga': int(row['Price_Clean']),
            'luas_bangunan': int(row['Luas_bangunan']),
            'kamar_tidur': int(row['Kamar_tidur_clean']),
            'jenis_properti': str(row['Jenis_Properti']),
            'kecamatan': str(row.get('Kecamatan', '-')),
            'kota': str(row.get('Kota_Kab', '-')),
            'provinsi': str(row.get('Provinsi', '-')),
            'skor': item['final_score'],
            'detail_skor': item['scores']
        })

    # ── Metrik Evaluasi ───────────────────────────────────────
    all_indices = [s['index'] for s in scores_sorted]
    # Threshold relevansi diturunkan ke 3.0 untuk eksperimen/debug
    relevant_set = [s['index'] for s in scores_sorted if s['final_score'] >= 3.0]
    prec = precision_at_k(all_indices, relevant_set, top_k)
    rec = recall_at_k(all_indices, relevant_set, top_k)
    f1 = f1_at_k(prec, rec)
    ndcg = ndcg_at_k(all_indices, relevant_set, top_k)
    mrr_val = mrr(all_indices, relevant_set)

    return jsonify({
        'recommendations': recommendations,
        'ahp': ahp_result,
        'metrics': {
            'precision_at_k': round(prec, 4),
            'recall_at_k': round(rec, 4),
            'f1_at_k': round(f1, 4),
            'ndcg_at_k': round(ndcg, 4),
            'mrr': round(mrr_val, 4),
            'k': top_k
        },
        'total_scored': len(scores)
    })

# ===================== EVALUASI MASSAL (BAB 4) =====================
@app.route('/api/evaluasi-massal', methods=['POST'])
def evaluasi_massal():
    """
    Generate N konsumen sintetik dan evaluasi sistem secara massal.
    Digunakan untuk keperluan Bab 4 skripsi.
    Body: { "n_konsumen": 1000, "top_k": 10, "seed": 42 }
    """
    data = request.json or {}
    n_konsumen = int(data.get('n_konsumen', 1000))
    top_k = int(data.get('top_k', 10))
    seed = int(data.get('seed', 42))

    if n_konsumen < 1 or n_konsumen > 5000:
        return jsonify({'error': 'n_konsumen harus antara 1 - 5000'}), 400
    if top_k < 1 or top_k > 50:
        return jsonify({'error': 'top_k harus antara 1 - 50'}), 400

    np.random.seed(seed)
    random.seed(seed)

    # Ambil data wilayah yang tersedia dari dataset
    wilayah_tersedia = df_properti[['Provinsi', 'Kota_Kab', 'Kecamatan']].dropna().drop_duplicates()
    wilayah_list = wilayah_tersedia.values.tolist()

    jenis_list = df_properti['Jenis_Properti'].dropna().unique().tolist()
    pekerjaan_list = ['Buruh', 'ASN', 'Pengusaha', 'Umum']
    price_min = int(df_properti['Price_Clean'].min())
    price_max = int(df_properti['Price_Clean'].max())
    luas_min = int(df_properti['Luas_bangunan'].min())
    luas_max = int(df_properti['Luas_bangunan'].max())

    # Akumulasi metrik
    all_precision, all_recall, all_f1, all_ndcg, all_mrr = [], [], [], [], []
    konsumen_results = []

    for i in range(n_konsumen):
        # Generate preferensi acak
        pekerjaan = random.choice(pekerjaan_list)
        wilayah = random.choice(wilayah_list)
        pref_provinsi = wilayah[0]
        pref_kota = wilayah[1] if random.random() > 0.3 else ''
        pref_kec = wilayah[2] if random.random() > 0.5 else ''

        preferences = {
            'max_price': random.randint(price_min, price_max),
            'min_luas': random.randint(luas_min, min(luas_max, 300)),
            'kamar_tidur': random.choice([0, 1, 2, 3, 4]),
            'jenis_properti': random.choice([''] + jenis_list),
            'provinsi': pref_provinsi,
            'kota': pref_kota,
            'kecamatan': pref_kec,
            'pekerjaan': pekerjaan,
        }

        weights = get_weights_by_pekerjaan(pekerjaan)

        # Hard filter
        df_scoring = df_properti.copy()
        if pref_provinsi:
            df_scoring = df_scoring[df_scoring['Provinsi'].str.lower() == pref_provinsi.lower()]
        if pref_kota:
            df_scoring = df_scoring[df_scoring['Kota_Kab'].str.lower() == pref_kota.lower()]
        if pref_kec:
            df_scoring = df_scoring[df_scoring['Kecamatan'].str.lower() == pref_kec.lower()]

        if df_scoring.empty:
            print(f"DEBUG: no candidates for pref: {preferences}")
            continue

        scores = profile_matching(df_scoring, preferences, weights)
        scores_sorted = sorted(scores, key=lambda x: x['final_score'], reverse=True)
        all_indices = [s['index'] for s in scores_sorted]
        # Threshold relevansi diturunkan ke 3.0 untuk eksperimen/debug
        relevant_set = [s['index'] for s in scores_sorted if s['final_score'] >= 3.0]

        prec = precision_at_k(all_indices, relevant_set, top_k)
        rec = recall_at_k(all_indices, relevant_set, top_k)
        f1 = f1_at_k(prec, rec)
        ndcg = ndcg_at_k(all_indices, relevant_set, top_k)
        mrr_val = mrr(all_indices, relevant_set)

        all_precision.append(prec)
        all_recall.append(rec)
        all_f1.append(f1)
        all_ndcg.append(ndcg)
        all_mrr.append(mrr_val)

        konsumen_results.append({
            'konsumen_id': i + 1,
            'pekerjaan': pekerjaan,
            'provinsi': pref_provinsi,
            'kota': pref_kota,
            'kecamatan': pref_kec,
            'total_kandidat': len(df_scoring),
            'total_relevan': len(relevant_set),
            'precision': round(prec, 4),
            'recall': round(rec, 4),
            'f1': round(f1, 4),
            'ndcg': round(ndcg, 4),
            'mrr': round(mrr_val, 4),
        })

    n_valid = len(all_precision)
    if n_valid == 0:
        return jsonify({'error': 'Tidak ada konsumen valid yang bisa dievaluasi.'}), 400

    # Hitung rata-rata per pekerjaan
    df_results = pd.DataFrame(konsumen_results)
    per_pekerjaan = {}
    for pek in pekerjaan_list:
        subset = df_results[df_results['pekerjaan'] == pek]
        if len(subset) > 0:
            per_pekerjaan[pek] = {
                'n': len(subset),
                'avg_precision': round(subset['precision'].mean(), 4),
                'avg_recall': round(subset['recall'].mean(), 4),
                'avg_f1': round(subset['f1'].mean(), 4),
                'avg_ndcg': round(subset['ndcg'].mean(), 4),
                'avg_mrr': round(subset['mrr'].mean(), 4),
            }

    return jsonify({
        'n_konsumen_diminta': n_konsumen,
        'n_konsumen_valid': n_valid,
        'top_k': top_k,
        'seed': seed,
        'rata_rata_keseluruhan': {
            'precision_at_k': round(np.mean(all_precision), 4),
            'recall_at_k': round(np.mean(all_recall), 4),
            'f1_at_k': round(np.mean(all_f1), 4),
            'ndcg_at_k': round(np.mean(all_ndcg), 4),
            'mrr': round(np.mean(all_mrr), 4),
        },
        'std_deviasi': {
            'precision': round(np.std(all_precision), 4),
            'recall': round(np.std(all_recall), 4),
            'f1': round(np.std(all_f1), 4),
            'ndcg': round(np.std(all_ndcg), 4),
            'mrr': round(np.std(all_mrr), 4),
        },
        'per_pekerjaan': per_pekerjaan,
        'detail_konsumen': konsumen_results[:50]  # Tampilkan 50 pertama saja di response
    })

# ===================== RUN SERVER =====================
if __name__ == '__main__':
    print(f"Dataset loaded: {len(df_properti)} properti")
    print(f"Provinsi terdeteksi: {sorted(df_properti['Provinsi'].unique().tolist())}")
    print("Server running at http://localhost:5000")
    app.run(debug=True, port=5000)
