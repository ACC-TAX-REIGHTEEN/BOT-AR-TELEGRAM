import configparser
from datetime import datetime
import io
import os
import re
import threading
import time

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from rapidfuzz import fuzz, process
import telebot
from telebot import types

session_lock = threading.Lock()
data_lock = threading.Lock()

authenticated_users = set()
user_sessions = {}

G_DATA = {
    'df_ar': None,
    'pay_summary': {},
    'ml_dict': {},
    'ml_list': [],
    'fb_dict': {},
    'fb_list': [],
    'group_keywords': [],
    'branch_rules': {},
    'depo_config': "",
    'show_pay_status': True,
    'last_updated': None
}

def load_config(config_file='config.conf'):
    if not os.path.exists(config_file):
        raise ValueError(f"CRITICAL ERROR: File konfigurasi '{config_file}' tidak ditemukan!")
    
    config = configparser.ConfigParser()
    config.read(config_file, encoding='utf-8')

    secret_key = config.get('AUTH', 'secret_key', fallback='').strip()
    bot_token = config.get('BOT', 'bot_token', fallback='').strip()

    if not secret_key:
        raise ValueError("CRITICAL ERROR: 'secret_key' di seksi [AUTH] pada config.conf wajib diisi!")
    if not bot_token:
        raise ValueError("CRITICAL ERROR: 'bot_token' di seksi [BOT] pada config.conf wajib diisi!")

    return config

config = load_config()
SECRET_KEY = config.get('AUTH', 'secret_key').strip()
BOT_TOKEN = config.get('BOT', 'bot_token').strip()

bot = telebot.TeleBot(BOT_TOKEN)

def get_custom_rules(config):
    group_keywords = []
    if config.has_section("GROUP_KEYWORDS"):
        for _, val in config.items("GROUP_KEYWORDS"):
            if val.strip():
                group_keywords.append(val.strip().lower())

    branch_rules = {}
    if config.has_section("BRANCH_RULES"):
        for key, val in config.items("BRANCH_RULES"):
            if key.strip() and val.strip():
                branch_rules[key.strip().lower()] = val.strip()

    return group_keywords, branch_rules

def bersihkan_teks(teks):
    if pd.isna(teks):
        return ""
    t = str(teks).lower()
    t = re.sub(r"[^\w\s]", " ", t)
    return " ".join(t.split())

def potong_teks(teks, max_char=35):
    if pd.isna(teks):
        return ""
    s = str(teks).strip()
    return s[:max_char-3] + '...' if len(s) > max_char else s

def standardize_code(code, depo_prefixes="SL|YY|MKS|MGL|PW|PWT|PLU|SG|SMG|TGL|PA|KDI"):
    if pd.isna(code):
        return "" 
    if isinstance(code, float) and code.is_integer():
        code = int(code)    
    s = str(code).strip().upper()
    s = re.sub(r'\s*-\s*', '-', s)
    pattern = rf'^({depo_prefixes})\s*(\d+)'
    s = re.sub(pattern, r'\1-\2', s)
    s = re.sub(r'(\d+)([A-Z])$', r'\1 \2', s)
    return s

def read_excel_auto_header(file_path, sheet_name=0, target_column=""):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    target_clean = str(target_column).strip().upper()

    for idx, row in df_raw.iterrows():
        row_cleaned = [str(val).strip().upper() for val in row.dropna()]
        if target_clean in row_cleaned:
            df_clean = df_raw.iloc[idx + 1 :].copy()
            df_clean.columns = df_raw.iloc[idx].astype(str).str.strip()
            return df_clean.reset_index(drop=True)

    raise KeyError(f"Kolom target '{target_column}' tidak ditemukan pada file '{file_path}'")

def parse_date_sort(val):
    if pd.isna(val):
        return pd.NaT
    indo_months = {
        "mei": "may", "ags": "aug", "agt": "aug", "agu": "aug",
        "okt": "oct", "nop": "nov", "des": "dec", "peb": "feb"
    }
    val_str = str(val).lower().strip()
    for indo, eng in indo_months.items():
        if indo in val_str:
            val_str = val_str.replace(indo, eng)
            break
    return pd.to_datetime(val_str, errors="coerce", format="mixed")

def load_ml_and_fback_datasets(config):
    ml_file = config.get('DIR', 'hasil_latihan', fallback='Hasil_Latihan_temp.xlsx')
    fb_file = config.get('DIR', 'fback_cust', fallback='FBackCust_temp.xlsx')

    ml_dict, ml_list = {}, []
    if os.path.exists(ml_file):
        df_ml = pd.read_excel(ml_file)
        df_ml.columns = df_ml.columns.str.strip()
        if "Nama Customer dan Kota" in df_ml.columns and "Hasil_Nama_Rekomendasi" in df_ml.columns:
            for _, r in df_ml.iterrows():
                k = bersihkan_teks(r["Nama Customer dan Kota"])
                v = str(r["Hasil_Nama_Rekomendasi"]).strip()
                if k and v and v.upper() not in ["TIDAK DITEMUKAN", "FAILED", "NAN", "NONE", ""]:
                    ml_dict[k] = v
                    if k not in ml_list:
                        ml_list.append(k)

    fb_dict, fb_list = {}, []
    if os.path.exists(fb_file):
        df_fb = pd.read_excel(fb_file)
        df_fb.columns = df_fb.columns.str.strip()
        if "NO." in df_fb.columns:
            df_fb["NO."] = pd.to_numeric(df_fb["NO."], errors="coerce")
            df_fb = df_fb.sort_values(by="NO.", ascending=True).drop_duplicates(subset=["KETERANGAN"], keep="last")
        if "KETERANGAN" in df_fb.columns and "NAMA" in df_fb.columns:
            for _, row in df_fb.iterrows():
                ket = bersihkan_teks(row["KETERANGAN"])
                nama = str(row["NAMA"]).strip()
                if ket and nama:
                    fb_dict[ket] = nama
                    fb_list.append(ket)

    return ml_dict, ml_list, fb_dict, fb_list

def resolve_target_name_fast(raw_key, ml_dict, ml_list, fb_dict, fb_list, cache_resolver, branch_rules):
    raw_clean = bersihkan_teks(raw_key)
    if not raw_clean:
        return str(raw_key).strip()

    if raw_clean in cache_resolver:
        return cache_resolver[raw_clean]

    for rule_key, target_name in branch_rules.items():
        if "|" in rule_key:
            parent_kw, branch_kw = rule_key.split("|", 1)
            if parent_kw.strip() in raw_clean and branch_kw.strip() in raw_clean:
                cache_resolver[raw_clean] = target_name
                return target_name
        else:
            if rule_key in raw_clean:
                cache_resolver[raw_clean] = target_name
                return target_name

    if raw_clean in fb_dict:
        res = fb_dict[raw_clean]
        cache_resolver[raw_clean] = res
        return res

    if raw_clean in ml_dict:
        res = ml_dict[raw_clean]
        cache_resolver[raw_clean] = res
        return res

    if ml_list:
        match_ml = process.extractOne(query=raw_clean, choices=ml_list, scorer=fuzz.WRatio, score_cutoff=75.0)
        if match_ml:
            res = ml_dict[match_ml[0]]
            cache_resolver[raw_clean] = res
            return res

    if fb_list:
        match_fb = process.extractOne(query=raw_clean, choices=list(fb_dict.keys()), scorer=fuzz.WRatio, score_cutoff=80.0)
        if match_fb:
            res = fb_dict[match_fb[0]]
            cache_resolver[raw_clean] = res
            return res

    res = str(raw_key).strip()
    cache_resolver[raw_clean] = res
    return res

def load_payment_summary(pay_file):
    pay_summary = {}
    if os.path.exists(pay_file):
        df_pay = pd.read_excel(pay_file, sheet_name=0, header=1)
        df_pay.columns = df_pay.columns.str.strip()

        col_nota = [c for c in df_pay.columns if 'nota' in c.lower() or 'invoice' in c.lower() or 'faktur' in c.lower()]
        col_bayar = [c for c in df_pay.columns if 'bayar' in c.lower() or 'nominal' in c.lower()]

        if col_nota and col_bayar:
            c_n, c_b = col_nota[0], col_bayar[0]
            for _, row in df_pay.dropna(subset=[c_n]).iterrows():
                nota = str(row[c_n]).strip()
                try:
                    nominal = float(row[c_b])
                except (ValueError, TypeError):
                    nominal = 0.0

                if nota not in pay_summary:
                    pay_summary[nota] = {'total_bayar': 0.0, 'count': 0}
                pay_summary[nota]['total_bayar'] += nominal
                pay_summary[nota]['count'] += 1
    return pay_summary

def get_payment_status(no_faktur, nilai_faktur, pay_summary):
    no_faktur_str = str(no_faktur).strip()
    try:
        nilai_faktur_val = float(nilai_faktur)
    except (ValueError, TypeError):
        nilai_faktur_val = 0.0

    if no_faktur_str in pay_summary:
        total_bayar = pay_summary[no_faktur_str]['total_bayar']
        jml_data = pay_summary[no_faktur_str]['count']

        if total_bayar == nilai_faktur_val:
            return f"LUNAS (Total {jml_data}x bayar)", "blue"
        elif total_bayar < nilai_faktur_val:
            kurang = nilai_faktur_val - total_bayar
            return f"DICICIL: Baru Bayar {total_bayar:,.0f} (Kurang: {kurang:,.0f})".replace(",", "."), "orange"
        else:
            return f"LEBIH BAYAR: Total {total_bayar:,.0f}".replace(",", "."), "magenta"
    else:
        return "Belum Ada Data", "black"

def load_ar_dataset_from_disk(config):
    ar_file = config.get('DIR', 'ar_clean', fallback='ARClean_temp.xlsx')
    pay_file = config.get('DIR', 'pay_ss', fallback='PaySS_temp.xlsx')
    depo_config = config.get('MAP', 'depo', fallback='SL|YY|MKS|MGL|PW|PWT|PLU|SG|SMG|TGL|PA|KDI').strip()
    show_pay_str = config.get('DISPLAY', 'show_pay_status', fallback='Ya').strip().lower()
    show_pay_status = show_pay_str in ['ya', 'yes', '1', 'true']

    df_ar = read_excel_auto_header(ar_file, sheet_name=0, target_column="Nama Pelanggan")

    if 'No. Pelanggan' in df_ar.columns:
        df_ar['Clean_Kode'] = df_ar['No. Pelanggan'].apply(lambda x: standardize_code(x, depo_config))
    elif 'Kode Pelanggan' in df_ar.columns:
        df_ar['Clean_Kode'] = df_ar['Kode Pelanggan'].apply(lambda x: standardize_code(x, depo_config))
    else:
        df_ar['Clean_Kode'] = ""

    pay_summary = load_payment_summary(pay_file)
    ml_dict, ml_list, fb_dict, fb_list = load_ml_and_fback_datasets(config)
    group_keywords, branch_rules = get_custom_rules(config)

    return df_ar, pay_summary, ml_dict, ml_list, fb_dict, fb_list, group_keywords, branch_rules, depo_config, show_pay_status

def background_data_refresher(config):
    while True:
        try:
            df_ar, pay_summary, ml_dict, ml_list, fb_dict, fb_list, group_keywords, branch_rules, depo_config, show_pay_status = load_ar_dataset_from_disk(config)
            
            with data_lock:
                G_DATA['df_ar'] = df_ar
                G_DATA['pay_summary'] = pay_summary
                G_DATA['ml_dict'] = ml_dict
                G_DATA['ml_list'] = ml_list
                G_DATA['fb_dict'] = fb_dict
                G_DATA['fb_list'] = fb_list
                G_DATA['group_keywords'] = group_keywords
                G_DATA['branch_rules'] = branch_rules
                G_DATA['depo_config'] = depo_config
                G_DATA['show_pay_status'] = show_pay_status
                G_DATA['last_updated'] = datetime.now()

            print(f"--> [{datetime.now().strftime('%H:%M:%S')}] RAM Data Caching Sukses Diperbarui!")
        except Exception as err:
            print(f"--> [ERROR CACHE REFRESH]: {err}")

        interval_minutes = int(config.get('AR', 'ar_time_interval', fallback=10))
        time.sleep(interval_minutes * 60)

def generate_ar_image(df_filtered, pay_summary, filter_jt=False, show_pay_status=True):
    df = df_filtered.copy()

    if 'No. Faktur' in df.columns:
        df = df.sort_values(by='No. Faktur', ascending=True).reset_index(drop=True)

    def parse_umur_jt(val):
        match = re.search(r'(-?\d+)', str(val))
        return int(match.group(1)) if match else 0

    df['Umur_JT_Num'] = df['Umur JT'].apply(parse_umur_jt)

    if filter_jt:
        df = df[df['Umur_JT_Num'] > 0].reset_index(drop=True)

    if df.empty:
        return None

    status_list, color_list = [], []
    for _, r in df.iterrows():
        st, col = get_payment_status(r.get('No. Faktur', ''), r.get('Nilai Faktur', 0), pay_summary)
        status_list.append(st)
        color_list.append(col)

    df['Cek Pelunasan SS Sales'] = status_list

    if show_pay_status:
        cols_to_show = [
            'No. Faktur', 'Tgl Faktur', 'Jatuh Tempo', 'Nilai Faktur', 
            'Sisa Piutang', 'Umur JT', 'Nama Pelanggan', 'Nama Penjual', 
            'Nama Kontak', 'Cek Pelunasan SS Sales'
        ]
        col_widths = [0.08, 0.07, 0.07, 0.08, 0.08, 0.06, 0.16, 0.09, 0.18, 0.13]
    else:
        cols_to_show = [
            'No. Faktur', 'Tgl Faktur', 'Jatuh Tempo', 'Nilai Faktur', 
            'Sisa Piutang', 'Umur JT', 'Nama Pelanggan', 'Nama Penjual', 
            'Nama Kontak'
        ]
        col_widths = [0.09, 0.08, 0.08, 0.09, 0.09, 0.06, 0.20, 0.10, 0.21]

    valid_cols = [c for c in cols_to_show if c in df.columns]
    df_display = df[valid_cols].copy()

    if 'Nama Kontak' in df_display.columns:
        df_display['Nama Kontak'] = df_display['Nama Kontak'].apply(lambda x: potong_teks(x, max_char=35))
    if 'Nama Pelanggan' in df_display.columns:
        df_display['Nama Pelanggan'] = df_display['Nama Pelanggan'].apply(lambda x: potong_teks(x, max_char=30))
    if 'Nama Penjual' in df_display.columns:
        df_display['Nama Penjual'] = df_display['Nama Penjual'].apply(lambda x: potong_teks(x, max_char=16))

    for col in ['Nilai Faktur', 'Sisa Piutang']:
        if col in df_display.columns:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,.0f}".replace(",", ".") if pd.notna(x) else "0")

    months = [parse_date_sort(val) for val in df['Tgl Faktur']]
    final_rows = []
    final_color_list = []

    for i in range(len(df_display)):
        if i > 0:
            prev_m = (months[i-1].year, months[i-1].month) if pd.notna(months[i-1]) else None
            curr_m = (months[i].year, months[i].month) if pd.notna(months[i]) else None
            if prev_m and curr_m and prev_m != curr_m:
                final_rows.append([""] * len(valid_cols))
                final_color_list.append("black")

        final_rows.append(df_display.iloc[i].tolist())
        final_color_list.append(color_list[i])

    df_final_display = pd.DataFrame(final_rows, columns=valid_cols)

    total_nilai = df['Nilai Faktur'].sum() if 'Nilai Faktur' in df.columns else 0
    total_sisa = df['Sisa Piutang'].sum() if 'Sisa Piutang' in df.columns else 0
    total_jt = df[df['Umur_JT_Num'] > 0]['Sisa Piutang'].sum() if 'Sisa Piutang' in df.columns else 0

    num_rows = len(df_final_display)
    width_inches = 18.0
    height_inches = max(1.8, num_rows * 0.32 + 1.0)
    
    target_dpi = 220
    max_pixels_limit = 9000

    if (width_inches * target_dpi) > max_pixels_limit or (height_inches * target_dpi) > max_pixels_limit:
        target_dpi = int(max_pixels_limit / max(width_inches, height_inches))

    fig, ax = plt.subplots(figsize=(width_inches, height_inches))
    ax.axis('off')

    table = ax.table(
        cellText=df_final_display.values, 
        colLabels=df_final_display.columns, 
        colWidths=col_widths,
        bbox=[0.0, 0.11, 1.0, 0.81],
        cellLoc='left'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9.5)

    for (row_idx, col_idx), cell in table.get_celld().items():
        if row_idx == 0:
            cell.set_facecolor('#002060')
            cell.set_text_props(color='white', weight='bold')
        else:
            if show_pay_status and df_final_display.columns[col_idx] == 'Cek Pelunasan SS Sales':
                cell.set_text_props(color=final_color_list[row_idx - 1], weight='bold')

    plt.figtext(0.5, 0.94, "PT PRIMA TUNGGAL MANDIRI", ha="center", va="bottom", fontsize=15, fontweight='bold', color='#002060')

    summary_title = f"Total Nilai Faktur: Rp {total_nilai:,.0f}  |  Total Sisa Piutang: Rp {total_sisa:,.0f}".replace(",", ".")
    if not filter_jt:
        summary_title += f"  |  Total Jatuh Tempo (JT): Rp {total_jt:,.0f}".replace(",", ".")
    elif filter_jt:
        summary_title += f"  |  Total Jatuh Tempo (JT): Rp {total_sisa:,.0f}".replace(",", ".")

    plt.figtext(0.5, 0.04, summary_title, ha="center", va="top", fontsize=11, fontweight='bold', color='#002060')

    img_buffer = io.BytesIO()
    plt.savefig(img_buffer, format='png', bbox_inches='tight', pad_inches=0.2, dpi=target_dpi)
    plt.close(fig)
    img_buffer.seek(0)

    return img_buffer

@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    user_id = message.from_user.id
    with session_lock:
        is_auth = user_id in authenticated_users

    if not is_auth:
        bot.reply_to(message, "Akses Terkunci.\nMasukkan Sandi/Token Internal untuk mengaktifkan bot:")
    else:
        bot.reply_to(message, "Bot AR Ready!\nMasukkan Nama Pelanggan, Kode Pelanggan, atau Kata Kunci Sales (Contoh: Wakid Kendal, 10000, YY-2223, 10000 & YY-2223).")

@bot.message_handler(func=lambda msg: True)
def handle_incoming_messages(message):
    user_id = message.from_user.id

    with session_lock:
        is_auth = user_id in authenticated_users

    if not is_auth:
        if message.text.strip() == SECRET_KEY:
            with session_lock:
                authenticated_users.add(user_id)
            bot.reply_to(message, "Akses Diterima! Silakan cari data piutang pelanggan.")
        else:
            bot.reply_to(message, "Sandi Salah! Silakan coba lagi.")
        return

    query = message.text.strip()

    with data_lock:
        df_ar = G_DATA['df_ar']
        pay_summary = G_DATA['pay_summary']
        ml_dict = G_DATA['ml_dict']
        ml_list = G_DATA['ml_list']
        fb_dict = G_DATA['fb_dict']
        fb_list = G_DATA['fb_list']
        group_keywords = G_DATA['group_keywords']
        branch_rules = G_DATA['branch_rules']
        depo_config = G_DATA['depo_config']

    if df_ar is None:
        bot.reply_to(message, "Data sedang dimuat ke memori RAM, silakan coba beberapa detik lagi.")
        return

    matched_df = pd.DataFrame()
    cache_resolver = {}

    clean_q = query.lower().replace('nopel:', '').strip()
    raw_codes = [k.strip() for k in clean_q.split('&') if k.strip()]
    std_codes = [standardize_code(k, depo_config) for k in raw_codes]

    matched_df = df_ar[df_ar['Clean_Kode'].isin(std_codes)]

    if matched_df.empty:
        nama_resmi = resolve_target_name_fast(query, ml_dict, ml_list, fb_dict, fb_list, cache_resolver, branch_rules)
        target_clean = bersihkan_teks(nama_resmi)

        matched_group = None
        for g_kw in group_keywords:
            if g_kw in target_clean or g_kw in bersihkan_teks(query):
                matched_group = g_kw
                break

        if matched_group:
            cond_group = (
                df_ar['Nama Pelanggan'].astype(str).apply(bersihkan_teks).str.contains(rf"\b{re.escape(matched_group)}\b", regex=True, na=False) |
                df_ar['Nama Kontak'].astype(str).apply(bersihkan_teks).str.contains(rf"\b{re.escape(matched_group)}\b", regex=True, na=False)
            )
            matched_df = df_ar[cond_group]
        else:
            pelanggan_clean = df_ar['Nama Pelanggan'].astype(str).apply(bersihkan_teks) if 'Nama Pelanggan' in df_ar.columns else pd.Series("", index=df_ar.index)
            kontak_clean = df_ar['Nama Kontak'].astype(str).apply(bersihkan_teks) if 'Nama Kontak' in df_ar.columns else pd.Series("", index=df_ar.index)

            cond_exact = (pelanggan_clean == target_clean) | (kontak_clean == target_clean)
            matched_df = df_ar[cond_exact]

            if matched_df.empty:
                names_list = df_ar['Nama Pelanggan'].dropna().unique().tolist()
                best_match = process.extractOne(nama_resmi, names_list, scorer=fuzz.token_sort_ratio, score_cutoff=80.0)
                if best_match:
                    matched_df = df_ar[df_ar['Nama Pelanggan'] == best_match[0]]

    if matched_df.empty:
        bot.reply_to(message, f"Data piutang tidak ditemukan untuk kata kunci: '{query}'.")
        return

    with session_lock:
        user_sessions[user_id] = {
            'data': matched_df,
            'pay_summary': pay_summary
        }

    markup = types.InlineKeyboardMarkup(row_width=3)
    markup.add(
        types.InlineKeyboardButton("IRC", callback_data="prod_IRC"),
        types.InlineKeyboardButton("ZN", callback_data="prod_ZN"),
        types.InlineKeyboardButton("SEMUA", callback_data="prod_ALL")
    )

    bot.send_message(message.chat.id, f"Ditemukan {len(matched_df)} faktur piutang.\n\nLangkah 1/3: Pilih Filter Produk:", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('prod_'))
def process_product_filter(call):
    user_id = call.from_user.id
    with session_lock:
        if user_id not in user_sessions:
            bot.answer_callback_query(call.id, "Sesi berakhir. Silakan cari ulang.")
            return
        prod = call.data.split('_')[1]
        user_sessions[user_id]['prod'] = prod

    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton("HANYA JT", callback_data="jt_YES"),
        types.InlineKeyboardButton("SEMUA DATA", callback_data="jt_NO")
    )

    bot.edit_message_text("Langkah 2/3: Pilih Filter Jatuh Tempo (JT):", chat_id=call.message.chat.id, message_id=call.message.message_id, reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('jt_'))
def process_jt_filter(call):
    user_id = call.from_user.id
    with session_lock:
        if user_id not in user_sessions:
            bot.answer_callback_query(call.id, "Sesi berakhir. Silakan cari ulang.")
            return
        is_jt = call.data == "jt_YES"
        user_sessions[user_id]['filter_jt'] = is_jt

    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
        types.InlineKeyboardButton("TANPA FRAUD", callback_data="fraud_NO"),
        types.InlineKeyboardButton("SERTAKAN FRAUD", callback_data="fraud_YES")
    )

    bot.edit_message_text("Langkah 3/3: Sertakan data Sales FRAUD?", chat_id=call.message.chat.id, message_id=call.message.message_id, reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data.startswith('fraud_'))
def process_fraud_filter(call):
    user_id = call.from_user.id
    with session_lock:
        if user_id not in user_sessions:
            bot.answer_callback_query(call.id, "Sesi berakhir. Silakan cari ulang.")
            return
        session = user_sessions[user_id]

    include_fraud = call.data == "fraud_YES"
    df_data = session['data'].copy()
    prod = session.get('prod', 'ALL')
    is_jt = session.get('filter_jt', False)

    with data_lock:
        show_pay_status = G_DATA['show_pay_status']

    if prod != 'ALL':
        cond_k = df_data['Nama Kontak'].astype(str).str.contains(prod, case=False, na=False)
        cond_p = df_data['Nama Penjual'].astype(str).str.contains(prod, case=False, na=False)
        df_data = df_data[cond_k | cond_p]

    if not include_fraud and 'Nama Penjual' in df_data.columns:
        df_data = df_data[~df_data['Nama Penjual'].astype(str).str.contains('FRAUD', case=False, na=False)]

    bot.edit_message_text("Mengolah tabel dan meng-generate gambar laporan...", chat_id=call.message.chat.id, message_id=call.message.message_id)

    img_buffer = generate_ar_image(df_data, session['pay_summary'], filter_jt=is_jt, show_pay_status=show_pay_status)

    if img_buffer:
        caption_msg = (
            f"Laporan Piutang ({prod})\n"
            f"• Status JT: {'Hanya JT (>0 hari)' if is_jt else 'Semua Faktur'}\n"
            f"• Sales FRAUD: {'Disertakan' if include_fraud else 'Dibuang (Tanpa Fraud)'}"
        )
        try:
            bot.send_photo(call.message.chat.id, img_buffer, caption=caption_msg)
        except Exception as e_send:
            img_buffer.seek(0)
            img_buffer.name = f"Laporan_Piutang_{prod}.png"
            bot.send_document(call.message.chat.id, img_buffer, caption=caption_msg)
    else:
        bot.send_message(call.message.chat.id, "Tidak ada data piutang setelah disaring.")

if __name__ == '__main__':
    print("--> Memulai Data RAM Preloader...")
    df_ar, pay_summary, ml_dict, ml_list, fb_dict, fb_list, group_keywords, branch_rules, depo_config, show_pay_status = load_ar_dataset_from_disk(config)
    
    with data_lock:
        G_DATA['df_ar'] = df_ar
        G_DATA['pay_summary'] = pay_summary
        G_DATA['ml_dict'] = ml_dict
        G_DATA['ml_list'] = ml_list
        G_DATA['fb_dict'] = fb_dict
        G_DATA['fb_list'] = fb_list
        G_DATA['group_keywords'] = group_keywords
        G_DATA['branch_rules'] = branch_rules
        G_DATA['depo_config'] = depo_config
        G_DATA['show_pay_status'] = show_pay_status
        G_DATA['last_updated'] = datetime.now()

    refresher_thread = threading.Thread(target=background_data_refresher, args=(config,), daemon=True)
    refresher_thread.start()

    print("--> Bot Telegram Piutang AR Aktif & Berjalan...")
    bot.infinity_polling()