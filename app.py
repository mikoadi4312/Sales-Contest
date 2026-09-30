from pathlib import Path
from io import BytesIO
import hashlib
import pickle
import re

import pandas as pd
import streamlit as st
import portal_ui as ui

from contest_engine import (read_template, read_revenue, calculate, resume, export_results,
                            FILTER_COLUMNS, clean, identity, store_id, validate_mapping)

ROOT = Path(__file__).resolve().parent
LOCAL_TEMPLATE = ROOT / 'Template Kriteria Upload Sales Contes.xlsx'
TEMPLATE = LOCAL_TEMPLATE if LOCAL_TEMPLATE.exists() else ROOT / 'assets' / 'Template Kriteria Sales Contest.xlsx'
REVENUE = ROOT / 'REVENUEERA 1-31 Agustus 2026.xlsx'
st.set_page_config(page_title='EraBlue | Sales Contest', page_icon=str(ROOT / 'erablue-logo.png'), layout='wide')
ui.apply_theme()


@st.cache_data(show_spinner=False, max_entries=6)
def load_revenue(content):
    return read_revenue(content)


@st.cache_data(show_spinner=False, max_entries=30)
def load_template(content):
    return read_template(content)


@st.cache_data(show_spinner=False, max_entries=12)
def mapping_sheet(content, sheet, header):
    return pd.read_excel(BytesIO(content), sheet_name=sheet, header=header - 1, dtype=str).fillna('')


@st.cache_data(show_spinner=False)
def mapping_example():
    stream = BytesIO()
    with pd.ExcelWriter(stream, engine='openpyxl') as writer:
        pd.DataFrame([{'Store ID': 'CONTOH001', 'Store Name': 'Toko contoh (ganti data ini)',
                       'AM': 'Nama AM', 'RSM': 'Nama RSM'}]).to_excel(writer, index=False, sheet_name='Mapping')
        pd.DataFrame({'Petunjuk': [
            'Ganti baris contoh dengan mapping toko aktual.',
            'Satu toko hanya boleh memiliki satu AM.',
            'Untuk beberapa penerima RSM/Asst./Admin RSM, ulangi Store ID dan AM dengan nama penerima RSM berbeda.',
            'Gunakan nama AM dan RSM yang sama dengan Criteria.',
            'File AM dan RSM terpisah juga dapat diunggah lalu digabung berdasarkan Store ID.']}).to_excel(writer, index=False, sheet_name='Petunjuk')
    return stream.getvalue()


def rupiah(value):
    return 'Rp' + f'{value:,.0f}'.replace(',', '.')


def show_table(frame):
    view = frame.copy()
    if 'Pencapaian' in view:
        view['Pencapaian (%)'] = view.pop('Pencapaian') * 100
    st.dataframe(view, hide_index=True, width='stretch', column_config={
        'Pencapaian (%)': st.column_config.NumberColumn(format='%.2f%%'),
        'Reward': st.column_config.NumberColumn(format='localized'),
        'Total': st.column_config.NumberColumn(format='localized')})


def suggest_filters(template, revenue):
    # Only suggest explicit terms in the scheme; users review all filter settings.
    text = template['scheme'].casefold()
    suggestions = {col: [] for col in FILTER_COLUMNS}
    for col in ['Main group', 'Sub group']:
        for value in sorted(revenue[col].unique()):
            term = re.sub(r'^indo\s*-\s*', '', value, flags=re.I).strip().casefold()
            if term and re.search(r'(?<!\w)' + re.escape(term) + r'(?!\w)', text):
                suggestions[col].append(value)
    if 'all brand' not in text:
        suggestions['Brand name'] = [v for v in sorted(revenue['Brand name'].unique())
                                     if v and re.search(r'(?<!\w)' + re.escape(v.casefold()) + r'(?!\w)', text)]
    suggestions['Inventory status'] = [v for v in sorted(revenue['Inventory status'].unique()) if v.casefold() in text]
    return suggestions


with st.sidebar:
    ui.sidebar_brand()
    modes = ['Upload file saya']
    if LOCAL_TEMPLATE.exists() and REVENUE.exists():
        modes.append('Coba contoh Agustus 2026')
    mode = st.radio('Sumber data', modes)
    st.divider()
    st.subheader('Pusat template')
    if TEMPLATE.exists():
        ui.template_download('Template kriteria', TEMPLATE.read_bytes(), TEMPLATE.name)
    ui.template_download('Contoh format mapping', mapping_example(), 'Template Mapping Toko AM RSM.xlsx')
    st.caption('Klik untuk menyimpan file Excel. Periksa daftar download browser (Ctrl+J).')
    st.caption('Satu template Criteria + Result untuk satu kategori. Jumlah kategori mengikuti file Anda.')
    st.divider()
    with st.expander('Panduan singkat', icon=':material/help_outline:'):
        st.write('1. Upload template dan Revenue.\n2. Lengkapi mapping toko.\n3. Periksa kriteria, lalu hitung.\n4. Download hasil Excel.')
        st.caption('File upload dan pengaturan berlaku pada sesi browser ini. Unduh hasil sebelum menutup sesi.')


demo = mode.startswith('Coba')
files = []
revenue_bytes = None
mapping_files = []
ui.section_title('upload-data', '01', 'Siapkan data contest', 'Satu template untuk setiap kategori. Tambahkan sebanyak yang Anda perlukan.')
with st.container(border=True, key='upload_section'):
    if demo:
        st.info('Mode contoh menggunakan template Laptop dan Revenue Agustus yang tersedia. Mapping AM dibaca dari Result. Mapping RSM belum tersedia, sehingga bonus RSM akan kosong sampai dilengkapi.')
        if not TEMPLATE.exists() or not REVENUE.exists():
            st.error('File contoh tidak ditemukan di folder aplikasi. Gunakan mode upload.')
            st.stop()
        files = [(TEMPLATE.name, TEMPLATE.read_bytes())]
        revenue_bytes = REVENUE.read_bytes()
    else:
        left, right, mapping_col = st.columns(3, gap='medium')
        with left:
            ui.upload_label('Template kategori', 'Criteria & Result · Bisa beberapa file')
            uploaded = st.file_uploader('Template kategori (.xlsx)', type=['xlsx'], accept_multiple_files=True,
                                       label_visibility='collapsed', help='Setiap file memiliki sheet Criteria dan Result, mengikuti template download.')
            files = [(f.name, f.getvalue()) for f in uploaded]
        with right:
            ui.upload_label('Revenue penjualan', 'Data transaksi untuk menghitung SO')
            uploaded_revenue = st.file_uploader('Revenue (.xlsx)', type=['xlsx'], label_visibility='collapsed')
            if uploaded_revenue:
                revenue_bytes = uploaded_revenue.getvalue()
    with (st.container() if demo else mapping_col):
        ui.upload_label('Mapping AM & RSM', 'File gabungan atau file terpisah')
        mapping_files = st.file_uploader('Mapping toko–AM–RSM (.xlsx), dapat lebih dari satu file', type=['xlsx'],
                                         label_visibility='collapsed', accept_multiple_files=True, help='Bisa satu file gabungan atau file AM dan RSM terpisah. Pilih kolom setelah upload.')
    st.caption(f"Format .xlsx · Maksimal {st.get_option('server.maxUploadSize')} MB per file · Template tersedia di panel kiri")

if not files or revenue_bytes is None:
    ui.pending_sections()
    ui.footer()
    st.stop()

try:
    with st.spinner('Membaca template dan Revenue. File besar dapat memerlukan beberapa saat…'):
        templates = [load_template(content) for _, content in files]
        revenue = load_revenue(revenue_bytes)
    if revenue.empty:
        st.error('Revenue tidak berisi transaksi.'); st.stop()
except Exception as exc:
    st.error(f'File belum dapat dibaca: {exc}')
    st.stop()

a, b, c = st.columns(3)
a.metric('Kategori diunggah', len(files))
b.metric('Baris Revenue', f'{len(revenue):,}'.replace(',', '.'))
c.metric('Toko dalam Revenue', revenue['Store code'].nunique())

with st.expander('Mapping toko, AM, dan RSM', expanded=bool(mapping_files)):
    if mapping_files:
        parts = []
        try:
            for file in mapping_files:
                content = file.getvalue()
                token = hashlib.sha256(content).hexdigest()[:10] + file.name
                st.markdown(f'**{file.name}**')
                workbook = pd.ExcelFile(BytesIO(content), engine='openpyxl')
                p, q = st.columns([2, 1])
                sheet = p.selectbox('Sheet mapping', workbook.sheet_names, key='sheet' + token)
                header = q.number_input('Baris header', 1, 100, 1, key='header' + token)
                raw = mapping_sheet(content, sheet, header)
                options = ['— Tidak ada —'] + list(raw.columns)
                def column_pick(label, aliases, container):
                    found = next((c for c in raw.columns if identity(c) in aliases), None)
                    return container.selectbox(label, options, index=options.index(found) if found is not None else 0,
                                               key=label + token)
                x, y, z = st.columns(3)
                sc = column_pick('Kolom kode toko', ['storeid', 'storecode', 'kodetoko', 'idstore'], x)
                ac = column_pick('Kolom AM', ['am', 'amname', 'areamanager', 'namaam'], y)
                rc = column_pick('Kolom RSM', ['rsm', 'rsmname', 'regionalmanager', 'regional sales manager'.replace(' ', ''), 'namarsm'], z)
                if sc == options[0]:
                    st.warning('Pilih kolom kode toko untuk setiap file.'); st.stop()
                frame = pd.DataFrame({'Store ID': raw[sc].map(store_id)})
                frame['AM'] = raw[ac].map(clean) if ac != options[0] else ''
                frame['RSM'] = raw[rc].map(clean) if rc != options[0] else ''
                parts.append(frame)
            combined = pd.concat(parts, ignore_index=True)
            ams = combined.loc[combined.AM != '', ['Store ID', 'AM']].drop_duplicates()
            rsms = combined.loc[combined.RSM != '', ['Store ID', 'RSM']].drop_duplicates()
            ids = combined[['Store ID']].drop_duplicates()
            mapping = ids.merge(ams, on='Store ID', how='left').merge(rsms, on='Store ID', how='left').fillna('')
            mapping = validate_mapping(mapping)
        except Exception as exc:
            st.error(f'Mapping belum valid: {exc}'); st.stop()
    else:
        mapping = pd.concat([t['mapping'] for t in templates], ignore_index=True)
        st.warning('Belum ada file mapping. Untuk mencoba, AM menggunakan nama pada sheet Result; RSM belum dihitung. Upload mapping aktual untuk hasil periode berjalan.')
    mapping = st.data_editor(mapping, num_rows='dynamic', hide_index=True, width='stretch',
                             key='mapping_editor_' + hashlib.sha256(pickle.dumps(mapping)).hexdigest()[:12])
    st.caption('Satu toko dapat memiliki beberapa penerima RSM/Asst. RSM: ulangi kode toko dengan AM yang sama dan nama RSM berbeda. SO dihitung sekali per toko untuk setiap penerima.')

ui.section_title('kriteria', '02', 'Periode & kriteria kategori', 'Sesuaikan filter transaksi dan skema reward dengan ketentuan contest Anda.')
with st.container(border=True, key='criteria_section'):
    left, right = st.columns(2)
    start = left.date_input('Tanggal mulai', revenue.Date.min().date())
    end = right.date_input('Tanggal akhir', revenue.Date.max().date())
    if start > end:
        st.error('Tanggal mulai harus sebelum atau sama dengan tanggal akhir.'); st.stop()
    st.caption('Aturan angka dibaca dari Criteria. Teks Scheme ditampilkan untuk diperiksa; ubah filter dan aturan di bawah jika diperlukan. Aplikasi menghitung pencapaian SO ÷ target, bukan pertumbuhan dari bulan sebelumnya.')
    configs = []
    for index, ((filename, content), template) in enumerate(zip(files, templates)):
        token = f'{index}_' + hashlib.sha256(content).hexdigest()[:12]
        with st.expander(f'{index + 1}. {filename}', expanded=len(files) == 1):
            config = {'start': start, 'end': end}
            a, b = st.columns([3, 1])
            config['name'] = a.text_input('Nama kategori', template['category'], key='name' + token).strip()
            config['metric'] = b.selectbox('Dasar SO', ['Revenue', 'Quantity'], key='metric' + token)
            st.caption('Target harus menggunakan satuan yang sama dengan dasar SO.')
            st.text_area('Scheme dari Criteria', template['scheme'], disabled=True, height=140, key='scheme' + token)
            defaults = suggest_filters(template, revenue)
            config['filters'] = {}
            st.markdown('**Filter transaksi**')
            st.caption('Pilihan kosong berarti semua nilai. Antar-kolom berlaku AND, beberapa pilihan dalam satu kolom berlaku OR. Pilihan awal hanya saran dari teks Scheme.')
            columns = st.columns(3)
            for i, field in enumerate(FILTER_COLUMNS):
                config['filters'][field] = columns[i % 3].multiselect(field, sorted(revenue[field].unique()),
                            default=defaults[field], key=field + token)
            if 'kredivo' in (config['name'] + template['scheme']).casefold():
                st.warning('Revenue contoh tidak memisahkan penyedia pembiayaan. Skema rasio Kredivo / seluruh financing belum didukung versi ini. Jangan gunakan filter cicilan umum sebagai pengganti transaksi Kredivo.')
            st.markdown('**Reward toko**')
            a, b = st.columns(2)
            config['minimum'] = a.number_input('Minimum pencapaian toko (%)', min_value=0., value=100., step=5., key='minimum' + token)
            config['ties'] = b.selectbox('Jika pencapaian sama', ['Peringkat sama (1, 1, 3)', 'Urutan Store ID'], key='ties' + token,
                                         help='Peringkat sama dapat membuat penerima hadiah melebihi kuota. Urutan Store ID membatasi kuota dengan kode toko sebagai pembeda.')
            config['ranks'] = st.data_editor(template['ranks'], num_rows='dynamic', hide_index=True, width='stretch', key='ranks' + token)
            st.markdown('**Reward AM dan RSM**')
            a, b = st.columns(2)
            config['manager_scope'] = a.selectbox('Cakupan SO AM/RSM', ['Toko peserta kategori', 'Seluruh toko dalam mapping'], key='scope' + token)
            config['count_source'] = b.selectbox('Jumlah toko untuk tier AM', ['Criteria', 'Mapping'], key='count' + token)
            config['am_rules'] = st.data_editor(template['am_rules'], num_rows='dynamic', hide_index=True, width='stretch', key='amrules' + token)
            st.caption('Reward 100% dipilih terlebih dahulu; jika belum tercapai, digunakan tier 90%. Target nominal tiap tier ada di tabel berikut.')
            st.markdown('Target AM')
            config['am'] = st.data_editor(template['am'], num_rows='dynamic', hide_index=True, width='stretch', key='am' + token)
            st.markdown('Target dan reward RSM / Asst. RSM / Admin RSM')
            config['rsm'] = st.data_editor(template['rsm'], num_rows='dynamic', hide_index=True, width='stretch', key='rsm' + token)
            st.caption('Nama penerima harus sama dengan mapping. Reward yang tidak terbaca dari Criteria tampil 0; periksa sebelum menghitung.')
            st.markdown('Target toko dari Criteria')
            show_table(template['stores'])
            configs.append(config)

names = [c['name'] for c in configs]
reserved = {'store id', 'store name', 'am', 'rsm', 'total', '_key', 'reward'}
valid_names = all(names) and len({n.casefold() for n in names}) == len(names) and not any(n.casefold() in reserved for n in names)
if not valid_names:
    st.warning('Gunakan nama kategori unik dan tidak kosong. Nama kolom seperti AM, RSM, Total, dan Reward tidak dapat dipakai sebagai nama kategori.')
signature = hashlib.sha256(pickle.dumps((hashlib.sha256(revenue_bytes).hexdigest(),
                  [hashlib.sha256(b).hexdigest() for _, b in files], mapping, configs))).hexdigest()
confirmed = st.checkbox('Saya sudah memeriksa filter transaksi, target, dan aturan reward setiap kategori.', key='review_' + signature[:16])
if st.button('Hitung semua kategori', type='primary', disabled=not (valid_names and confirmed), width='stretch'):
    try:
        with st.spinner('Menghitung SO, ranking, dan bonus…'):
            results = [calculate(t, revenue, mapping, c) for t, c in zip(templates, configs)]
            period = f'{start:%d-%m-%Y} s.d. {end:%d-%m-%Y}'
            output = export_results(results, [b for _, b in files], period)
        st.session_state['calculated'] = {'signature': signature, 'results': results, 'output': output}
    except Exception as exc:
        st.error(f'Perhitungan belum dapat diselesaikan: {exc}')

saved = st.session_state.get('calculated')
if saved and saved['signature'] != signature:
    st.info('Input atau aturan berubah. Periksa pengaturan lalu hitung ulang untuk memperbarui hasil dan file download.')
ui.section_title('hasil', '03', 'Resume & hasil perhitungan', 'Bonus per penerima, detail setiap kategori, dan laporan Excel.')
if saved and saved['signature'] == signature:
    results = saved['results']
    notes = [f"{r['name']}: {n}" for r in results for n in r['notes']]
    if notes:
        st.warning('Hasil percobaan memiliki catatan. Bonus yang belum bisa dihitung ditampilkan kosong. Total yang tersedia bukan total final jika data penerima belum lengkap.')
        with st.expander('Lihat catatan perhitungan', expanded=True):
            for note in notes: st.write('• ' + note)
    cols = st.columns(3)
    for col, level, title in zip(cols, ['stores', 'am', 'rsm'], ['Bonus toko tersedia', 'Bonus AM tersedia', 'Bonus RSM tersedia']):
        table = resume(results, level)
        values = pd.to_numeric(table.Total, errors='coerce')
        col.metric(title, rupiah(values.sum()) if values.notna().any() else 'Belum tersedia')
    st.download_button('↓ Download hasil Excel', saved['output'], f'Result Sales Contest {start:%Y%m%d}-{end:%Y%m%d}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary', on_click='ignore')
    tabs = st.tabs(['Resume Toko', 'Resume AM', 'Resume RSM', 'Result per kategori'])
    for tab, level in zip(tabs[:3], ['stores', 'am', 'rsm']):
        with tab: show_table(resume(results, level))
    with tabs[3]:
        chosen = st.selectbox('Kategori hasil', names)
        result = next(r for r in results if r['name'] == chosen)
        st.caption(f"{result['rows']:,} baris transaksi memenuhi filter kategori dan periode.")
        st.markdown('**Toko**'); show_table(result['stores'])
        st.markdown('**Area Manager**'); show_table(result['am'])
        st.markdown('**Regional Sales Manager**'); show_table(result['rsm'])
else:
    st.info('Hasil akan tampil di sini setelah Anda memeriksa kriteria dan memilih “Hitung semua kategori”.', icon=':material/analytics:')
ui.footer()
