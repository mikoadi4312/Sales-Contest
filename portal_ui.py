"""EraBlue-inspired presentation shared by the contest workspace."""
import base64
from html import escape
from pathlib import Path

import streamlit as st

ROOT = Path(__file__).resolve().parent


def template_download(label, content, filename):
    """Small templates use native downloads, independent of Streamlit media/session URLs."""
    payload = base64.b64encode(content).decode('ascii')
    mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    st.markdown(
        f'<a class="template-download" href="data:{mime};base64,{payload}" '
        f'download="{escape(filename, quote=True)}" title="Download {escape(filename, quote=True)}">'
        f'<span aria-hidden="true">↓</span> {escape(label)}</a>',
        unsafe_allow_html=True,
    )


def icon(name, color='#0066b3'):
    variant = 'white' if color == 'white' else 'blue'
    data = base64.b64encode((ROOT / 'assets' / f'{name}-{variant}.png').read_bytes()).decode()
    return f'<img class="ui-icon" src="data:image/png;base64,{data}" alt="" width="24" height="24">'



def logo():
    data = base64.b64encode((ROOT / 'erablue-logo.png').read_bytes()).decode()
    return f'<img src="data:image/png;base64,{data}" alt="EraBlue Electronics" class="era-logo">'


def apply_theme():
    st.html('<style>' + (ROOT / 'portal.css').read_text(encoding='utf-8') + '</style>')


def sidebar_brand():
    st.html(f'<div class="sidebar-brand">{logo()}<span>SALES CONTEST</span></div>')
    st.html(f'<div class="workspace-label">{icon("grid")} <span>Contest workspace</span></div>')


def section_title(anchor, number, title, subtitle):
    st.html(f'<div id="{anchor}" class="section-heading"><span class="section-number">{number}</span><div><h2>{title}</h2><p>{subtitle}</p></div></div>')


def upload_label(name, text):
    st.html(f'<div class="upload-label">{icon("file")}<div><strong>{name}</strong><span>{text}</span></div></div>')


def pending_sections():
    st.html('''<div class="pending-grid"><section id="kriteria" class="pending-card"><span class="pending-step">02</span><div><h3>Kriteria siap setelah data diunggah</h3><p>Upload template kategori dan Revenue untuk mengatur periode, filter, serta reward.</p></div></section><section id="hasil" class="pending-card"><span class="pending-step">03</span><div><h3>Hasil perhitungan Anda</h3><p>Resume toko, AM, dan RSM akan tersedia setelah Anda menjalankan perhitungan.</p></div></section></div>''')


def footer():
    st.html('<footer class="portal-footer"><span><b class="era-red">era</b><b class="era-blue">blue</b> &nbsp; Sales Contest Workspace</span><span>Kelola data. Hitung pencapaian. Apresiasi hasil.</span></footer>')
