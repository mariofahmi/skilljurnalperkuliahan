#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_jurnal.py — Generator Dokumen JURNAL PERKULIAHAN Standar Resmi UNIROW Tuban (.docx & .md)

Tabel 6 Kolom: No | Hari | Tanggal | Jam | Makul | Pokok Bahasan
Desain presisi tinggi (KOP 3-kolom simetris, identitas tabular laser-aligned, 
distribusi kolom proporsional, border ganda resmi, dan blok tanda tangan rapi).

Pemakaian:
  py -3 generate_jurnal.py data.json output.docx [--md output.md]
  py -3 generate_jurnal.py rps.docx output.docx [--hari "Senin"] [--jam "08:00-09:40"] [--md output.md]
"""

import os
import sys
import json
import re
import argparse
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

FONT_NAME = "Times New Roman"

def set_cell_shading(cell, color_hex):
    """Menyetel warna latar sel tabel."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_padding(cell, top=60, bottom=60, left=70, right=70):
    """Menyetel margin dalam (padding) sel tabel dalam satuan dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcMar'):
            tcPr.remove(child)
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:left w:w="{left}" w:type="dxa"/>'
        f'  <w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="000000"):
    """Menyetel border garis sel tabel."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('tcBorders'):
            tcPr.remove(child)
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'  <w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

def clear_cell_borders(cell):
    """Menghapus semua garis batas sel tabel."""
    set_cell_borders(cell, top="none", bottom="none", left="none", right="none", sz="0", color="auto")

def prevent_row_split(row):
    """Mencegah baris tabel terpotong antar halaman."""
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:cantSplit'):
        cantSplit = parse_xml(f'<w:cantSplit {nsdecls("w")}/>')
        trPr.append(cantSplit)

def set_repeat_header(row):
    """Mengulang baris judul tabel di tiap halaman baru."""
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:tblHeader'):
        tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
        trPr.append(tblHeader)

def set_cell_valign(cell, val="center"):
    """Menyetel vertical alignment sel."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('vAlign'):
            tcPr.remove(child)
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{val}"/>')
    tcPr.append(vAlign)

def set_table_widths(table, col_widths_dxa):
    """Mengunci lebar kolom tabel secara presisi pada level tblGrid dan tcW dxa."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    tblPr = table._tbl.tblPr
    total_w = sum(col_widths_dxa)
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total_w}" w:type="dxa"/>')
    tblPr.append(tblW)
    
    tblGrid = table._tbl.tblGrid
    for c in list(tblGrid):
        tblGrid.remove(c)
    for w in col_widths_dxa:
        gridCol = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w}"/>')
        tblGrid.append(gridCol)

    for row in table.rows:
        prevent_row_split(row)
        for i, w in enumerate(col_widths_dxa):
            if i < len(row.cells):
                cell = row.cells[i]
                tcPr = cell._tc.get_or_add_tcPr()
                for child in list(tcPr):
                    if child.tag.endswith('tcW'):
                        tcPr.remove(child)
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w}" w:type="dxa"/>')
                tcPr.append(tcW)

def format_run(run, bold=False, italic=False, underline=False, size_pt=10, font_name=FONT_NAME):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rFonts.set(qn('w:cs'), font_name)

def add_p_run(paragraph, text, bold=False, italic=False, underline=False, size_pt=10, font_name=FONT_NAME):
    run = paragraph.add_run(text)
    format_run(run, bold=bold, italic=italic, underline=underline, size_pt=size_pt, font_name=font_name)
    return run

def clean_str(s):
    if not s: return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(s)).strip()

def parse_rps_to_jurnal_data(rps_path):
    """Mengekstrak data untuk jurnal perkuliahan dari dokumen RPS (.docx atau .md)."""
    data = {
        "prodi": "Pendidikan Pancasila dan Kewarganegaraan (PPKn)",
        "dosen": "Mario Fahmi Syahrial, M.Pd.",
        "kaprodi": "Mario Fahmi Syahrial, M.Pd.",
        "semester_kelas": "1 / ..........",
        "tahun_akademik": "20... / 20...",
        "jumlah_pertemuan": 16,
        "tanggal_penutup": ".........................",
        "baris": []
    }
    
    ext = os.path.splitext(rps_path)[1].lower()
    if ext == '.docx':
        doc = docx.Document(rps_path)
        base = os.path.splitext(os.path.basename(rps_path))[0]
        
        # Ekstrak semester dari nama file jika ada
        m_sem_file = re.match(r'^0?(\d+)_', base)
        if m_sem_file:
            s_num = int(m_sem_file.group(1))
            data["semester_kelas"] = f"{s_num} / .........."
            data["tahun_akademik"] = "20... / 20..."

        # Ekstrak identitas dari tabel 0 jika ada
        mk_name = ""
        if len(doc.tables) > 0:
            t0 = doc.tables[0]
            if len(t0.rows) > 3:
                r2 = [clean_str(c.text) for c in t0.rows[2].cells]
                r3 = [clean_str(c.text) for c in t0.rows[3].cells]
                if r3 and r3[0] and "MATA KULIAH" not in r3[0].upper():
                    mk_name = r3[0]
                for idx, col in enumerate(r2):
                    if idx < len(r3):
                        col_u = col.upper()
                        val = r3[idx]
                        if "SEMESTER" in col_u and val:
                            m = re.search(r'(\d+)', val)
                            if m:
                                sem_num = int(m.group(1))
                                data["semester_kelas"] = f"{sem_num} / .........."
                                data["tahun_akademik"] = "20... / 20..."
            
            # Cari dosen pengampu / pengembang RPS & Kaprodi
            for t in doc.tables[:3]:
                for r_idx, row in enumerate(t.rows):
                    for c_idx, cell in enumerate(row.cells):
                        txt = clean_str(cell.text).replace('\n', ' ')
                        # Cek Pengembang RPS
                        if any(k in txt for k in ["Pengembang RPS", "Dosen Pengembang RPS", "Dosen Pengembang"]):
                            if r_idx + 1 < len(t.rows):
                                val = clean_str(t.rows[r_idx + 1].cells[c_idx].text)
                                lines = [l.strip() for l in val.split("\n") if l.strip()]
                                for l in lines:
                                    if not any(k in l for k in ["NIDN", "Tanda Tangan", "Nama Dosen", "Pengembang RPS"]):
                                        data["dosen"] = l
                                        break
                            if not data.get("dosen") or data["dosen"] == "Mario Fahmi Syahrial, M.Pd.":
                                lines = [l.strip() for l in cell.text.split("\n") if l.strip()]
                                if len(lines) > 1:
                                    for l in lines[1:]:
                                        if not any(k in l for k in ["NIDN", "Tanda Tangan"]):
                                            data["dosen"] = clean_str(l)
                                            break
                        # Cek Kaprodi
                        if any(k in txt for k in ["Ketua PRODI", "Ketua Program Studi", "Kaprodi"]):
                            if r_idx + 1 < len(t.rows):
                                val = clean_str(t.rows[r_idx + 1].cells[c_idx].text)
                                lines = [l.strip() for l in val.split("\n") if l.strip()]
                                for l in lines:
                                    if not any(k in l for k in ["NIDN", "Tanda Tangan"]):
                                        data["kaprodi"] = l
                                        break
        
        # Fallback nama MK dari nama berkas jika kosong
        if not mk_name or mk_name == "Mata Kuliah" or len(mk_name) > 60:
            clean_base = re.sub(r'^\d+_\d+_RPS_', '', base)
            clean_base = re.sub(r'^\d+_RPS_', '', clean_base)
            clean_base = re.sub(r'^RPS_', '', clean_base)
            clean_base = re.sub(r'_(?:FINAL|UNIROW|OBE|COMPLETE)$', '', clean_base)
            mk_name = clean_base.replace('_', ' ').strip()
            
        # Ekstrak 16 minggu pembelajaran dari tabel 1
        if len(doc.tables) > 1:
            t1 = doc.tables[1]
            for row in t1.rows:
                cells = [clean_str(c.text) for c in row.cells]
                if not cells: continue
                col0 = cells[0]
                m_mg = re.match(r'^\s*(\d{1,2})\b', col0)
                is_uts = 'UTS' in col0.upper() or 'TENGAH SEMESTER' in col0.upper()
                is_uas = 'UAS' in col0.upper() or 'AKHIR SEMESTER' in col0.upper()
                
                if m_mg or is_uts or is_uas:
                    mg_no = int(m_mg.group(1)) if m_mg else (8 if is_uts else 16)
                    pokok = ""
                    if is_uts:
                        if len(cells) > 6 and len(cells[6]) > 10:
                            pokok = cells[6]
                        else:
                            pokok = "Ujian Tengah Semester (UTS)"
                    elif is_uas:
                        if len(cells) > 6 and len(cells[6]) > 10:
                            pokok = cells[6]
                        else:
                            pokok = "Ujian Akhir Semester (UAS)"
                    else:
                        for candidate_idx in [6, 5, 4, 1]:
                            if candidate_idx < len(cells) and cells[candidate_idx]:
                                pokok = cells[candidate_idx]
                                break
                                
                    pokok_clean = re.sub(r'\s+', ' ', pokok).strip()
                    if len(pokok_clean) > 220:
                        pokok_clean = pokok_clean[:217] + "..."
                    
                    data["baris"].append({
                        "no": mg_no,
                        "hari": "",
                        "tanggal": "",
                        "jam": "",
                        "makul": mk_name,
                        "pokok_bahasan": pokok_clean
                    })
    
    seen_no = set()
    unique_baris = []
    for b in sorted(data["baris"], key=lambda x: x.get("no", 0)):
        no = b.get("no", 0)
        if no not in seen_no and 1 <= no <= 16:
            seen_no.add(no)
            unique_baris.append(b)
    data["baris"] = unique_baris
    return data

def build_jurnal_docx(data, output_path, logo_path=None):
    """Menghasilkan file DOCX Jurnal Perkuliahan UNIROW presisi tinggi, rapi, dan estetis."""
    doc = Document()
    
    # 1. Konfigurasi Halaman (A4 Portrait, Margin Kompak Proporsional)
    # A4 = 11906 x 16838 dxa. Margin: Top 850 dxa (0.59 in), Bottom 850 dxa (0.59 in), Left 1080 dxa (0.75 in), Right 1080 dxa (0.75 in)
    # Lebar Konten = 11906 - 2160 = 9746 dxa.
    section = doc.sections[0]
    section.page_width = Inches(8.267)    # 11906 dxa
    section.page_height = Inches(11.693)  # 16838 dxa
    section.top_margin = Inches(0.59)     # ~850 dxa
    section.bottom_margin = Inches(0.59)  # ~850 dxa
    section.left_margin = Inches(0.75)    # 1080 dxa
    section.right_margin = Inches(0.75)   # 1080 dxa
    
    TOTAL_CONTENT_WIDTH = 9746
    
    # Path Logo
    if not logo_path or not os.path.exists(logo_path):
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        candidate = os.path.join(base_dir, "assets", "logo_unirow.png")
        if os.path.exists(candidate):
            logo_path = candidate
        else:
            logo_path = None

    # 2. Kop Surat UNIROW 3-Kolom Simetris (Logo Kiri, Teks Tengah Dead-Center, Spacer Kanan)
    # Lebar: Logo 1200 dxa, Teks 7346 dxa, Spacer 1200 dxa = 9746 dxa
    kop_table = doc.add_table(rows=1, cols=3)
    kop_widths = [1200, 7346, 1200]
    set_table_widths(kop_table, kop_widths)
    
    kop_row = kop_table.rows[0]
    cell_logo = kop_row.cells[0]
    cell_text = kop_row.cells[1]
    cell_spacer = kop_row.cells[2]
    
    clear_cell_borders(cell_logo)
    clear_cell_borders(cell_text)
    clear_cell_borders(cell_spacer)
    
    set_cell_padding(cell_logo, top=0, bottom=0, left=0, right=40)
    set_cell_padding(cell_text, top=0, bottom=0, left=20, right=20)
    set_cell_padding(cell_spacer, top=0, bottom=0, left=40, right=0)
    
    set_cell_valign(cell_logo, "center")
    set_cell_valign(cell_text, "center")
    set_cell_valign(cell_spacer, "center")
    
    # Logo UNIROW (Kiri)
    p_logo = cell_logo.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_logo.paragraph_format.space_before = Pt(0)
    p_logo.paragraph_format.space_after = Pt(0)
    if logo_path and os.path.exists(logo_path):
        p_logo.add_run().add_picture(logo_path, width=Inches(0.82))
    
    # Teks Kop (Pusat Simetris)
    p_head = cell_text.paragraphs[0]
    p_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_head.paragraph_format.space_before = Pt(0)
    p_head.paragraph_format.space_after = Pt(0)
    p_head.paragraph_format.line_spacing = 1.05
    
    add_p_run(p_head, "JURNAL PERKULIAHAN\n", bold=True, size_pt=15)
    th_akademik = data.get('tahun_akademik') or "20... / 20..."
    add_p_run(p_head, f"TAHUN AKADEMIK {th_akademik}\n", bold=True, size_pt=11)
    add_p_run(p_head, "UNIVERSITAS PGRI RONGGOLAWE TUBAN\n", bold=True, size_pt=14)
    add_p_run(p_head, "Jl. Manunggal 61 Tuban, Telp. (0356) 322233, Fax (0356) 331578\n", bold=False, size_pt=9)
    add_p_run(p_head, "Website: www.unirow.ac.id, Email: prospective@unirow.ac.id", bold=False, size_pt=9)
    
    # Garis Pembatas KOP (Garis Ganda Khas Surat Resmi)
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(2)
    p_div.paragraph_format.space_after = Pt(6)
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="3" w:color="000000"/></w:pBdr>')
    p_div._p.get_or_add_pPr().append(pBdr)
    
    # 3. Blok Identitas Mata Kuliah (Tabular Laser-Aligned tanpa Border)
    # Lebar: Label 1900 dxa, Colon 200 dxa, Value 7646 dxa = 9746 dxa
    identitas = [
        ("Fakultas", "Keguruan dan Ilmu Pendidikan (FKIP)"),
        ("Program Studi", data.get("prodi", "")),
        ("Dosen Pengampu", data.get("dosen", "")),
        ("Semester/Kelas", data.get("semester_kelas", "")),
    ]
    
    id_table = doc.add_table(rows=len(identitas), cols=3)
    set_table_widths(id_table, [1900, 200, 7646])
    
    for idx, (label, val) in enumerate(identitas):
        row = id_table.rows[idx]
        prevent_row_split(row)
        
        c_lbl, c_col, c_val = row.cells[0], row.cells[1], row.cells[2]
        clear_cell_borders(c_lbl)
        clear_cell_borders(c_col)
        clear_cell_borders(c_val)
        
        set_cell_padding(c_lbl, top=15, bottom=15, left=0, right=0)
        set_cell_padding(c_col, top=15, bottom=15, left=0, right=0)
        set_cell_padding(c_val, top=15, bottom=15, left=0, right=0)
        
        set_cell_valign(c_lbl, "center")
        set_cell_valign(c_col, "center")
        set_cell_valign(c_val, "center")
        
        p0 = c_lbl.paragraphs[0]
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        add_p_run(p0, label, bold=True, size_pt=10)
        
        p1 = c_col.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        add_p_run(p1, ":", bold=True, size_pt=10)
        
        p2 = c_val.paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        add_p_run(p2, f" {val}", bold=False, size_pt=10)
        
    p_sp_table = doc.add_paragraph()
    p_sp_table.paragraph_format.space_before = Pt(0)
    p_sp_table.paragraph_format.space_after = Pt(4)
    
    # 4. Tabel Jurnal Perkuliahan (6 Kolom: No, Hari, Tanggal, Jam, Makul, Pokok Bahasan)
    # Distribusi Lebar: No 450, Hari 950, Tanggal 1250, Jam 1050, Makul 1950, Pokok Bahasan 4096 = 9746 dxa
    COL_WIDTHS = [450, 950, 1250, 1050, 1950, 4096]
    HEADERS = ["No", "Hari", "Tanggal", "Jam", "Makul", "Pokok Bahasan"]
    
    jumlah_pertemuan = data.get("jumlah_pertemuan", 16)
    baris_list = data.get("baris", [])
    
    table = doc.add_table(rows=1 + jumlah_pertemuan, cols=6)
    set_table_widths(table, COL_WIDTHS)
    
    # Header Row
    header_row = table.rows[0]
    set_repeat_header(header_row)
    prevent_row_split(header_row)
    
    for idx, name in enumerate(HEADERS):
        cell = header_row.cells[idx]
        set_cell_shading(cell, "E2E2E2")
        set_cell_padding(cell, top=60, bottom=60, left=50, right=50)
        set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="000000")
        set_cell_valign(cell, "center")
        
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_p_run(p, name, bold=True, size_pt=9.5)
        
    # Data Rows
    for i in range(1, jumlah_pertemuan + 1):
        row = table.rows[i]
        prevent_row_split(row)
        
        row_data = next((b for b in baris_list if int(b.get("no", 0)) == i), {})
        
        cell_values = [
            (str(row_data.get("no", i)), WD_ALIGN_PARAGRAPH.CENTER),
            (str(row_data.get("hari", "")), WD_ALIGN_PARAGRAPH.CENTER),
            (str(row_data.get("tanggal", "")), WD_ALIGN_PARAGRAPH.CENTER),
            (str(row_data.get("jam", "")), WD_ALIGN_PARAGRAPH.CENTER),
            (str(row_data.get("makul", "")), WD_ALIGN_PARAGRAPH.LEFT),
            (str(row_data.get("pokok_bahasan", "")), WD_ALIGN_PARAGRAPH.LEFT),
        ]
        
        for idx, (val, alignment) in enumerate(cell_values):
            cell = row.cells[idx]
            set_cell_padding(cell, top=35, bottom=35, left=50, right=50)
            set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", color="000000")
            set_cell_valign(cell, "center")
            
            p = cell.paragraphs[0]
            p.alignment = alignment
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            add_p_run(p, val, bold=False, size_pt=8.5)
            
    # 5. Blok Tanda Tangan Penutup (Simetris & Presisi di Bawah Kolom Tabel)
    p_sp_sig = doc.add_paragraph()
    p_sp_sig.paragraph_format.space_before = Pt(8)
    p_sp_sig.paragraph_format.space_after = Pt(0)
    p_sp_sig.paragraph_format.line_spacing = Pt(1)
    r_sp = p_sp_sig.add_run()
    r_sp.font.size = Pt(1)
    
    sig_table = doc.add_table(rows=1, cols=2)
    sig_widths = [4873, 4873]
    set_table_widths(sig_table, sig_widths)
    
    sig_row = sig_table.rows[0]
    prevent_row_split(sig_row)
    
    cell_kiri = sig_row.cells[0]
    cell_kanan = sig_row.cells[1]
    
    clear_cell_borders(cell_kiri)
    clear_cell_borders(cell_kanan)
    set_cell_padding(cell_kiri, top=0, bottom=0, left=0, right=0)
    set_cell_padding(cell_kanan, top=0, bottom=0, left=0, right=0)
    set_cell_valign(cell_kiri, "top")
    set_cell_valign(cell_kanan, "top")
    
    # Kolom Kiri: Ketua Program Studi
    p_kiri1 = cell_kiri.paragraphs[0]
    p_kiri1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kiri1.paragraph_format.space_before = Pt(0)
    p_kiri1.paragraph_format.space_after = Pt(0)
    p_kiri1.paragraph_format.line_spacing = 1.15
    add_p_run(p_kiri1, "Mengetahui,\nKetua Program Studi", bold=False, size_pt=10)
    
    p_kiri_sp = cell_kiri.add_paragraph()
    p_kiri_sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kiri_sp.paragraph_format.space_before = Pt(45)
    p_kiri_sp.paragraph_format.space_after = Pt(0)
    kaprodi_name = data.get("kaprodi", "Mario Fahmi Syahrial, M.Pd.")
    add_p_run(p_kiri_sp, kaprodi_name, bold=True, underline=True, size_pt=10)
    
    # Kolom Kanan: Dosen Pengampu (Tanggal presisi terpusat di atas jabatan)
    p_kanan1 = cell_kanan.paragraphs[0]
    p_kanan1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kanan1.paragraph_format.space_before = Pt(0)
    p_kanan1.paragraph_format.space_after = Pt(0)
    p_kanan1.paragraph_format.line_spacing = 1.15
    tgl_penutup = data.get("tanggal_penutup") or "........................."
    add_p_run(p_kanan1, f"Tuban, {tgl_penutup}\nDosen Pengampu,", bold=False, size_pt=10)
    
    p_kanan_sp = cell_kanan.add_paragraph()
    p_kanan_sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_kanan_sp.paragraph_format.space_before = Pt(45)
    p_kanan_sp.paragraph_format.space_after = Pt(0)
    dosen_name = data.get("dosen", "........................................")
    add_p_run(p_kanan_sp, dosen_name, bold=True, underline=True, size_pt=10)
    
    # Simpan file DOCX
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    print(f"[OK] Berhasil menghasilkan Jurnal Perkuliahan UNIROW: {output_path}")

def build_jurnal_md(data, output_path):
    """Menghasilkan file Markdown Jurnal Perkuliahan UNIROW standar GFM."""
    lines = []
    lines.append("# JURNAL PERKULIAHAN")
    lines.append(f"## TAHUN AKADEMIK {data.get('tahun_akademik') or '20... / 20...'}")
    lines.append("### UNIVERSITAS PGRI RONGGOLAWE TUBAN")
    lines.append("Jl. Manunggal 61 Tuban, Telp. (0356) 322233, Fax (0356) 331578  ")
    lines.append("Website: www.unirow.ac.id, Email: prospective@unirow.ac.id\n")
    lines.append("---\n")
    lines.append(f"- **Fakultas** : Keguruan dan Ilmu Pendidikan (FKIP)")
    lines.append(f"- **Program Studi** : {data.get('prodi', '')}")
    lines.append(f"- **Dosen Pengampu** : {data.get('dosen', '')}")
    lines.append(f"- **Semester/Kelas** : {data.get('semester_kelas', '')}\n")
    lines.append("---\n")
    lines.append("### Catatan Pokok Bahasan Mingguan\n")
    lines.append("| No | Hari | Tanggal | Jam | Makul | Pokok Bahasan |")
    lines.append("|:---:|:---:|:---:|:---:|:---|:---|")
    
    jumlah_pertemuan = data.get("jumlah_pertemuan", 16)
    baris_list = data.get("baris", [])
    
    for i in range(1, jumlah_pertemuan + 1):
        row_data = next((b for b in baris_list if int(b.get("no", 0)) == i), {})
        no = str(row_data.get("no", i))
        hari = str(row_data.get("hari", ""))
        tanggal = str(row_data.get("tanggal", ""))
        jam = str(row_data.get("jam", ""))
        makul = str(row_data.get("makul", "")).replace('|', '\\|')
        pokok = str(row_data.get("pokok_bahasan", "")).replace('|', '\\|')
        lines.append(f"| {no} | {hari} | {tanggal} | {jam} | {makul} | {pokok} |")
        
    lines.append("\n---\n")
    tgl_penutup = data.get('tanggal_penutup') or '.........................'
    lines.append(f"| Mengetahui,<br>Ketua Program Studi | Tuban, {tgl_penutup}<br>Dosen Pengampu, |")
    lines.append("|:---:|:---:|")
    kaprodi = data.get("kaprodi", "Mario Fahmi Syahrial, M.Pd.")
    dosen = data.get("dosen", "........................................")
    lines.append(f"| <br><br><br><u>**{kaprodi}**</u> | <br><br><br><u>**{dosen}**</u> |")
    
    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    print(f"[OK] Berhasil menghasilkan Markdown: {output_path}")

def main():
    parser = argparse.ArgumentParser(description="Generator Jurnal Perkuliahan UNIROW")
    parser.add_argument("input_path", help="Path ke data.json ATAU berkas RPS (.docx / .md)")
    parser.add_argument("output_path", help="Path keluaran file .docx")
    parser.add_argument("--md", help="Path opsional keluaran file .md", default=None)
    parser.add_argument("--logo", help="Path kustom ke file logo_unirow.png", default=None)
    parser.add_argument("--hari", help="Nama hari perkuliahan (mis. 'Senin')", default=None)
    parser.add_argument("--jam", help="Jam perkuliahan (mis. '08:00-09:40')", default=None)
    parser.add_argument("--prodi", help="Override nama Program Studi", default=None)
    parser.add_argument("--dosen", help="Override nama Dosen Pengampu", default=None)
    parser.add_argument("--semester_kelas", help="Override semester dan kelas", default=None)
    parser.add_argument("--tahun_akademik", help="Override tahun akademik", default=None)
    
    args = parser.parse_args()
    
    in_ext = os.path.splitext(args.input_path)[1].lower()
    if in_ext == '.json':
        with open(args.input_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    elif in_ext in ['.docx', '.md']:
        print(f"[*] Mengekstrak data dari RPS: {args.input_path}")
        data = parse_rps_to_jurnal_data(args.input_path)
    else:
        try:
            with open(args.input_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            raise ValueError(f"Format berkas input tidak dikenali: {args.input_path}")

    if args.prodi: data["prodi"] = args.prodi
    if args.dosen: data["dosen"] = args.dosen
    if args.semester_kelas: data["semester_kelas"] = args.semester_kelas
    if args.tahun_akademik: data["tahun_akademik"] = args.tahun_akademik
    if args.hari or args.jam:
        for b in data.get("baris", []):
            if args.hari and not b.get("hari"): b["hari"] = args.hari
            if args.jam and not b.get("jam"): b["jam"] = args.jam

    build_jurnal_docx(data, args.output_path, logo_path=args.logo)
    if args.md:
        build_jurnal_md(data, args.md)

if __name__ == "__main__":
    main()
