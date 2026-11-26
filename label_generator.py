"""
label_generator.py
Modul untuk membaca data Excel daftar barang, memvalidasi data,
dan membuat dokumen Word (.docx) label barang berdasarkan template 'Label New Bold.docx'.
"""

import copy
import io
import math
import os
import re
from typing import Dict, List, Optional, Tuple, Union

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
import openpyxl
import pandas as pd


def normalize_header(header: str) -> str:
    """Membersihkan dan menormalisasi string header untuk pencocokan."""
    if not header:
        return ""
    return re.sub(r"\s+", " ", str(header).strip().lower())


def find_header_row(ws) -> Tuple[Optional[int], Dict[str, int]]:
    """
    Mencari baris header pada worksheet openpyxl secara dinamis.
    Mencari kolom: 'No', 'List Item' (atau Item/Nama Barang), 'Spec' (atau Spesifikasi).
    
    Mengembalikan:
        (header_row_index, mapping_column_name_to_col_index)
    """
    for r in range(1, min(ws.max_row + 1, 25)):
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        norm_vals = [normalize_header(v) for v in row_vals]

        no_col = None
        item_col = None
        spec_col = None
        uom_col = None

        for idx, val in enumerate(norm_vals, start=1):
            if not val:
                continue
            if val in ["no", "no.", "nomor"]:
                no_col = idx
            elif any(k in val for k in ["list item", "nama barang", "item barang", "nama item"]):
                item_col = idx
            elif val == "item":
                item_col = idx
            elif any(k in val for k in ["spec", "spesifikasi"]):
                spec_col = idx
            elif any(k in val for k in ["uom", "satuan"]):
                uom_col = idx

        if no_col is not None and item_col is not None and spec_col is not None:
            return r, {"no": no_col, "item": item_col, "spec": spec_col, "uom": uom_col}

    return None, {}


def read_excel_data(
    file_source: Union[str, io.BytesIO],
    skip_hidden_rows: bool = True
) -> Tuple[Optional[pd.DataFrame], Optional[str]]:
    """
    Membaca file Excel yang berisi daftar barang.
    
    Ketentuan:
    - Mengambil kolom No, List Item, dan Spec.
    - Kolom UOM tidak dimasukkan ke dalam hasil akhir Word.
    - Membaca nama kolom berdasarkan header, bukan posisi kolom.
    - Mengabaikan baris tersembunyi (hidden rows) di Excel jika skip_hidden_rows=True.
    - Mengabaikan baris dengan No non-angka (seperti 'A' pada PERLENGKAPAN KERJA).
    - Menampilkan data apa adanya sesuai Excel jika ada cell yang kosong.
    
    Returns:
        (DataFrame dengan kolom ['No', 'List Item', 'Spec'], error_message)
    """
    try:
        wb = openpyxl.load_workbook(file_source, data_only=True)
    except Exception as e:
        return None, f"Gagal membaca file Excel. Pastikan file berformat .xlsx yang valid. Detail: {str(e)}"

    sheet_names = wb.sheetnames
    if not sheet_names:
        return None, "File Excel tidak memiliki lembar kerja (worksheet)."

    ws = wb.active

    header_row, col_map = find_header_row(ws)
    if header_row is None:
        return None, (
            "Header kolom tidak ditemukan. Pastikan file Excel memiliki kolom "
            "'No', 'List Item', dan 'Spec'."
        )

    no_col = col_map["no"]
    item_col = col_map["item"]
    spec_col = col_map["spec"]
    uom_col = col_map.get("uom")

    records = []
    for r in range(header_row + 1, ws.max_row + 1):
        # Abaikan baris tersembunyi di Excel (misal baris arsip/lama)
        if skip_hidden_rows:
            rd = ws.row_dimensions.get(r)
            if rd and rd.hidden:
                continue

        no_val = ws.cell(r, no_col).value
        item_val = ws.cell(r, item_col).value
        spec_val = ws.cell(r, spec_col).value
        uom_val = ws.cell(r, uom_col).value if uom_col else None

        # Abaikan jika baris benar-benar kosong
        if no_val is None and item_val is None and spec_val is None and uom_val is None:
            continue

        # Cek jika baris adalah header berulang
        if str(no_val).strip().lower() in ["no", "no.", "nomor"]:
            continue

        # Validasi kolom No (harus berupa nomor/angka)
        if no_val is None:
            continue

        is_numeric_no = False
        parsed_no = None

        if isinstance(no_val, (int, float)):
            try:
                parsed_no = int(no_val)
                is_numeric_no = True
            except (ValueError, OverflowError):
                is_numeric_no = False
        elif isinstance(no_val, str):
            clean_str = no_val.strip()
            if clean_str.isdigit():
                parsed_no = int(clean_str)
                is_numeric_no = True
            else:
                is_numeric_no = False

        if not is_numeric_no:
            # Baris dengan no non-angka (misal 'A' pada 'PERLENGKAPAN KERJA') dilewati
            continue

        item_str = str(item_val).strip() if item_val is not None else ""
        spec_str = str(spec_val).strip() if spec_val is not None else ""
        if spec_str.lower() in ["none", "nan"]:
            spec_str = ""

        # Simpan persis sesuai data di Excel (termasuk jika nama barang atau spec kosong)
        records.append({
            "No": parsed_no,
            "List Item": item_str,
            "Spec": spec_str
        })

    if not records:
        return None, "Tidak ada data barang yang valid ditemukan dalam file Excel."

    df = pd.DataFrame(records)
    return df, None


def filter_by_range(df: pd.DataFrame, start_no: int, end_no: int) -> Tuple[pd.DataFrame, List[int], Optional[str]]:
    """
    Menyaring data barang berdasarkan rentang nomor [start_no, end_no].
    
    Returns:
        (filtered_df, missing_nos, error_message)
    """
    if start_no > end_no:
        return pd.DataFrame(), [], f"Nomor awal ({start_no}) tidak boleh lebih besar dari nomor akhir ({end_no})."

    # Ambil barang dengan nomor di antara start_no dan end_no
    mask = (df["No"] >= start_no) & (df["No"] <= end_no)
    filtered_df = df[mask].copy()

    # Cek nomor yang ada di rentang tapi tidak ada di Excel
    existing_nos = set(filtered_df["No"].unique())
    requested_range = range(start_no, end_no + 1)
    missing_nos = [n for n in requested_range if n not in existing_nos]

    if filtered_df.empty:
        return filtered_df, missing_nos, f"Tidak ada data barang yang ditemukan untuk rentang nomor {start_no} sampai {end_no}."

    return filtered_df, missing_nos, None


def filter_by_manual_list(df: pd.DataFrame, selected_nos: List[int]) -> Tuple[pd.DataFrame, List[int], Optional[str]]:
    """
    Menyaring data barang berdasarkan daftar nomor manual yang dipilih user.
    Mempertahankan urutan dari Excel sesuai instruksi.
    
    Returns:
        (filtered_df, missing_nos, error_message)
    """
    if not selected_nos:
        return pd.DataFrame(), [], "Silakan masukkan atau pilih setidaknya satu nomor barang."

    existing_nos_in_excel = set(df["No"].unique())
    missing_nos = [n for n in selected_nos if n not in existing_nos_in_excel]

    filtered_df = df[df["No"].isin(selected_nos)].copy()

    if filtered_df.empty:
        return filtered_df, missing_nos, "Nomor-nomor yang dipilih tidak ditemukan dalam data Excel."

    return filtered_df, missing_nos, None


def set_cell_properties(cell, width_dxa: int = 1440):
    """
    Memastikan properti cell sesuai referensi Label New Bold.docx:
    - tcW: width_dxa (1440 dxa = 1 inci / 72 pt)
    - textDirection: tbRl (vertikal dari atas ke bawah, kanan ke kiri)
    - vAlign: center
    """
    tcPr = cell._tc.get_or_add_tcPr()

    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width_dxa))
    tcW.set(qn("w:type"), "dxa")

    textDir = tcPr.find(qn("w:textDirection"))
    if textDir is None:
        textDir = OxmlElement("w:textDirection")
        tcPr.append(textDir)
    textDir.set(qn("w:val"), "tbRl")

    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), "center")


def populate_label_cell(cell, item_name: str, spec: str = ""):
    """
    Mengisi sebuah sel label dengan nama barang dan spesifikasi sesuai format Label New Bold.docx:
    - Jika nama barang ada & spec ada: P0=Nama Barang (22pt Bold), P1=Spec (20pt Bold)
    - Jika nama barang ada & spec kosong: P0=Nama Barang (22pt Bold), tanpa P1
    - Jika nama barang kosong & spec ada: P0=Spec (22pt Bold), tanpa P1
    - Jika keduanya kosong: sel dikosongkan
    """
    set_cell_properties(cell)

    # Hapus semua paragraf lama di dalam cell
    for p in list(cell.paragraphs):
        p_elem = p._p
        p_elem.getparent().remove(p_elem)

    clean_item = item_name.strip() if item_name else ""
    clean_spec = spec.strip() if spec else ""

    # Jika keduanya kosong
    if not clean_item and not clean_spec:
        cell.add_paragraph()
        return

    # Tentukan teks utama (P0) dan teks kedua (P1)
    if clean_item and clean_spec:
        # P0: List Item (Nama Barang)
        p_item = cell.add_paragraph()
        p_item.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pr = p_item._p.get_or_add_pPr()

        ind = p_pr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            p_pr.append(ind)
        ind.set(qn("w:left"), "113")
        ind.set(qn("w:right"), "113")

        run_item = p_item.add_run(clean_item)
        run_item.font.name = "Nyala"
        run_item.font.size = Pt(22)
        run_item.bold = True
        run_item.font.color.rgb = RGBColor(0, 0, 0)

        # P1: Spec (Spesifikasi)
        p_spec = cell.add_paragraph()
        p_spec.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_spec_pr = p_spec._p.get_or_add_pPr()

        ind_spec = p_spec_pr.find(qn("w:ind"))
        if ind_spec is None:
            ind_spec = OxmlElement("w:ind")
            p_spec_pr.append(ind_spec)
        ind_spec.set(qn("w:left"), "36")
        ind_spec.set(qn("w:right"), "36")

        run_spec = p_spec.add_run(clean_spec)
        run_spec.font.name = "Nyala"
        run_spec.font.size = Pt(20)
        run_spec.bold = True
        run_spec.font.color.rgb = RGBColor(0, 0, 0)

    else:
        # Hanya salah satu yang ada (item saja atau spec saja)
        single_text = clean_item if clean_item else clean_spec
        p = cell.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_pr = p._p.get_or_add_pPr()

        ind = p_pr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            p_pr.append(ind)
        ind.set(qn("w:left"), "113")
        ind.set(qn("w:right"), "113")

        run = p.add_run(single_text)
        run.font.name = "Nyala"
        run.font.size = Pt(22)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)


def clear_label_cell(cell):
    """Mengosongkan sel label."""
    set_cell_properties(cell)
    for p in list(cell.paragraphs):
        p_elem = p._p
        p_elem.getparent().remove(p_elem)
    cell.add_paragraph()


def generate_word_document(
    items: List[Dict[str, str]],
    template_path: str = "templates/Label New Bold.docx"
) -> io.BytesIO:
    """
    Membuat dokumen Word (.docx) berdasarkan daftar barang dan template referensi.
    Setiap halaman template menampung 30 label (3 baris label x 10 kolom).
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"File template tidak ditemukan di: {template_path}")

    doc = docx.Document(template_path)
    if not doc.tables:
        raise ValueError("File template Word tidak memiliki tabel.")

    base_table = doc.tables[0]
    base_tbl_xml = copy.deepcopy(base_table._tbl)

    LABEL_ROWS = [0, 2, 4]
    COLS_PER_ROW = 10
    LABELS_PER_PAGE = len(LABEL_ROWS) * COLS_PER_ROW  # 30

    total_items = len(items)
    num_pages = max(1, math.ceil(total_items / LABELS_PER_PAGE))

    body = doc._body._body
    curr_trailing_p = None
    for el in body:
        if el.tag.split("}")[-1] == "p":
            curr_trailing_p = el

    for page_idx in range(1, num_pages):
        if curr_trailing_p is not None:
            pPr = curr_trailing_p.get_or_add_pPr()
            rPr = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz")
            sz.set(qn("w:val"), "2")
            rPr.append(sz)
            vanish = OxmlElement("w:vanish")
            rPr.append(vanish)
            pPr.append(rPr)

            r = OxmlElement("w:r")
            r.append(copy.deepcopy(rPr))
            br = OxmlElement("w:br")
            br.set(qn("w:type"), "page")
            r.append(br)
            curr_trailing_p.append(r)

            new_tbl = copy.deepcopy(base_tbl_xml)
            curr_trailing_p.addnext(new_tbl)

            new_trailing_p = OxmlElement("w:p")
            new_pPr = OxmlElement("w:pPr")
            new_rPr = copy.deepcopy(rPr)
            new_pPr.append(new_rPr)
            new_trailing_p.append(new_pPr)

            new_tbl.addnext(new_trailing_p)
            curr_trailing_p = new_trailing_p
        else:
            p = doc.add_paragraph()
            r = p.add_run()
            r.add_break(docx.enum.text.WD_BREAK.PAGE)
            new_tbl = copy.deepcopy(base_tbl_xml)
            p._p.addnext(new_tbl)

    all_tables = doc.tables

    item_idx = 0
    for page_idx, tbl in enumerate(all_tables):
        for r_idx in LABEL_ROWS:
            row = tbl.rows[r_idx]
            for c_idx in range(COLS_PER_ROW):
                cell = row.cells[c_idx]
                if item_idx < total_items:
                    item_data = items[item_idx]
                    item_name = str(item_data.get("List Item", ""))
                    spec = str(item_data.get("Spec", ""))
                    populate_label_cell(cell, item_name, spec)
                    item_idx += 1
                else:
                    clear_label_cell(cell)

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io
