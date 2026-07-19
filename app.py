# -*- coding: utf-8 -*-
"""
Teklif → AX Dönüştürme API Servisi
Lovable uygulamasından gelen teklif Excel dosyasını okur,
orijinal formatta AX Excel dosyası üretir ve geri döner.
"""

import io
import re
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
CORS(app)  # Lovable'dan gelen isteklere izin ver

# ─── Stiller ─────────────────────────────────────────────
RED = "FF0000"
BLUE = "0000FF"
BLACK = "000000"

thin = Side(style="thin", color=BLACK)
medium_blue = Side(style="medium", color=BLUE)


def cell_border(top=False, bottom=False, left=False, right=False):
    """İç ince kenarlık + dış kalın mavi kenarlık kombinasyonu"""
    return Border(
        top=medium_blue if top else thin,
        bottom=medium_blue if bottom else thin,
        left=medium_blue if left else thin,
        right=medium_blue if right else thin,
    )


# ─── Teklif Okuma ────────────────────────────────────────
def parse_teklif(file_bytes):
    """Teklif Excel'ini oku, üst bilgileri ve kalemleri ayıkla"""
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.active

    info = {"magaza_adi": "", "magaza_kodu": "", "srv_no": ""}
    kalemler = []

    for row in ws.iter_rows(values_only=False):
        cells = [c.value for c in row]
        # Üst bilgileri bul: etiket hücresinin sağındaki dolu hücre değerdir
        for i, v in enumerate(cells):
            if v is None:
                continue
            text = str(v).strip()
            low = text.lower().replace(":", "").replace("ı", "i")
            if low == "magaza adi" or low == "mağaza adi" or "mağaza adı" in text.lower():
                info["magaza_adi"] = _next_value(cells, i)
            elif "mağaza kodu" in text.lower() or "magaza kodu" in low:
                info["magaza_kodu"] = _next_value(cells, i)
            elif low.startswith("srv no"):
                info["srv_no"] = _next_value(cells, i)

        # Kalem satırı mı? ET kodu içeren hücre ara
        et_idx = None
        for i, v in enumerate(cells):
            if v is not None and re.match(r"^ET\d+", str(v).strip()):
                et_idx = i
                break
        if et_idx is None:
            continue

        try:
            et_kod = str(cells[et_idx]).strip()
            aciklama = str(cells[et_idx + 1]).strip() if cells[et_idx + 1] else ""
            birim = str(cells[et_idx + 2]).strip() if cells[et_idx + 2] else ""
            miktar = _num(cells[et_idx + 3])
            birim_fiyat = _num(cells[et_idx + 4])
            tutar = _num(cells[et_idx + 5])
        except (IndexError, TypeError):
            continue

        # Sadece tutari 0'dan büyük kalemler AX'e girer
        if tutar and tutar > 0:
            kalemler.append({
                "aciklama": aciklama,
                "birim": birim,
                "miktar": miktar or 0,
                "birim_fiyat": birim_fiyat or 0,
                "et_kod": et_kod,
            })

    return info, kalemler


def _next_value(cells, start_idx):
    """Etiketten sonraki ilk dolu hücreyi döndür"""
    for j in range(start_idx + 1, len(cells)):
        if cells[j] is not None and str(cells[j]).strip() != "":
            return str(cells[j]).strip()
    return ""


def _num(v):
    """Hücre değerini sayıya çevir"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(".", "").replace(",", "."))
    except ValueError:
        return None


# ─── AX Dosyası Üretme ───────────────────────────────────
def build_ax(info, kalemler):
    """Orijinal formatta AX Excel dosyası oluştur"""
    wb = Workbook()
    ws = wb.active
    ws.title = "OZET"

    # Kılavuz çizgilerini kapat — orijinal AX görünümü
    ws.sheet_view.showGridLines = False
    # Sayfa Sonu Önizleme modu — orijinal AX'teki gri dış alan görünümü
    ws.sheet_view.view = "pageBreakPreview"
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.zoomScaleNormal = 100
    ws.sheet_view.zoomScaleSheetLayoutView = 100

    # Sütun genişlikleri
    widths = {"A": 8, "B": 70, "C": 8, "D": 8, "E": 14, "F": 10, "G": 8, "H": 8, "I": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    n = len(kalemler)
    last_row = n + 1  # başlık + veri satırları

    # ── Başlık satırı ──
    headers = ["Poz no", "Açıklama", "Birim", "Miktar", "BF", "Para birimi", "", "", "Madde Kodu"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h if h else None)
        c.font = Font(name="Calibri", size=11, bold=True, color=RED)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = cell_border(
            top=True,
            bottom=(n == 0),
            left=(ci == 1),
            right=(ci == 9),
        )

    # ── Veri satırları ──
    for r_i, k in enumerate(kalemler):
        r = r_i + 2
        is_last = (r == last_row)

        vals = [
            r_i + 1,           # Poz no
            k["aciklama"],     # Açıklama
            "ADET",            # Birim — her zaman ADET
            k["miktar"],       # Miktar
            k["birim_fiyat"],  # BF
            "TRY",             # Para birimi
            None, None,        # boş sütunlar
            k["et_kod"],       # Madde Kodu
        ]

        for ci, v in enumerate(vals, 1):
            c = ws.cell(row=r, column=ci, value=v)
            c.font = Font(name="Calibri", size=11)
            c.border = cell_border(
                bottom=is_last,
                left=(ci == 1),
                right=(ci == 9),
            )
            # Hizalama ve format
            if ci == 1:      # Poz no
                c.alignment = Alignment(horizontal="right", vertical="center")
            elif ci == 3 or ci == 6:  # Birim, Para birimi
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif ci == 4:    # Miktar
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = "#,##0.00"
            elif ci == 5:    # BF
                c.alignment = Alignment(horizontal="right", vertical="center")
                c.number_format = '"₺" #,##0.00'
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")

    # Yazdırma alanı — tablo bölgesi (grinin dışında kalan beyaz alan)
    ws.print_area = f"A1:I{last_row}"

    # Dosyayı belleğe kaydet
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def safe_filename(s):
    """Dosya adı için geçersiz karakterleri temizle"""
    return re.sub(r'[/\\:*?"<>|]', "-", s).strip()


# ─── API Endpoint'leri ───────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/ax", methods=["POST"])
def convert_ax():
    """Teklif Excel'i al, AX Excel'i döndür"""
    if "file" not in request.files:
        return jsonify({"error": "Dosya bulunamadı. 'file' alanında Excel gönderin."}), 400

    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "Sadece .xlsx veya .xlsm dosyaları kabul edilir."}), 400

    try:
        file_bytes = f.read()
        info, kalemler = parse_teklif(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Dosya okunamadı: {str(e)}"}), 422

    if not kalemler:
        return jsonify({"error": "Dosyada tutarı 0'dan büyük kalem bulunamadı."}), 422

    ax_buf = build_ax(info, kalemler)

    parts = ["AX"]
    if info["magaza_kodu"]:
        parts.append(info["magaza_kodu"])
    if info["magaza_adi"]:
        parts.append(info["magaza_adi"])
    if info["srv_no"]:
        parts.append(info["srv_no"])
    fname = safe_filename(" - ".join(parts)) + ".xlsx"

    return send_file(
        ax_buf,
        as_attachment=True,
        download_name=fname,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/ax/preview", methods=["POST"])
def preview_ax():
    """İndirmeden önce önizleme verisi döndür (JSON)"""
    if "file" not in request.files:
        return jsonify({"error": "Dosya bulunamadı."}), 400

    f = request.files["file"]
    try:
        file_bytes = f.read()
        info, kalemler = parse_teklif(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Dosya okunamadı: {str(e)}"}), 422

    return jsonify({
        "info": info,
        "kalemler": kalemler,
        "kalem_sayisi": len(kalemler),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
