from __future__ import annotations

import io
import re
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, Tuple

import fitz  # PyMuPDF
from PIL import Image
import pytesseract
from openpyxl import load_workbook

MONTHS = {
    "ENE": 1, "ENERO": 1, "JAN": 1,
    "FEB": 2, "FEBRERO": 2,
    "MAR": 3, "MARZO": 3,
    "ABR": 4, "ABRIL": 4, "APR": 4,
    "MAY": 5, "MAYO": 5,
    "JUN": 6, "JUNIO": 6,
    "JUL": 7, "JULIO": 7,
    "AGO": 8, "AGOSTO": 8, "AUG": 8,
    "SEP": 9, "SEPT": 9, "SEPTIEMBRE": 9,
    "OCT": 10, "OCTUBRE": 10,
    "NOV": 11, "NOVIEMBRE": 11,
    "DIC": 12, "DICIEMBRE": 12, "DEC": 12,
}

@dataclass
class ExtractedData:
    nombre: Optional[str] = None
    cedula: Optional[str] = None
    ingreso: Optional[str] = None          # YYYY-MM-DD
    valor_credito: Optional[float] = None
    seguro_vida_mensual: Optional[int] = None
    extraprima: Optional[str] = None
    fecha_nacimiento: Optional[str] = None # YYYY-MM-DD
    raw_confidence_notes: Optional[str] = None


def _normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text or " ").strip()


def _normalize_doc(raw: str | None) -> Optional[str]:
    if not raw:
        return None
    digits = re.sub(r"\D", "", raw).lstrip("0")
    if not digits:
        return None
    return f"{int(digits):,}".replace(",", ".")


def _parse_money(raw: str | None) -> Optional[float]:
    if not raw:
        return None
    try:
        return float(raw.replace("$", "").replace(",", "").strip())
    except ValueError:
        return None


def _parse_date(raw: str | None) -> Optional[str]:
    if not raw:
        return None
    s = raw.strip().upper().replace(".", "")
    # 05-06-2026 or 05/06/2026
    m = re.search(r"(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})", s)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 1900 if y > 30 else 2000
        return datetime(y, mo, d).strftime("%Y-%m-%d")
    # 05/JUN/2026 or 20-AGO-1981
    m = re.search(r"(\d{1,2})[-/]([A-ZÁÉÍÓÚÑ]{3,10})[-/](\d{2,4})", s)
    if m:
        d = int(m.group(1)); mon = m.group(2).replace("Á", "A").replace("É", "E").replace("Í", "I").replace("Ó", "O").replace("Ú", "U")
        y = int(m.group(3))
        if y < 100:
            y += 1900 if y > 30 else 2000
        mo = MONTHS.get(mon[:3]) or MONTHS.get(mon)
        if mo:
            return datetime(y, mo, d).strftime("%Y-%m-%d")
    return None


def extract_pdf_text(pdf_path: str | Path, ocr_first_pages: int = 3) -> Tuple[str, str]:
    """Returns (embedded_text, ocr_text). OCR is applied to scanned/empty pages only."""
    doc = fitz.open(str(pdf_path))
    embedded_parts = []
    ocr_parts = []
    for idx, page in enumerate(doc):
        embedded = page.get_text("text") or ""
        embedded_parts.append(f"\n--- PAGE {idx+1} TEXT ---\n{embedded}")
        if idx < ocr_first_pages and len(embedded.strip()) < 30:
            # 2.5x is usually enough; increasing helps OCR but slows down processing.
            pix = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                ocr = pytesseract.image_to_string(img, lang="eng", config="--psm 6", timeout=25)
            except (RuntimeError, EnvironmentError, OSError) as e:
                if "tesseract" in str(e).lower():
                    raise EnvironmentError(
                        "Tesseract no está instalado. Instálalo con: brew install tesseract"
                    ) from e
                ocr = ""
            ocr_parts.append(f"\n--- PAGE {idx+1} OCR ---\n{ocr}")
    return "\n".join(embedded_parts), "\n".join(ocr_parts)


def extract_fields(pdf_path: str | Path) -> Dict[str, Any]:
    embedded, ocr = extract_pdf_text(pdf_path)
    all_text = _normalize_spaces(embedded + "\n" + ocr)
    embedded_norm = _normalize_spaces(embedded)
    ocr_norm = _normalize_spaces(ocr)
    data = ExtractedData()

    # Projection page: reliable for nombre, documento, ingreso, valor, seguro mensual.
    m = re.search(r"\b36\s+([A-ZÁÉÍÓÚÑ ]{8,90}?)\s+1008\b", embedded_norm)
    if not m:
        m = re.search(r"\b\d{6,}\s+\d{2}-\d{2}-\d{4}\s+\d+\s+([A-ZÁÉÍÓÚÑ ]{8,90}?)\s+1008\b", embedded_norm)
    data.nombre = m.group(1).strip() if m else None

    m = re.search(r"\b(0{3,}\d{6,12})\b", embedded_norm) or re.search(r"(?:C[. ]?C|DOCUMENTO).*?(\d[\d .]{5,15})", all_text, re.I)
    data.cedula = _normalize_doc(m.group(1)) if m else None

    m = re.search(r"(\d{2}-\d{2}-\d{4})\s+\d+\s+[A-ZÁÉÍÓÚÑ ]+\s+1008", embedded_norm)
    if not m:
        m = re.search(r"(\d{1,2}/[A-Z]{3}\.?/\d{4})", embedded_norm, re.I)
    data.ingreso = _parse_date(m.group(1)) if m else None

    m = re.search(r"1008\s+PMO\s+CONSUMO\s+SEG\s+(\d{1,3}(?:,\d{3})+\.\d{2})", embedded_norm, re.I)
    data.valor_credito = _parse_money(m.group(1)) if m else None

    # IMPORTANT BUSINESS RULE:
    # The Excel column named JUNIO / seguro_vida_mensual must use the initial
    # proportional insurance charged at disbursement, not the first amortization
    # row. In the projection PDF this appears at the bottom as:
    # "Seguro Cartera proporcional 2,937.00".
    m = re.search(
        r"Seguro\s+(?:Cartera|Vida)\s+proporcional\s+(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)",
        embedded_norm,
        re.I,
    )
    if not m:
        # Some PDF text extractors output the total before the description.
        m = re.search(
            r"(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)\s+Seguro\s+(?:Cartera|Vida)\s+proporcional",
            embedded_norm,
            re.I,
        )
    if m:
        data.seguro_vida_mensual = int(round(_parse_money(m.group(1)) or 0))
    else:
        # Fallback only: old logic used the Seguro Vida value from cycle 1.
        m = re.search(r"\b1\s+20\d{4}\s+[\d,]+\s+0\s+[\d,]+\s+([\d,]+)\s+0\s+[\d,]+\s+[\d,]+", embedded_norm)
        data.seguro_vida_mensual = int(m.group(1).replace(",", "")) if m else None

    # Declaracion page OCR: fecha de nacimiento can be handwritten as 20/08/81.
    # Look for a handwritten/typed DOB near the FECHA NACIMIENTO label.
    m = None
    pos = ocr_norm.upper().find("FECHA NACIMIENTO")
    if pos >= 0:
        window = ocr_norm[pos:pos + 220]
        m = re.search(r"(\d{1,2}[-/][A-Z]{3}[-/]\d{2,4}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", window, re.I)
    if not m:
        # In this form it often appears right after the ID number. OCR can confuse 9/4 and 6/2, so keep it loose.
        m = re.search(r"\b[0-9. ]{7,15}\s+[|I]?\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\s+[|I]?", ocr_norm, re.I)
    data.fecha_nacimiento = _parse_date(m.group(1)) if m else None

    # Extraprima rule:
    # If the health declaration has at least one disease/lesion marked under "SI",
    # Excel should say "EXTRAPRIMA INCLUIDA". For the current sample all health
    # conditions are marked under "NO", so this remains blank.
    data.extraprima = "EXTRAPRIMA INCLUIDA" if _detect_medical_yes(pdf_path, ocr_norm) else ""
    notes = []
    if not data.fecha_nacimiento:
        notes.append("Fecha de nacimiento no detectada con alta confianza; revisar manualmente.")
    if not data.valor_credito:
        notes.append("Valor del credito no detectado.")
    data.raw_confidence_notes = " ".join(notes) or "OK"
    result = asdict(data)
    result["ingreso_excel"] = to_excel_date_display(data.ingreso)
    result["fecha_nacimiento_excel"] = to_excel_date_display(data.fecha_nacimiento)
    return result



def _detect_medical_yes(pdf_path: str | Path, ocr_norm: str = "") -> bool:
    """Best-effort detection of any positive checkbox in the medical history table.

    This combines OCR keywords with a conservative image heuristic for the known
    Previsora insurance health declaration format. It intentionally prefers
    false negatives over false positives because charging extraprima incorrectly
    is worse than asking for manual review.
    """
    txt = (ocr_norm or "").upper()
    positive_words = [
        "ENFERMEDAD DEL CORAZON SI",
        "DIABETES SI",
        "CANCER SI",
        "PARALISIS SI",
        "DROGADICCION SI",
        "VIH/SIDA SI",
    ]
    if any(w in txt for w in positive_words):
        return True

    try:
        doc = fitz.open(str(pdf_path))
        if len(doc) < 2:
            return False
        page = doc[1]
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
        img = Image.open(io.BytesIO(pix.tobytes("png"))).convert("L")
        w, h = img.size
        # Approximate SI checkbox columns in the two medical tables on page 2.
        # The bands are intentionally narrow and the threshold high.
        bands = [
            (0.755, 0.785),  # left table SI column
            (0.930, 0.960),  # right table SI column
        ]
        y1, y2 = int(h * 0.515), int(h * 0.705)
        for x1r, x2r in bands:
            x1, x2 = int(w * x1r), int(w * x2r)
            crop = img.crop((x1, y1, x2, y2))
            # Count very dark pixels. Empty boxes and printed vertical lines have
            # little ink; a handwritten/marked X creates a denser cluster.
            pixels = crop.load()
            dark = 0
            total = crop.size[0] * crop.size[1]
            for yy in range(crop.size[1]):
                for xx in range(crop.size[0]):
                    if pixels[xx, yy] < 80:
                        dark += 1
            if total and dark / total > 0.09:
                return True
    except Exception:
        return False
    return False


def fill_excel_template(template_path: str | Path, output_path: str | Path, data: Dict[str, Any], month_sheet: str = "JUNIO") -> str:
    from copy import copy

    wb = load_workbook(template_path)
    ws = wb[month_sheet]
    # Expected columns: A NOMBRE, D CEDULA, E INGRESO, F VALOR CREDITO, G MES, H EXTRAPRIMA, I FECHA NAC
    row = 4
    while ws.cell(row=row, column=1).value:
        row += 1

    # Copy the style from the row above so generated rows keep the same Excel layout.
    template_row = row - 1 if row > 4 else 4
    for col in range(1, ws.max_column + 1):
        src = ws.cell(row=template_row, column=col)
        dst = ws.cell(row=row, column=col)
        if src.has_style:
            dst._style = copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy(src.alignment)
        if src.border:
            dst.border = copy(src.border)
        if src.fill:
            dst.fill = copy(src.fill)
        if src.font:
            dst.font = copy(src.font)

    ws.cell(row=row, column=1).value = data.get("nombre")
    ws.cell(row=row, column=4).value = data.get("cedula")
    ws.cell(row=row, column=5).value = _parse_date_to_datetime(data.get("ingreso"))
    ws.cell(row=row, column=6).value = data.get("valor_credito")
    ws.cell(row=row, column=7).value = data.get("seguro_vida_mensual")
    ws.cell(row=row, column=8).value = data.get("extraprima")
    ws.cell(row=row, column=9).value = _parse_date_to_datetime(data.get("fecha_nacimiento"))

    # Force date cells as real Excel dates, matching the other sheets: MM-DD-YY.
    ws.cell(row=row, column=5).number_format = "mm-dd-yy"
    ws.cell(row=row, column=9).number_format = "mm-dd-yy"
    wb.save(output_path)
    return str(output_path)


def _parse_date_to_datetime(value: Optional[str]):
    if not value:
        return None
    return datetime.strptime(value, "%Y-%m-%d")


def to_excel_date_display(value: Optional[str]) -> str:
    """Return MM-DD-YY text for copy/paste preview, matching the template display."""
    if not value:
        return ""
    try:
        return datetime.strptime(value, "%Y-%m-%d").strftime("%m-%d-%y")
    except Exception:
        return ""

if __name__ == "__main__":
    import argparse, json
    parser = argparse.ArgumentParser()
    parser.add_argument("pdf")
    parser.add_argument("--template")
    parser.add_argument("--out", default="salida.xlsx")
    parser.add_argument("--month", default="JUNIO")
    args = parser.parse_args()
    result = extract_fields(args.pdf)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.template:
        fill_excel_template(args.template, args.out, result, args.month)
        print(f"Excel generado: {args.out}")
