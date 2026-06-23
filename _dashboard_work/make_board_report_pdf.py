import json
import sys
from datetime import datetime
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


def fit_image(image_width, image_height, box_width, box_height):
    scale = min(box_width / image_width, box_height / image_height)
    return image_width * scale, image_height * scale


def repair_text(value):
    text = str(value or "")
    if any(marker in text for marker in ("Ã", "Â")):
        try:
            text = text.encode("latin1").decode("utf-8")
        except UnicodeError:
            pass
    return text


def wrap_line(text, max_width, font_name="Helvetica", font_size=7.2):
    words = repair_text(text).split()
    if not words:
        return [""]
    lines = []
    current = words[0]
    for word in words[1:]:
        candidate = f"{current} {word}"
        if stringWidth(candidate, font_name, font_size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    lines.append(current)
    return lines


def build_note_lines(notes, max_width):
    if not notes:
        return [{"text": "Sem notas explicativas cadastradas.", "bold": False}]

    lines = []
    for note in notes:
        title = repair_text(note.get("title") or "Nota explicativa")
        lines.append({"text": title, "bold": True})
        body = repair_text(note.get("note") or "").strip()
        for paragraph in body.splitlines():
            if not paragraph.strip():
                lines.append({"text": "", "bold": False})
                continue
            for wrapped in wrap_line(paragraph, max_width - 8):
                lines.append({"text": wrapped, "bold": False})
        lines.append({"text": "", "bold": False})
    while lines and not lines[-1]["text"]:
        lines.pop()
    return lines or [{"text": "Sem notas explicativas cadastradas.", "bold": False}]


def draw_header(pdf, title, page_number, page_count, width, height):
    pdf.setFillColor(colors.HexColor("#00113f"))
    pdf.setFont("Helvetica-Bold", 13)
    pdf.drawString(28, height - 28, title)
    pdf.setFillColor(colors.HexColor("#63708d"))
    pdf.setFont("Helvetica", 8)
    pdf.drawRightString(width - 28, height - 26, f"Pagina {page_number} de {page_count}")
    pdf.setStrokeColor(colors.HexColor("#dfe6f2"))
    pdf.setLineWidth(0.6)
    pdf.line(28, height - 38, width - 28, height - 38)


def draw_footer(pdf, generated_at, width):
    pdf.setFillColor(colors.HexColor("#63708d"))
    pdf.setFont("Helvetica", 7)
    pdf.drawString(28, 16, "Redefrete - Demonstrativo de Resultados DRE")
    pdf.drawRightString(width - 28, 16, f"Gerado em {generated_at}")


def draw_notes_box(pdf, lines, x, y, width, height, start=0):
    pdf.setStrokeColor(colors.HexColor("#dfe6f2"))
    pdf.setFillColor(colors.HexColor("#fbfcff"))
    pdf.roundRect(x, y, width, height, 6, stroke=1, fill=1)
    pdf.setFillColor(colors.HexColor("#00113f"))
    pdf.setFont("Helvetica-Bold", 8.5)
    pdf.drawString(x + 10, y + height - 15, "Notas explicativas")

    line_height = 8.4
    cursor_y = y + height - 29
    max_y = y + 9
    index = start
    while index < len(lines) and cursor_y >= max_y:
        line = lines[index]
        pdf.setFont("Helvetica-Bold" if line.get("bold") else "Helvetica", 7.2)
        pdf.setFillColor(colors.HexColor("#00113f") if line.get("bold") else colors.HexColor("#33415f"))
        pdf.drawString(x + 10, cursor_y, line["text"])
        cursor_y -= line_height
        index += 1
    return index


def draw_notes_page(pdf, title, lines, start, generated_at, page_number, page_count, width, height):
    draw_header(pdf, f"Notas explicativas - {title}", page_number, page_count, width, height)
    draw_footer(pdf, generated_at, width)
    x = 28
    y = 34
    box_w = width - 56
    box_h = height - 86
    return draw_notes_box(pdf, lines, x, y, box_w, box_h, start)


def continuation_count(lines, first_capacity, continuation_capacity):
    remaining = max(0, len(lines) - first_capacity)
    if remaining == 0:
        return 0
    return (remaining + continuation_capacity - 1) // continuation_capacity


def main():
    if len(sys.argv) < 2:
        raise SystemExit("Uso: make_board_report_pdf.py manifest.json")

    manifest_path = Path(sys.argv[1])
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    output_dir = Path(manifest["outputDir"])
    output_dir.mkdir(parents=True, exist_ok=True)

    generated_dt = datetime.now()
    generated_label = generated_dt.strftime("%d/%m/%Y %H:%M")
    output = output_dir / f"relatorio_conselho_dre_{generated_dt.strftime('%Y%m%d_%H%M')}.pdf"

    width, height = landscape(A4)
    pdf = canvas.Canvas(str(output), pagesize=(width, height))
    pages = manifest["pages"]
    note_width = width - 76
    initial_note_height = 138
    continuation_note_height = height - 86
    line_height = 8.4
    first_capacity = int((initial_note_height - 38) / line_height)
    continuation_capacity = int((continuation_note_height - 38) / line_height)
    note_lines_by_page = [build_note_lines(page.get("notes", []), note_width) for page in pages]
    page_count = len(pages) + sum(
        continuation_count(lines, first_capacity, continuation_capacity)
        for lines in note_lines_by_page
    )
    current_page = 1

    for page, note_lines in zip(pages, note_lines_by_page):
        draw_header(pdf, repair_text(page["title"]), current_page, page_count, width, height)
        draw_footer(pdf, generated_label, width)

        image = ImageReader(page["image"])
        image_width, image_height = image.getSize()
        box_x = 24
        notes_x = 28
        notes_y = 31
        notes_w = width - 56
        notes_h = initial_note_height
        box_y = notes_y + notes_h + 10
        box_w = width - 48
        box_h = height - box_y - 44
        draw_w, draw_h = fit_image(image_width, image_height, box_w, box_h)
        x = box_x + (box_w - draw_w) / 2
        y = box_y + (box_h - draw_h) / 2
        pdf.drawImage(image, x, y, width=draw_w, height=draw_h, preserveAspectRatio=True, anchor="c")

        next_note = draw_notes_box(pdf, note_lines, notes_x, notes_y, notes_w, notes_h, 0)
        pdf.showPage()
        current_page += 1

        while next_note < len(note_lines):
            next_note = draw_notes_page(
                pdf,
                repair_text(page["title"]),
                note_lines,
                next_note,
                generated_label,
                current_page,
                page_count,
                width,
                height,
            )
            pdf.showPage()
            current_page += 1

    pdf.save()
    print(output)


if __name__ == "__main__":
    main()
