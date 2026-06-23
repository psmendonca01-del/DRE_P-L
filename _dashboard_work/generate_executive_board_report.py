import json
import math
import textwrap
import unicodedata
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


ROOT = Path("C:/Users/PauloMendonça/OneDrive - Redefrete/Documentos/DRE")
WORK = ROOT / "_dashboard_work"
DATA = WORK / "dashboard_data.json"
NOTES = WORK / "dashboard_notes.json"
REPORT_IMAGES = WORK / "executive_report_assets"
OUTPUT = ROOT / "output" / "pdf"
LOGO = WORK / "assets" / "redefrete-logo-branco.png"
FONT_DIR = Path("C:/Windows/Fonts")

NAVY = "#071b4d"
BLUE = "#0d5be1"
GREEN = "#0aa36f"
RED = "#e23b3b"
PURPLE = "#7c3aed"
TEAL = "#0891b2"
ORANGE = "#f59e0b"
MUTED = "#63708d"
LINE = "#dfe6f2"
BG = "#eef2f8"
INK = "#00113f"


def font(size=12, bold=False):
    candidates = [
        FONT_DIR / ("arialbd.ttf" if bold else "arial.ttf"),
        FONT_DIR / ("segoeuib.ttf" if bold else "segoeui.ttf"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return ImageFont.truetype(str(candidate), size)
    return ImageFont.load_default()

CLIENT_MAP = {
    "AGROPAD": "REDEFRETE",
    "AMAZON": "AMAZON",
    "AMAZON EDSP": "AMAZON",
    "AMAZON FM": "AMAZON",
    "AMAZON LM": "AMAZON",
    "AMAZON SD": "AMAZON",
    "BIMBO LM": "BIMBO",
    "CASAS BAHIA": "CASAS BAHIA",
    "FAST SHOP": "FAST SHOP",
    "GPA LM": "GPA",
    "MAGALU": "MAGALU",
    "MERCADO LIVRE LM": "MERCADO LIVRE",
    "MERCADO LIVRE RENTAL": "MERCADO LIVRE",
    "PETLOVE": "PETLOVE",
    "PETLOVE FM": "PETLOVE",
    "PETLOVE FM (INATIVO)": "PETLOVE",
    "PETLOVE LM": "PETLOVE",
    "SHOPEE LM": "SHOPEE",
    "SHOPEE": "SHOPEE",
    "SHOPEE LH": "SHOPEE",
    "VIA VAREJO": "VIA VAREJO",
    "REDEFRETE": "REDEFRETE",
}


def repair_text(value):
    text = str(value or "")
    if any(marker in text for marker in ("Ã", "Â")):
        try:
            text = text.encode("latin1").decode("utf-8")
        except UnicodeError:
            pass
    return text


def norm(value):
    text = str(value or "").lower()
    return "".join(ch for ch in unicodedata.normalize("NFKD", text) if not unicodedata.combining(ch))


def is_account(row, text):
    return norm(text) in norm(row.get("conta"))


def is_category(row, text):
    return norm(text) in norm(row.get("categoria"))


def is_iss(row):
    category = norm(row.get("categoria"))
    return category == "iss" or category.startswith("iss ") or "iss retido" in category


def is_irrf(row):
    return "irrf" in norm(row.get("categoria")) or "irrf" in norm(row.get("conta"))


def client_name(row):
    return row.get("clienteDashboard") or CLIENT_MAP.get(str(row.get("projeto") or "").strip(), "Outros")


def create_bucket(period):
    return defaultdict(float, {"period": period})


def fill_bucket(bucket, row):
    value = row["valor"]
    is_pis = is_category(row, "PIS")
    is_cofins = is_category(row, "COFINS")
    iss = is_iss(row)
    irrf = is_irrf(row)
    csll = is_category(row, "CSLL") or is_account(row, "CSLL")
    if is_account(row, "Receita Bruta"):
        bucket["gross"] += value
    if is_pis:
        bucket["pis"] += value
    if is_cofins:
        bucket["cofins"] += value
    if iss:
        bucket["iss"] += value
    if is_account(row, "Deduções de Receita"):
        bucket["revenue_deductions"] += value
    if is_account(row, "Custo dos Serviços"):
        bucket["service_costs"] += value
    if is_account(row, "Outras Receitas"):
        bucket["other_revenue"] += value
    if is_category(row, "Depreciação"):
        bucket["depreciation"] += value
    elif is_account(row, "Despesas Administrativas"):
        bucket["admin"] += value
    if is_account(row, "Despesas com Pessoal"):
        bucket["people"] += value
    if is_account(row, "Despesas de Vendas"):
        bucket["sales"] += value
    if is_account(row, "Outros Tributos") or (is_account(row, "Impostos") and not is_pis and not is_cofins and not iss and not irrf and not csll):
        bucket["other_taxes"] += value
    if is_account(row, "Receitas Financeiras"):
        bucket["financial_revenue"] += value
    if is_account(row, "Despesas Financeiras"):
        bucket["financial_expense"] += value
    if is_account(row, "IRPJ") or is_category(row, "IRPJ") or irrf:
        bucket["irpj"] += value
    if is_category(row, "CSLL (Serviço)"):
        bucket["csll"] += value
    if (is_account(row, "CSLL") or is_category(row, "CSLL")) and not is_category(row, "CSLL (Serviço)"):
        bucket["csll"] += value


def finalize(bucket):
    bucket["taxes"] = bucket["pis"] + bucket["cofins"] + bucket["iss"] + bucket["revenue_deductions"]
    bucket["net_revenue"] = bucket["gross"] + bucket["taxes"]
    bucket["costs"] = bucket["service_costs"] + bucket["other_revenue"]
    bucket["contribution"] = bucket["net_revenue"] + bucket["costs"]
    bucket["expenses"] = bucket["admin"] + bucket["people"] + bucket["sales"] + bucket["other_taxes"]
    bucket["ebitda"] = bucket["contribution"] + bucket["expenses"]
    bucket["ebit"] = bucket["ebitda"] + bucket["depreciation"]
    bucket["financial"] = bucket["financial_revenue"] + bucket["financial_expense"]
    bucket["ebt"] = bucket["ebit"] + bucket["financial"]
    bucket["result_taxes"] = bucket["irpj"] + bucket["csll"]
    bucket["net_result"] = bucket["ebt"] + bucket["result_taxes"]
    return bucket


def aggregate(rows, periods):
    bucket = create_bucket("total")
    allowed = set(periods)
    for row in rows:
        if row["periodo"] in allowed:
            fill_bucket(bucket, row)
    return finalize(bucket)


def month_bucket(rows, period):
    return aggregate(rows, [period])


def fmt_money(value):
    sign = "-" if value < 0 else ""
    value = abs(value) / 1000
    return f"{sign}R$ {value:,.0f}".replace(",", ".")


def fmt_pct(value):
    if value is None or not math.isfinite(value):
        return "-"
    return f"{value * 100:.1f}%".replace(".", ",")


def variation(current, previous):
    if abs(previous) < 0.01:
        return None
    return (current - previous) / abs(previous)


def draw_text_box(pdf, x, y, w, h, title, body, accent=BLUE):
    pdf.setFillColor(colors.white)
    pdf.setStrokeColor(colors.HexColor(LINE))
    pdf.roundRect(x, y, w, h, 8, stroke=1, fill=1)
    pdf.setFillColor(colors.HexColor(accent))
    pdf.roundRect(x, y + h - 5, w, 5, 3, stroke=0, fill=1)
    pdf.setFillColor(colors.HexColor(INK))
    pdf.setFont("Helvetica-Bold", 12)
    pdf.drawString(x + 14, y + h - 24, title)
    pdf.setFillColor(colors.HexColor("#33415f"))
    pdf.setFont("Helvetica", 8.4)
    cy = y + h - 42
    for para in body:
        for line in wrap(para, w - 28, "Helvetica", 8.4):
            if cy < y + 10:
                return
            pdf.drawString(x + 14, cy, line)
            cy -= 10
        cy -= 5


def wrap(text, max_width, font="Helvetica", size=8):
    lines = []
    for raw in repair_text(text).splitlines() or [""]:
        words = raw.split()
        if not words:
            lines.append("")
            continue
        current = words[0]
        for word in words[1:]:
            candidate = f"{current} {word}"
            if stringWidth(candidate, font, size) <= max_width:
                current = candidate
            else:
                lines.append(current)
                current = word
        lines.append(current)
    return lines


def fit_image(iw, ih, bw, bh):
    scale = min(bw / iw, bh / ih)
    return iw * scale, ih * scale


def save_monthly_chart(months, buckets, path):
    img = Image.new("RGB", (1280, 520), "white")
    d = ImageDraw.Draw(img)
    small = font(18)
    tiny = font(15)
    title_font = font(22, True)
    left, top, right, bottom = 80, 50, 40, 80
    plot_w = img.width - left - right
    plot_h = img.height - top - bottom
    values = [b["net_revenue"] for b in buckets] + [b["ebitda"] for b in buckets] + [b["net_result"] for b in buckets]
    max_v = max(values + [1])
    min_v = min(values + [0])
    rng = max(max_v - min_v, 1)
    y = lambda v: top + (max_v - v) / rng * plot_h
    zero = y(0)
    d.line((left, zero, img.width - right, zero), fill="#dfe6f2", width=2)
    d.line((left, top, img.width - right, top), fill="#dfe6f2", width=2)
    d.text((left, 18), "Evolucao mensal - Receita liquida, EBITDA e Resultado liquido (R$ mil)", fill=INK, font=title_font)
    slot = plot_w / max(len(months), 1)
    xs = []
    for i, (month, b) in enumerate(zip(months, buckets)):
        x = left + slot * i + slot / 2
        xs.append(x)
        bar_h = abs(y(max(b["net_revenue"], 0)) - zero)
        d.rounded_rectangle((x - 18, zero - bar_h, x + 18, zero), radius=8, fill=BLUE)
        d.text((x - 22, zero + 18), month[5:] + "/" + month[2:4], fill="#33415f", font=small)
        d.text((x - 28, zero - bar_h - 24), fmt_money(b["net_revenue"]), fill=INK, font=tiny)
    for key, color in [("ebitda", GREEN), ("net_result", PURPLE)]:
        pts = [(xs[i], y(b[key])) for i, b in enumerate(buckets)]
        if len(pts) > 1:
            d.line(pts, fill=color, width=5, joint="curve")
        for pt in pts:
            d.ellipse((pt[0] - 6, pt[1] - 6, pt[0] + 6, pt[1] + 6), fill=color)
    d.rounded_rectangle((left, img.height - 36, left + 28, img.height - 26), radius=5, fill=BLUE)
    d.text((left + 36, img.height - 42), "Receita liquida", fill="#33415f", font=small)
    d.rounded_rectangle((left + 190, img.height - 36, left + 218, img.height - 26), radius=5, fill=GREEN)
    d.text((left + 226, img.height - 42), "EBITDA", fill="#33415f", font=small)
    d.rounded_rectangle((left + 310, img.height - 36, left + 338, img.height - 26), radius=5, fill=PURPLE)
    d.text((left + 346, img.height - 42), "Resultado liquido", fill="#33415f", font=small)
    img.save(path)


def save_donut(data, path):
    img = Image.new("RGB", (680, 430), "white")
    d = ImageDraw.Draw(img)
    small = font(15)
    title_font = font(20, True)
    total = sum(v for _, v in data) or 1
    colors_list = [BLUE, GREEN, ORANGE, PURPLE, TEAL, "#94a3b8"]
    start = -90
    box = (60, 55, 360, 355)
    for i, (_, value) in enumerate(data):
        extent = value / total * 360
        d.pieslice(box, start, start + extent, fill=colors_list[i % len(colors_list)])
        start += extent
    d.ellipse((130, 125, 290, 285), fill="white")
    d.text((156, 178), fmt_money(total), fill=INK, font=title_font)
    d.text((172, 205), "R$ mil", fill=MUTED, font=small)
    d.text((36, 20), "Faturamento bruto por cliente", fill=INK, font=title_font)
    y = 70
    for i, (name, value) in enumerate(data):
        d.ellipse((405, y + 3, 417, y + 15), fill=colors_list[i % len(colors_list)])
        d.text((426, y), name, fill=INK, font=small)
        d.text((560, y), fmt_pct(value / total), fill=INK, font=small)
        d.text((610, y), fmt_money(value), fill=INK, font=small)
        y += 34
    img.save(path)


def save_bridge_chart(bucket, path):
    img = Image.new("RGB", (1280, 430), "white")
    d = ImageDraw.Draw(img)
    small = font(15)
    tiny = font(13)
    title_font = font(22, True)
    labels = ["Receita Líquida", "Custos", "Despesas", "Deprec.", "Financ.", "Imp. Resultado", "Total"]
    values = [bucket["net_revenue"], bucket["costs"], bucket["expenses"], bucket["depreciation"], bucket["financial"], bucket["result_taxes"], bucket["net_result"]]
    max_abs = max(abs(v) for v in values) or 1
    left, top, bottom = 60, 50, 80
    plot_h = img.height - top - bottom
    zero = top + plot_h * 0.5
    scale = plot_h * 0.45 / max_abs
    slot = (img.width - 120) / len(values)
    d.text((left, 20), "Ponte gerencial do resultado do ultimo mes (R$ mil)", fill=INK, font=title_font)
    d.line((left, zero, img.width - 40, zero), fill="#dfe6f2", width=2)
    for i, (label, value) in enumerate(zip(labels, values)):
        x = left + slot * i + slot / 2
        h = abs(value) * scale
        color = GREEN if value >= 0 else RED
        y0, y1 = (zero - h, zero) if value >= 0 else (zero, zero + h)
        if label == "Total":
            color = BLUE if value >= 0 else RED
        d.rounded_rectangle((x - 30, y0, x + 30, y1), radius=7, fill=color)
        d.text((x - 46, y1 + 12 if value >= 0 else y1 + 8), label, fill="#33415f", font=tiny)
        d.text((x - 36, y0 - 22 if value >= 0 else y1 + 24), fmt_money(value), fill=INK, font=small)
    img.save(path)


def page_header(pdf, title, page, total, width, height):
    pdf.setFillColor(colors.HexColor(INK))
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(34, height - 28, title)
    pdf.setFillColor(colors.HexColor(MUTED))
    pdf.setFont("Helvetica", 8)
    pdf.drawRightString(width - 34, height - 26, f"Página {page} de {total}")
    pdf.setStrokeColor(colors.HexColor(LINE))
    pdf.line(34, height - 39, width - 34, height - 39)


def page_footer(pdf, generated, width):
    pdf.setFillColor(colors.HexColor(MUTED))
    pdf.setFont("Helvetica", 7)
    pdf.drawString(34, 18, "Redefrete - Parecer Executivo DRE")
    pdf.drawRightString(width - 34, 18, f"Gerado em {generated}")


def draw_kpi(pdf, x, y, w, h, title, value, note, accent):
    pdf.setFillColor(colors.white)
    pdf.setStrokeColor(colors.HexColor(LINE))
    pdf.roundRect(x, y, w, h, 8, stroke=1, fill=1)
    pdf.setFillColor(colors.HexColor(accent))
    pdf.roundRect(x, y + h - 5, w, 5, 3, stroke=0, fill=1)
    pdf.setFillColor(colors.HexColor(accent))
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(x + 12, y + h - 20, title.upper())
    pdf.setFillColor(colors.HexColor(INK))
    pdf.setFont("Helvetica-Bold", 16)
    pdf.drawString(x + 12, y + h - 43, value)
    pdf.setFillColor(colors.HexColor(MUTED))
    pdf.setFont("Helvetica-Bold", 8)
    pdf.drawString(x + 12, y + 12, note)


def draw_image(pdf, path, x, y, w, h):
    image = ImageReader(str(path))
    iw, ih = image.getSize()
    dw, dh = fit_image(iw, ih, w, h)
    pdf.drawImage(image, x + (w - dw) / 2, y + (h - dh) / 2, width=dw, height=dh, preserveAspectRatio=True)


def load_data():
    data = json.loads(DATA.read_text(encoding="utf-8"))
    notes = json.loads(NOTES.read_text(encoding="utf-8")) if NOTES.exists() else {}
    return data, notes


def main():
    REPORT_IMAGES.mkdir(parents=True, exist_ok=True)
    OUTPUT.mkdir(parents=True, exist_ok=True)
    data, notes = load_data()
    rows = data["rows"]
    periods = data["meta"]["periods"]
    current_year = max(int(p[:4]) for p in periods)
    current_periods = [p for p in periods if p.startswith(str(current_year))]
    same_months_previous = [f"{current_year-1}-{p[5:]}" for p in current_periods]
    current = aggregate(rows, current_periods)
    previous = aggregate(rows, same_months_previous)
    month_buckets = [month_bucket(rows, p) for p in current_periods]
    last = month_buckets[-1]

    monthly_png = REPORT_IMAGES / "monthly.png"
    donut_png = REPORT_IMAGES / "donut.png"
    bridge_png = REPORT_IMAGES / "bridge.png"
    save_monthly_chart(current_periods, month_buckets, monthly_png)
    revenue_by_client = defaultdict(float)
    for row in rows:
        if row["periodo"] in current_periods and is_account(row, "Receita Bruta"):
            revenue_by_client[client_name(row)] += row["valor"]
    donut_data = sorted(revenue_by_client.items(), key=lambda item: item[1], reverse=True)[:6]
    save_donut(donut_data, donut_png)
    save_bridge_chart(last, bridge_png)

    generated = datetime.now().strftime("%d/%m/%Y %H:%M")
    output = OUTPUT / f"parecer_executivo_dre_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
    width, height = landscape(A4)
    pdf = canvas.Canvas(str(output), pagesize=(width, height))
    total_pages = 10
    page = 1

    # Cover
    pdf.setFillColor(colors.HexColor(NAVY))
    pdf.rect(0, 0, width, height, stroke=0, fill=1)
    if LOGO.exists():
        pdf.drawImage(str(LOGO), 42, height - 98, width=118, height=34, mask="auto", preserveAspectRatio=True)
    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 30)
    pdf.drawString(48, height - 170, "Parecer Executivo DRE")
    pdf.setFont("Helvetica-Bold", 20)
    pdf.drawString(48, height - 205, f"Exercício {current_year} - YTD")
    pdf.setFont("Helvetica", 11)
    pdf.drawString(48, height - 245, "Análise consolidada de desempenho, evolução, margens e principais variações.")
    pdf.setFillColor(colors.HexColor(BLUE))
    pdf.roundRect(48, 72, 230, 40, 8, stroke=0, fill=1)
    pdf.setFillColor(colors.white)
    pdf.setFont("Helvetica-Bold", 11)
    pdf.drawString(66, 88, f"Data de emissão: {generated}")
    pdf.showPage()
    page += 1

    # Introduction
    page_header(pdf, "Introdução e Escopo", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    draw_text_box(pdf, 38, 330, 360, 180, "Objetivo do relatório", [
        "Este parecer consolida as informações do DRE em formato executivo para apoiar a leitura do Conselho. A proposta é transformar os dados operacionais e contábeis em uma narrativa gerencial: onde a receita avançou, onde houve pressão de custos e despesas, e como esses movimentos chegaram ao resultado líquido.",
        "Os valores estão expressos em R$ mil e respeitam a estrutura gerencial aplicada no dashboard."
    ], BLUE)
    draw_text_box(pdf, 420, 330, 380, 180, "Critérios utilizados", [
        "A visão de cliente respeita o mapeamento Projeto x Cliente definido para o dashboard, incluindo rateio de REDEFRETE/AGROPAD proporcional ao faturamento mensal dos clientes.",
        "Comparativos trimestrais e semestrais usam o mesmo período do ano anterior. Quando o trimestre ou semestre ainda está em aberto, a comparação é parcial até o mês mais recente disponível."
    ], GREEN)
    draw_text_box(pdf, 38, 112, 762, 170, "Síntese executiva", [
        f"No acumulado de {current_year}, a Receita Líquida totalizou {fmt_money(current['net_revenue'])}, com EBITDA de {fmt_money(current['ebitda'])} e Resultado Líquido de {fmt_money(current['net_result'])}.",
        f"Em relação ao mesmo intervalo de {current_year-1}, a Receita Líquida variou {fmt_pct(variation(current['net_revenue'], previous['net_revenue']))}, enquanto o Resultado Líquido variou {fmt_pct(variation(current['net_result'], previous['net_result']))}.",
        "A leitura central deve separar efeito volume/receita de eficiência operacional: custos, despesas, depreciação e resultado financeiro explicam a qualidade da conversão da receita em lucro."
    ], PURPLE)
    pdf.showPage()
    page += 1

    # Executive KPIs
    page_header(pdf, "Sumário Executivo", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    kpis = [
        ("Receita líquida YTD", fmt_money(current["net_revenue"]), f"{fmt_pct(variation(current['net_revenue'], previous['net_revenue']))} vs {current_year-1}", BLUE),
        ("EBITDA YTD", fmt_money(current["ebitda"]), f"Margem {fmt_pct(current['ebitda']/current['net_revenue'] if current['net_revenue'] else None)}", GREEN),
        ("Resultado líquido", fmt_money(current["net_result"]), f"Margem {fmt_pct(current['net_result']/current['net_revenue'] if current['net_revenue'] else None)}", PURPLE),
        ("Custos + despesas", fmt_money(abs(current["costs"]) + abs(current["expenses"])), f"{fmt_pct((abs(current['costs'])+abs(current['expenses']))/current['net_revenue'] if current['net_revenue'] else None)} da RL", RED),
    ]
    x = 38
    for title, value, note, color in kpis:
        draw_kpi(pdf, x, 435, 182, 82, title, value, note, color)
        x += 194
    draw_image(pdf, monthly_png, 38, 96, 762, 300)
    pdf.showPage()
    page += 1

    # Result bridge
    page_header(pdf, "Ponte do Resultado", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    draw_text_box(pdf, 38, 396, 762, 104, "Como ler esta página", [
        f"A ponte abaixo explica o resultado do último mês disponível ({current_periods[-1][5:]}/{current_periods[-1][:4]}). A soma de Receita Líquida, Custos, Despesas, Depreciação, Resultado Financeiro e Impostos antes do Resultado fecha no Resultado Líquido.",
        "Esta leitura é útil para separar performance operacional recorrente de efeitos financeiros e tributários."
    ], RED)
    draw_image(pdf, bridge_png, 38, 95, 762, 270)
    pdf.showPage()
    page += 1

    # Client analysis
    page_header(pdf, "Receita por Cliente", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    draw_image(pdf, donut_png, 38, 122, 382, 360)
    top_client = donut_data[0] if donut_data else ("N/D", 0)
    total_gross = sum(v for _, v in donut_data) or 1
    draw_text_box(pdf, 450, 310, 350, 172, "Concentração de faturamento", [
        f"O principal cliente no acumulado é {top_client[0]}, com {fmt_money(top_client[1])}, equivalente a {fmt_pct(top_client[1]/total_gross)} do faturamento bruto mapeado entre os principais clientes.",
        "Essa composição deve ser acompanhada sob a ótica de concentração comercial, sazonalidade de campanhas e margem por operação."
    ], TEAL)
    draw_text_box(pdf, 450, 122, 350, 150, "Leitura executiva", [
        "A análise por cliente deve ser cruzada com custos e despesas alocados, pois crescimento de receita não implica, necessariamente, maior geração de resultado.",
        "O rateio das estruturas REDEFRETE/AGROPAD foi aplicado mensalmente conforme a participação de faturamento de cada cliente."
    ], BLUE)
    pdf.showPage()
    page += 1

    # DRE table
    page_header(pdf, "DRE Gerencial - Acumulado", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    rows_table = [
        ("Receita Bruta", current["gross"]),
        ("Deduções da Receita", current["taxes"]),
        ("Receita Líquida", current["net_revenue"]),
        ("Custos", current["costs"]),
        ("Margem de Contribuição", current["contribution"]),
        ("Despesas", current["expenses"]),
        ("EBITDA", current["ebitda"]),
        ("Depreciação", current["depreciation"]),
        ("Resultado Financeiro", current["financial"]),
        ("Impostos antes do Resultado", current["result_taxes"]),
        ("Resultado Líquido", current["net_result"]),
    ]
    x, y = 50, 460
    pdf.setFont("Helvetica-Bold", 9)
    pdf.setFillColor(colors.HexColor(INK))
    pdf.drawString(x, y, "Linha DRE")
    pdf.drawRightString(450, y, f"{current_year} YTD")
    pdf.drawRightString(560, y, f"{current_year-1} comp.")
    pdf.drawRightString(700, y, "Variação")
    y -= 14
    pdf.setStrokeColor(colors.HexColor(LINE))
    pdf.line(x, y, 760, y)
    y -= 16
    previous_map = {
        "Receita Bruta": previous["gross"], "Deduções da Receita": previous["taxes"], "Receita Líquida": previous["net_revenue"],
        "Custos": previous["costs"], "Margem de Contribuição": previous["contribution"], "Despesas": previous["expenses"],
        "EBITDA": previous["ebitda"], "Depreciação": previous["depreciation"], "Resultado Financeiro": previous["financial"],
        "Impostos antes do Resultado": previous["result_taxes"], "Resultado Líquido": previous["net_result"],
    }
    for name, value in rows_table:
        prev = previous_map[name]
        var = variation(value, prev)
        pdf.setFillColor(colors.HexColor(INK))
        pdf.setFont("Helvetica-Bold" if name in ("Receita Líquida", "EBITDA", "Resultado Líquido") else "Helvetica", 8.5)
        pdf.drawString(x, y, name)
        pdf.drawRightString(450, y, fmt_money(value))
        pdf.drawRightString(560, y, fmt_money(prev))
        pdf.setFillColor(colors.HexColor(GREEN if (var or 0) >= 0 else RED))
        pdf.drawRightString(700, y, fmt_pct(var))
        y -= 18
    pdf.showPage()
    page += 1

    # Notes
    page_header(pdf, "Notas Explicativas", page, total_pages, width, height)
    page_footer(pdf, generated, width)
    note_items = [v for v in notes.values() if str(v.get("note", "")).strip()]
    if not note_items:
        draw_text_box(pdf, 38, 300, 762, 150, "Notas", ["Não há notas explicativas cadastradas nos cards."], BLUE)
    else:
        y = 460
        for note in note_items[:4]:
            title = repair_text(note.get("title", "Nota explicativa"))
            body = repair_text(note.get("note", ""))
            draw_text_box(pdf, 38, y - 120, 762, 106, title, textwrap.wrap(body.replace("\n", " "), width=150)[:4], ORANGE)
            y -= 132
            if y < 110:
                break
    pdf.showPage()
    page += 1

    # Appendix screenshots
    for title, image_name in [
        ("Anexo - Dashboard", "dashboard.png"),
        ("Anexo - Evolução", "evolution.png"),
        ("Anexo - Indicadores", "indicators.png"),
    ]:
        page_header(pdf, title, page, total_pages, width, height)
        page_footer(pdf, generated, width)
        image_path = WORK / "report" / image_name
        if image_path.exists():
            draw_image(pdf, image_path, 34, 42, width - 68, height - 90)
        pdf.showPage()
        page += 1

    pdf.save()
    print(output)


if __name__ == "__main__":
    main()
