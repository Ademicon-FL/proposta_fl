"""
app.py — Gerador de Proposta PDF · Ademicon Crédito Estruturado
"""
import io
import datetime

import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import Flowable

from leitor import ler_planilha

# ──────────────────────────────────────────────────────────────────────────────
# Palette
# ──────────────────────────────────────────────────────────────────────────────
C_DARK_BLUE = colors.HexColor("#1F3864")
C_MED_BLUE  = colors.HexColor("#2F75B6")
C_LITE_BLUE = colors.HexColor("#D6E4F0")
C_GREEN     = colors.HexColor("#E2EFDA")
C_GRAY      = colors.HexColor("#F5F5F5")
C_WHITE     = colors.white
C_BLACK     = colors.black

PAGE_W, PAGE_H = A4
MARGIN = 18 * mm


# ──────────────────────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────────────────────
def _brl(val, prefix="R$ "):
    """Format number as BRL currency."""
    if val is None:
        return "—"
    try:
        return f"{prefix}{float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


def _pct(val, decimals=2):
    """Format number as percentage."""
    if val is None:
        return "—"
    try:
        return f"{float(val)*100:.{decimals}f}%".replace(".", ",")
    except Exception:
        return str(val)


def _fmt_mes(dt):
    """Format datetime to MM/YYYY."""
    if isinstance(dt, datetime.datetime):
        return dt.strftime("%m/%Y")
    return str(dt)


def _fmt_num(val, decimals=0):
    """Format a plain number."""
    if val is None:
        return "—"
    try:
        if decimals == 0:
            return f"{int(round(float(val))):,}".replace(",", ".")
        return f"{float(val):,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(val)


# ──────────────────────────────────────────────────────────────────────────────
# PDF Styles
# ──────────────────────────────────────────────────────────────────────────────
def make_styles():
    base = getSampleStyleSheet()

    def ps(name, parent="Normal", **kw):
        kw.setdefault("fontName", "Helvetica")
        return ParagraphStyle(name, parent=base[parent], **kw)

    return {
        "title": ps("Title", fontSize=16, textColor=C_WHITE, alignment=TA_CENTER,
                    fontName="Helvetica-Bold", leading=20),
        "subtitle": ps("Subtitle", fontSize=9, textColor=C_WHITE, alignment=TA_CENTER,
                       leading=12),
        "section": ps("Section", fontSize=11, textColor=C_DARK_BLUE,
                      fontName="Helvetica-Bold", spaceBefore=8, spaceAfter=4, leading=14),
        "info_label": ps("InfoLabel", fontSize=9, textColor=C_DARK_BLUE,
                         fontName="Helvetica-Bold", leading=12),
        "info_val": ps("InfoVal", fontSize=9, textColor=C_BLACK, leading=12),
        "body": ps("Body", fontSize=8.5, textColor=C_BLACK, leading=12),
        "disclaimer": ps("Disclaimer", fontSize=7.5, textColor=colors.HexColor("#555555"),
                         fontName="Helvetica-Oblique", leading=11, spaceBefore=6),
        "footer": ps("Footer", fontSize=8, textColor=C_DARK_BLUE, alignment=TA_CENTER, leading=12),
        "footer_small": ps("FooterSmall", fontSize=7, textColor=colors.gray,
                           alignment=TA_CENTER, leading=10),
    }


# ──────────────────────────────────────────────────────────────────────────────
# Cover / Header block
# ──────────────────────────────────────────────────────────────────────────────
class CoverHeader(Flowable):
    """Blue banner at the top of the PDF."""
    def __init__(self, width, height=28*mm):
        super().__init__()
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        c.setFillColor(C_DARK_BLUE)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)

        c.setFillColor(C_WHITE)
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(self.width / 2, self.height / 2 + 4, "PROPOSTA DE CRÉDITO ESTRUTURADO")
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.HexColor("#A8C4E0"))
        c.drawCentredString(self.width / 2, self.height / 2 - 8, "Consórcio Imobiliário · Ademicon")


class SectionHeader(Flowable):
    """Blue bar for section titles."""
    def __init__(self, text, width, height=10*mm):
        super().__init__()
        self.text = text
        self.width = width
        self.height = height

    def draw(self):
        c = self.canv
        c.setFillColor(C_MED_BLUE)
        c.rect(0, 0, self.width, self.height, fill=1, stroke=0)
        c.setFillColor(C_WHITE)
        c.setFont("Helvetica-Bold", 10)
        c.drawString(4 * mm, self.height / 2 - 3, self.text.upper())


# ──────────────────────────────────────────────────────────────────────────────
# Table helpers
# ──────────────────────────────────────────────────────────────────────────────
def _base_table_style(header_rows=1):
    cmds = [
        # Header
        ("BACKGROUND",    (0, 0), (-1, header_rows - 1), C_DARK_BLUE),
        ("TEXTCOLOR",     (0, 0), (-1, header_rows - 1), C_WHITE),
        ("FONTNAME",      (0, 0), (-1, header_rows - 1), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, header_rows - 1), 8),
        ("ALIGN",         (0, 0), (-1, header_rows - 1), "CENTER"),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0, header_rows), (-1, -1), [C_WHITE, C_GRAY]),
        ("FONTNAME",      (0, header_rows), (-1, -1), "Helvetica"),
        ("FONTSIZE",      (0, header_rows), (-1, -1), 8),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]
    return cmds


# ──────────────────────────────────────────────────────────────────────────────
# Section builders
# ──────────────────────────────────────────────────────────────────────────────
CONTENT_WIDTH = PAGE_W - 2 * MARGIN


def build_info_block(cfg, dados, styles):
    """Client/proposal info block."""
    ref_date = cfg["data_referencia"].strftime("%d/%m/%Y")
    gen_date = datetime.date.today().strftime("%d/%m/%Y")

    rows = [
        ["Cliente:", cfg["nome_cliente"],      "Data de Referência:", ref_date],
        ["Gerente:", cfg["gerente"],            "Data de Geração:",    gen_date],
        ["Cargo:",   cfg["cargo"],              "Tipo:",               "Consórcio Imobiliário"],
        ["Unidade:", cfg["unidade"],            "",                    ""],
    ]

    col_w = [30*mm, 60*mm, 42*mm, 38*mm]
    data = []
    for r in rows:
        data.append([
            Paragraph(r[0], styles["info_label"]),
            Paragraph(r[1], styles["info_val"]),
            Paragraph(r[2], styles["info_label"]),
            Paragraph(r[3], styles["info_val"]),
        ])

    t = Table(data, colWidths=col_w)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), C_LITE_BLUE),
        ("VALIGN",     (0, 0), (-1, -1), "MIDDLE"),
        ("GRID",       (0, 0), (-1, -1), 0.3, C_MED_BLUE),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]))
    return t


def build_resumo_executivo(dados, cfg, styles):
    """Resumo executivo: TIR, crédito total, parcela, prazo."""
    elems = []
    elems.append(Spacer(1, 4*mm))
    elems.append(SectionHeader("Resumo Executivo", CONTENT_WIDTH))
    elems.append(Spacer(1, 3*mm))

    resumo = dados.get("resumo") or {}
    fluxo  = dados.get("fluxo")  or {}

    credito = resumo.get("credito_total") or fluxo.get("credito_total")
    parcela = resumo.get("valor_parcela") or fluxo.get("parcela")
    prazo   = resumo.get("qtd_parcelas")
    tir_m   = resumo.get("tir_mensal") or fluxo.get("tir_mensal")
    tir_a   = resumo.get("tir_anual")  or fluxo.get("tir_anual")
    taxa_e  = resumo.get("taxa_estatica")

    versao_tir = cfg.get("versao_tir", "Com FIDC")

    kpis = []
    kpis.append(("Crédito Total Levantado", _brl(credito)))
    kpis.append(("Valor da Parcela (Pré-Contemplação)", _brl(parcela)))
    kpis.append(("Quantidade de Parcelas", _fmt_num(prazo)))

    if versao_tir in ("Com FIDC", "Ambas"):
        kpis.append(("TIR Mensal (com FIDC)", _pct(tir_m)))
        kpis.append(("TIR Anual (com FIDC)", _pct(tir_a)))
    if versao_tir in ("Sem FIDC", "Ambas"):
        kpis.append(("Taxa Estática Mensal", _pct(taxa_e)))

    # 2-column KPI grid
    half = (len(kpis) + 1) // 2
    left  = kpis[:half]
    right = kpis[half:]

    while len(right) < len(left):
        right.append(("", ""))

    rows = []
    for (lk, lv), (rk, rv) in zip(left, right):
        rows.append([
            Paragraph(f"<b>{lk}</b>", styles["body"]),
            Paragraph(lv, styles["body"]),
            Paragraph(f"<b>{rk}</b>", styles["body"]) if rk else Paragraph("", styles["body"]),
            Paragraph(rv, styles["body"]) if rv else Paragraph("", styles["body"]),
        ])

    col_w = [55*mm, 35*mm, 55*mm, 25*mm]
    t = Table(rows, colWidths=col_w)
    cmds = [
        ("BACKGROUND",    (0, 0), (-1, -1), C_LITE_BLUE),
        ("GRID",          (0, 0), (-1, -1), 0.3, C_MED_BLUE),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("FONTSIZE",      (0, 0), (-1, -1), 8.5),
    ]
    # Alternate row backgrounds
    for i in range(len(rows)):
        bg = C_LITE_BLUE if i % 2 == 0 else C_WHITE
        cmds.append(("BACKGROUND", (0, i), (-1, i), bg))

    t.setStyle(TableStyle(cmds))
    elems.append(t)
    return elems


def build_custo_fidc(dados, styles):
    """FIDC cost block."""
    elems = []
    elems.append(Spacer(1, 4*mm))
    elems.append(SectionHeader("Custo FIDC", CONTENT_WIDTH))
    elems.append(Spacer(1, 3*mm))

    fluxo = dados.get("fluxo") or {}
    resumo = dados.get("resumo") or {}
    carteira = dados.get("carteira") or {}

    credito = fluxo.get("credito_total") or resumo.get("credito_total")
    taxa_fidc = fluxo.get("taxa_fidc")
    lance_embutido = None

    for k, v in (carteira.get("totais") or {}).items():
        if "total carteira" in k.lower() or "lance livre" in k.lower():
            if v.get("lance_embutido"):
                lance_embutido = v["lance_embutido"]
                break

    fee_total = None
    if credito and taxa_fidc:
        try:
            fee_total = float(credito) * float(taxa_fidc)
        except Exception:
            pass

    rows = [
        [Paragraph("<b>Indicador</b>", styles["body"]), Paragraph("<b>Valor</b>", styles["body"])],
        [Paragraph("Taxa FIDC (fee sobre crédito)", styles["body"]), Paragraph(_pct(taxa_fidc), styles["body"])],
        [Paragraph("Crédito Total (base de cálculo)", styles["body"]), Paragraph(_brl(credito), styles["body"])],
        [Paragraph("Fee Total FIDC", styles["body"]), Paragraph(_brl(fee_total), styles["body"])],
    ]
    if lance_embutido:
        rows.append([
            Paragraph("Lance Embutido Total (recursos FIDC)", styles["body"]),
            Paragraph(_brl(lance_embutido), styles["body"])
        ])

    col_w = [110*mm, 60*mm]
    t = Table(rows, colWidths=col_w)
    cmds = _base_table_style(1)
    for i in range(1, len(rows)):
        bg = C_LITE_BLUE if i % 2 == 1 else C_WHITE
        cmds.append(("BACKGROUND", (0, i), (-1, i), bg))
    t.setStyle(TableStyle(cmds))
    elems.append(t)
    return elems


def build_fluxo_12m(dados, styles):
    """First 12 months flow table, contemplation rows highlighted in green."""
    elems = []
    elems.append(Spacer(1, 4*mm))
    elems.append(SectionHeader("Fluxo dos Primeiros 12 Meses", CONTENT_WIDTH))
    elems.append(Spacer(1, 3*mm))

    fluxo_data = (dados.get("fluxo") or {}).get("fluxo", [])
    rows12 = fluxo_data[:12]

    if not rows12:
        elems.append(Paragraph("Dados de fluxo não disponíveis.", styles["body"]))
        return elems

    header = [
        Paragraph("<b>Mês</b>", styles["body"]),
        Paragraph("<b>Ctas\nContemp.</b>", styles["body"]),
        Paragraph("<b>Valor Pago</b>", styles["body"]),
        Paragraph("<b>Lance Pago</b>", styles["body"]),
        Paragraph("<b>Créd. Liberado</b>", styles["body"]),
        Paragraph("<b>Créd. Líq. Acum.</b>", styles["body"]),
    ]

    table_rows = [header]
    contemplated_rows = []

    for i, r in enumerate(rows12, 1):
        if r.get("contemplado"):
            contemplated_rows.append(i)
        table_rows.append([
            Paragraph(_fmt_mes(r.get("mes")), styles["body"]),
            Paragraph(_fmt_num(r.get("cotas_contempladas")), styles["body"]),
            Paragraph(_brl(r.get("valor_pago"), ""), styles["body"]),
            Paragraph(_brl(r.get("lance_pago"), ""), styles["body"]),
            Paragraph(_brl(r.get("credito_liberado"), ""), styles["body"]),
            Paragraph(_brl(r.get("credito_liquido_acumulado"), ""), styles["body"]),
        ])

    col_w = [22*mm, 20*mm, 32*mm, 32*mm, 32*mm, 32*mm]
    t = Table(table_rows, colWidths=col_w, repeatRows=1)
    cmds = _base_table_style(1)

    for idx in contemplated_rows:
        cmds.append(("BACKGROUND", (0, idx), (-1, idx), C_GREEN))
        cmds.append(("FONTNAME",   (0, idx), (-1, idx), "Helvetica-Bold"))

    # Align numbers right
    for col in range(1, 6):
        cmds.append(("ALIGN", (col, 1), (col, -1), "RIGHT"))

    t.setStyle(TableStyle(cmds))
    elems.append(t)

    if contemplated_rows:
        elems.append(Spacer(1, 2*mm))
        elems.append(Paragraph(
            "* Meses <font color='#2E7D32'><b>destacados em verde</b></font> indicam contemplação de cotas.",
            styles["body"]
        ))
    return elems


def build_carteira(dados, styles):
    """Portfolio / groups table."""
    elems = []
    elems.append(Spacer(1, 4*mm))
    elems.append(SectionHeader("Carteira de Cotas", CONTENT_WIDTH))
    elems.append(Spacer(1, 3*mm))

    carteira = dados.get("carteira") or {}
    grupos = carteira.get("grupos", [])

    if not grupos:
        elems.append(Paragraph("Dados de carteira não disponíveis.", styles["body"]))
        return elems

    header = [
        Paragraph("<b>Seção</b>", styles["body"]),
        Paragraph("<b>Grupo</b>", styles["body"]),
        Paragraph("<b>Créd. Contratado</b>", styles["body"]),
        Paragraph("<b>Parc. Pré</b>", styles["body"]),
        Paragraph("<b>Prazo</b>", styles["body"]),
        Paragraph("<b>Lance Emb.</b>", styles["body"]),
        Paragraph("<b>Lance Livre</b>", styles["body"]),
        Paragraph("<b>Cotas</b>", styles["body"]),
        Paragraph("<b>Créd. Líq.</b>", styles["body"]),
    ]

    table_rows = [header]
    last_secao = None
    secao_rows = []

    for g in grupos:
        secao = g.get("secao", "")
        if secao != last_secao:
            secao_rows.append(len(table_rows))
            last_secao = secao

        table_rows.append([
            Paragraph(secao, styles["body"]),
            Paragraph(str(g.get("grupo", "")), styles["body"]),
            Paragraph(_brl(g.get("credito_contratado"), ""), styles["body"]),
            Paragraph(_brl(g.get("parcelas_pre"), ""), styles["body"]),
            Paragraph(_fmt_num(g.get("prazo")), styles["body"]),
            Paragraph(_brl(g.get("lance_embutido"), ""), styles["body"]),
            Paragraph(_brl(g.get("lance_livre"), ""), styles["body"]),
            Paragraph(_fmt_num(g.get("qtde_cotas")), styles["body"]),
            Paragraph(_brl(g.get("credito_novo"), ""), styles["body"]),
        ])

    col_w = [23*mm, 14*mm, 23*mm, 18*mm, 12*mm, 22*mm, 22*mm, 11*mm, 25*mm]
    t = Table(table_rows, colWidths=col_w, repeatRows=1)
    cmds = _base_table_style(1)

    # Section separator rows
    for ri in secao_rows:
        cmds.append(("BACKGROUND", (0, ri), (-1, ri), C_LITE_BLUE))
        cmds.append(("FONTNAME",   (0, ri), (-1, ri), "Helvetica-Bold"))
        cmds.append(("TEXTCOLOR",  (0, ri), (-1, ri), C_DARK_BLUE))

    # Alternate rows
    for i in range(1, len(table_rows)):
        if i not in secao_rows:
            bg = C_WHITE if i % 2 == 0 else C_GRAY
            cmds.append(("BACKGROUND", (0, i), (-1, i), bg))

    # Right-align numeric columns
    for col in range(2, 9):
        cmds.append(("ALIGN", (col, 1), (col, -1), "RIGHT"))

    t.setStyle(TableStyle(cmds))
    elems.append(t)

    # Totals summary
    totais = carteira.get("totais", {})
    if totais:
        elems.append(Spacer(1, 3*mm))
        tot_rows = [[
            Paragraph("<b>Seção</b>", styles["body"]),
            Paragraph("<b>Créd. Contratado</b>", styles["body"]),
            Paragraph("<b>Lance Emb.</b>", styles["body"]),
            Paragraph("<b>Lance Livre</b>", styles["body"]),
            Paragraph("<b>Cotas</b>", styles["body"]),
        ]]
        for k, v in totais.items():
            tot_rows.append([
                Paragraph(f"<b>{k}</b>", styles["body"]),
                Paragraph(_brl(v.get("credito_contratado"), ""), styles["body"]),
                Paragraph(_brl(v.get("lance_embutido"), ""), styles["body"]),
                Paragraph(_brl(v.get("lance_livre"), ""), styles["body"]),
                Paragraph(_fmt_num(v.get("qtde_cotas")), styles["body"]),
            ])

        tcol_w = [55*mm, 36*mm, 36*mm, 36*mm, 20*mm]
        tt = Table(tot_rows, colWidths=tcol_w)
        tcmds = _base_table_style(1)
        for i in range(1, len(tot_rows)):
            tcmds.append(("BACKGROUND", (0, i), (-1, i), C_LITE_BLUE if i % 2 else C_WHITE))
        for col in range(1, 5):
            tcmds.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
        tt.setStyle(TableStyle(tcmds))
        elems.append(Paragraph("<b>Resumo por Seção</b>", styles["section"]))
        elems.append(tt)

    return elems


def build_prazos(dados, styles):
    """Prazo breakdown per group."""
    elems = []
    elems.append(Spacer(1, 4*mm))
    elems.append(SectionHeader("Detalhamento de Prazos por Grupo", CONTENT_WIDTH))
    elems.append(Spacer(1, 3*mm))

    carteira = dados.get("carteira") or {}
    grupos = carteira.get("grupos", [])
    prazo_medio = carteira.get("prazo_medio")

    if not grupos:
        elems.append(Paragraph("Dados não disponíveis.", styles["body"]))
        return elems

    header = [
        Paragraph("<b>Grupo</b>", styles["body"]),
        Paragraph("<b>Seção</b>", styles["body"]),
        Paragraph("<b>Prazo (meses)</b>", styles["body"]),
        Paragraph("<b>Parcelas Pré-Contemp.</b>", styles["body"]),
        Paragraph("<b>Crédito Contratado</b>", styles["body"]),
    ]
    rows = [header]
    for g in grupos:
        rows.append([
            Paragraph(str(g.get("grupo", "")), styles["body"]),
            Paragraph(g.get("secao", ""), styles["body"]),
            Paragraph(_fmt_num(g.get("prazo")), styles["body"]),
            Paragraph(_brl(g.get("parcelas_pre"), ""), styles["body"]),
            Paragraph(_brl(g.get("credito_contratado"), ""), styles["body"]),
        ])

    col_w = [22*mm, 35*mm, 35*mm, 45*mm, 40*mm]
    t = Table(rows, colWidths=col_w, repeatRows=1)
    cmds = _base_table_style(1)
    for i in range(1, len(rows)):
        cmds.append(("BACKGROUND", (0, i), (-1, i), C_LITE_BLUE if i % 2 else C_WHITE))
    for col in [2, 3, 4]:
        cmds.append(("ALIGN", (col, 0), (col, -1), "RIGHT"))
    t.setStyle(TableStyle(cmds))
    elems.append(t)

    if prazo_medio:
        elems.append(Spacer(1, 2*mm))
        elems.append(Paragraph(f"<b>Prazo médio da carteira:</b> {_fmt_num(prazo_medio)} meses", styles["body"]))

    return elems


DISCLAIMER_TEXT = (
    "Esta proposta tem caráter ilustrativo e informativo, não constituindo compromisso, "
    "garantia ou promessa de contemplação. As simulações apresentadas são baseadas em médias "
    "históricas de lances livres, fixos, limitados e de fidelidade, bem como na probabilidade "
    "de contemplação por sorteio, e podem variar conforme as condições de mercado e as regras "
    "de cada grupo de consórcio. A TIR (Taxa Interna de Retorno) e demais indicadores financeiros "
    "são estimativas e não representam rendimento garantido. O percentual de lance engloba a média "
    "histórica de lances livres e a probabilidade de contemplação por sorteio, além dos lances "
    "fixo, limitado e fidelidade, quando aplicável. As condições aqui apresentadas estão sujeitas "
    "à aprovação cadastral e às normas vigentes da Ademicon Administradora de Consórcios. "
    "Documentação sujeita a alterações sem aviso prévio."
)


def build_disclaimer(styles):
    elems = []
    elems.append(Spacer(1, 6*mm))
    elems.append(HRFlowable(width=CONTENT_WIDTH, thickness=0.5, color=colors.gray))
    elems.append(Spacer(1, 2*mm))
    elems.append(Paragraph("<b>DISCLAIMER</b>", styles["section"]))
    elems.append(Paragraph(DISCLAIMER_TEXT, styles["disclaimer"]))
    return elems


# ──────────────────────────────────────────────────────────────────────────────
# Footer callback
# ──────────────────────────────────────────────────────────────────────────────
def make_footer_fn(cfg):
    def footer(canvas, doc):
        canvas.saveState()
        w, h = A4
        y_line = 18 * mm
        y_sig  = 11 * mm

        canvas.setStrokeColor(C_MED_BLUE)
        canvas.setLineWidth(0.5)
        canvas.line(MARGIN, y_line, w - MARGIN, y_line)

        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(C_DARK_BLUE)
        canvas.drawString(MARGIN, y_sig + 4, cfg["gerente"])

        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.gray)
        canvas.drawString(MARGIN, y_sig - 4, f"{cfg['cargo']} · {cfg['unidade']}")

        canvas.setFont("Helvetica", 7)
        pag = f"Pág. {doc.page}"
        data = datetime.date.today().strftime("%d/%m/%Y")
        canvas.drawRightString(w - MARGIN, y_sig + 4, pag)
        canvas.drawRightString(w - MARGIN, y_sig - 4, f"Gerado em {data}")

        canvas.restoreState()
    return footer


# ──────────────────────────────────────────────────────────────────────────────
# Full PDF builder
# ──────────────────────────────────────────────────────────────────────────────
def gerar_pdf(cfg, dados):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=26 * mm,
        title="Proposta de Crédito Estruturado – Ademicon",
    )

    styles = make_styles()
    story  = []

    # ── Cover header ──────────────────────────────────────────────────────────
    story.append(CoverHeader(CONTENT_WIDTH, 30 * mm))
    story.append(Spacer(1, 5 * mm))

    # ── Info block ────────────────────────────────────────────────────────────
    story.append(build_info_block(cfg, dados, styles))
    story.append(Spacer(1, 4 * mm))

    # ── Sections (conditional) ────────────────────────────────────────────────
    if cfg.get("sec_resumo"):
        story += build_resumo_executivo(dados, cfg, styles)

    if cfg.get("sec_fidc"):
        story += build_custo_fidc(dados, styles)

    if cfg.get("sec_fluxo"):
        story += build_fluxo_12m(dados, styles)

    if cfg.get("sec_carteira"):
        story += build_carteira(dados, styles)

    if cfg.get("sec_prazos"):
        story += build_prazos(dados, styles)

    if cfg.get("sec_disclaimer"):
        story += build_disclaimer(styles)

    # ── Build ─────────────────────────────────────────────────────────────────
    footer_fn = make_footer_fn(cfg)
    doc.build(story, onFirstPage=footer_fn, onLaterPages=footer_fn)
    buffer.seek(0)
    return buffer.read()


# ══════════════════════════════════════════════════════════════════════════════
# STREAMLIT APP
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Gerador de Proposta PDF · Ademicon",
    page_icon="📄",
    layout="wide",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .block-container { padding-top: 1.5rem; }
  .st-emotion-cache-1629p8f h1 { color: #1F3864; }
  div[data-testid="stSidebar"] { background: #F0F4FB; }
  .metric-card {
    background: #D6E4F0; border-radius: 8px; padding: 12px 16px;
    margin-bottom: 8px;
  }
  .metric-card b { color: #1F3864; }
  .preview-section {
    background: #F5F5F5; border-left: 4px solid #2F75B6;
    padding: 8px 12px; border-radius: 0 6px 6px 0; margin-bottom: 8px;
  }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="background:#1F3864;padding:18px 24px;border-radius:10px;margin-bottom:20px">
  <h2 style="color:white;margin:0">📄 Gerador de Proposta PDF</h2>
  <p style="color:#A8C4E0;margin:4px 0 0">Crédito Estruturado · Consórcio Imobiliário · Ademicon</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Configuração da Proposta")
    st.markdown("---")

    st.markdown("**👤 Identificação**")
    nome_cliente = st.text_input("Nome do Cliente", placeholder="Ex.: Grupo Acme S/A")
    gerente      = st.text_input("Gerente Responsável", value="Julio Cesar Santos")
    cargo        = st.text_input("Cargo", value="Gerente de Crédito Estruturado")
    unidade      = st.text_input("Unidade", value="Ademicon Faria Lima")
    data_ref     = st.date_input("Data de Referência", value=datetime.date.today())

    st.markdown("---")
    st.markdown("**📑 Seções da Proposta**")
    sec_resumo     = st.toggle("Resumo Executivo",           value=True)
    sec_fidc       = st.toggle("Custo FIDC",                 value=True)
    sec_fluxo      = st.toggle("Fluxo — Primeiros 12 Meses", value=True)
    sec_carteira   = st.toggle("Carteira de Cotas",          value=True)
    sec_prazos     = st.toggle("Detalhamento de Prazos",     value=True)
    sec_disclaimer = st.toggle("Disclaimer Padrão Ademicon",  value=True)

    st.markdown("---")
    st.markdown("**📈 Versão da TIR**")
    versao_tir = st.radio("Exibir TIR:", ["Com FIDC", "Sem FIDC", "Ambas"], index=0)

# ── Main area ─────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "📂 Faça upload da planilha de simulação (.xlsx)",
    type=["xlsx"],
    help="Planilha no formato Ademicon com abas RESUMO, CARTEIRA e FLUXO",
)

if not uploaded:
    st.info("⬆️ Faça o upload de uma planilha para começar.")
    st.markdown("""
    **Formato esperado:** planilha `.xlsx` no padrão Ademicon contendo as abas:
    - **RESUMO** — indicadores financeiros e fluxo mensal resumido
    - **CARTEIRA** — grupos de consórcio selecionados
    - **FLUXO** — fluxo detalhado mês a mês das contemplações
    """)
    st.stop()

# ── Read spreadsheet ──────────────────────────────────────────────────────────
with st.spinner("Lendo planilha…"):
    dados = ler_planilha(uploaded.read())

# ── Errors ────────────────────────────────────────────────────────────────────
if dados.get("erros"):
    for err in dados["erros"]:
        st.warning(f"⚠️ {err}")

if not any([dados.get("resumo"), dados.get("carteira"), dados.get("fluxo")]):
    st.error("Não foi possível extrair dados da planilha. Verifique o formato do arquivo.")
    st.stop()

# ── Preview ───────────────────────────────────────────────────────────────────
st.markdown("### 🔍 Preview dos Dados Extraídos")

tabs = st.tabs(["📊 Resumo", "🏦 Carteira", "📈 Fluxo"])

with tabs[0]:
    resumo = dados.get("resumo") or {}
    fluxo_d = dados.get("fluxo") or {}

    c1, c2, c3 = st.columns(3)
    with c1:
        credito = resumo.get("credito_total") or fluxo_d.get("credito_total")
        st.metric("Crédito Total", _brl(credito))
    with c2:
        st.metric("Valor da Parcela", _brl(resumo.get("valor_parcela") or fluxo_d.get("parcela")))
    with c3:
        prazo = resumo.get("qtd_parcelas")
        st.metric("Qtd. Parcelas", _fmt_num(prazo) if prazo else "—")

    c4, c5, c6 = st.columns(3)
    with c4:
        st.metric("TIR Mensal", _pct(resumo.get("tir_mensal") or fluxo_d.get("tir_mensal")))
    with c5:
        st.metric("TIR Anual",  _pct(resumo.get("tir_anual")  or fluxo_d.get("tir_anual")))
    with c6:
        st.metric("Taxa Estática Mensal", _pct(resumo.get("taxa_estatica")))

    fluxo_rows = resumo.get("fluxo", [])
    if fluxo_rows:
        import pandas as pd
        df = pd.DataFrame([{
            "Mês": _fmt_mes(r["mes"]),
            "Parcela (R$)": _brl(r["parcela"], ""),
            "Crédito (R$)": _brl(r["credito"], ""),
            "Crédito Acum. (R$)": _brl(r["credito_acumulado"], ""),
        } for r in fluxo_rows])
        st.markdown("**Fluxo Mensal (RESUMO)**")
        st.dataframe(df, use_container_width=True, hide_index=True)

with tabs[1]:
    import pandas as pd
    carteira = dados.get("carteira") or {}
    grupos = carteira.get("grupos", [])
    if grupos:
        df_cart = pd.DataFrame([{
            "Seção": g["secao"],
            "Grupo": g["grupo"],
            "Créd. Contratado (R$)": _brl(g["credito_contratado"], ""),
            "Parcelas Pré": _brl(g["parcelas_pre"], ""),
            "Prazo (m)": _fmt_num(g["prazo"]),
            "Lance Emb. (R$)": _brl(g["lance_embutido"], ""),
            "Lance Livre (R$)": _brl(g["lance_livre"], ""),
            "Cotas": _fmt_num(g["qtde_cotas"]),
            "Créd. Líq. (R$)": _brl(g["credito_novo"], ""),
        } for g in grupos])
        st.dataframe(df_cart, use_container_width=True, hide_index=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Crédito Total (Carteira)", _brl(carteira.get("credito_total")))
        with c2:
            st.metric("Prazo Médio", f"{_fmt_num(carteira.get('prazo_medio'))} meses"
                      if carteira.get("prazo_medio") else "—")
        with c3:
            st.metric("% Lance Fixo/Limitado", _pct(carteira.get("pct_fixo")))
    else:
        st.info("Nenhum grupo encontrado na aba CARTEIRA.")

with tabs[2]:
    import pandas as pd
    fluxo_d = dados.get("fluxo") or {}

    cx1, cx2, cx3 = st.columns(3)
    with cx1:
        st.metric("Taxa FIDC", _pct(fluxo_d.get("taxa_fidc")))
    with cx2:
        st.metric("TIR Mensal (FIDC)", _pct(fluxo_d.get("tir_mensal")))
    with cx3:
        st.metric("Total de Cotas", _fmt_num(fluxo_d.get("total_cotas")))

    fluxo_rows = fluxo_d.get("fluxo", [])
    if fluxo_rows:
        rows_display = []
        for r in fluxo_rows[:24]:
            rows_display.append({
                "Mês": _fmt_mes(r["mes"]),
                "Cotas Contemp.": _fmt_num(r["cotas_contempladas"]),
                "Valor Pago (R$)": _brl(r["valor_pago"], ""),
                "Lance Pago (R$)": _brl(r["lance_pago"], ""),
                "Créd. Liberado (R$)": _brl(r["credito_liberado"], ""),
                "Créd. Líq. Acum. (R$)": _brl(r["credito_liquido_acumulado"], ""),
                "Contemp.?": "✅" if r["contemplado"] else "",
            })
        df_f = pd.DataFrame(rows_display)
        st.markdown(f"**Aba:** `{dados.get('fluxo_sheet_name', 'FLUXO')}`")

        def highlight_contemplated(row):
            if row["Contemp.?"] == "✅":
                return ["background-color: #E2EFDA"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_f.style.apply(highlight_contemplated, axis=1),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Nenhum dado de fluxo encontrado.")

# ── Generate PDF ──────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("### 🖨️ Gerar Proposta em PDF")

if not nome_cliente.strip():
    st.warning("⚠️ Preencha o **Nome do Cliente** na barra lateral antes de gerar o PDF.")
else:
    cfg = {
        "nome_cliente":    nome_cliente.strip(),
        "gerente":         gerente,
        "cargo":           cargo,
        "unidade":         unidade,
        "data_referencia": data_ref,
        "versao_tir":      versao_tir,
        "sec_resumo":      sec_resumo,
        "sec_fidc":        sec_fidc,
        "sec_fluxo":       sec_fluxo,
        "sec_carteira":    sec_carteira,
        "sec_prazos":      sec_prazos,
        "sec_disclaimer":  sec_disclaimer,
    }

    if st.button("⚡ Gerar PDF", type="primary", use_container_width=True):
        with st.spinner("Gerando PDF…"):
            try:
                pdf_bytes = gerar_pdf(cfg, dados)
                filename = f"Proposta_{nome_cliente.replace(' ', '_')}_{data_ref.strftime('%Y%m%d')}.pdf"
                st.success("✅ PDF gerado com sucesso!")
                st.download_button(
                    label="⬇️ Baixar Proposta PDF",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro ao gerar PDF: {e}")
                raise
