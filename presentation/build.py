"""Generate the two-slide Russian-language presentation from data.csv."""

import csv
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


ROOT = Path(__file__).parent
with (ROOT / "data.csv").open(encoding="utf-8", newline="") as stream:
    DATA = list(csv.DictReader(stream))

assert len(DATA) == 10 and [int(row["year"]) for row in DATA] == list(range(2015, 2025))

NAVY = RGBColor(12, 24, 44)
PANEL = RGBColor(23, 41, 66)
WHITE = RGBColor(241, 246, 250)
MUTED = RGBColor(163, 183, 202)
GRID = RGBColor(61, 83, 107)
TEAL = RGBColor(87, 219, 207)
ORANGE = RGBColor(255, 174, 122)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


def box(slide, x, y, w, h, fill, rounded=False):
    kind = MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    if rounded:
        shape.adjustments[0] = 0.08
    return shape


def text(slide, value, x, y, w, h, size=14, color=WHITE, bold=False, align=PP_ALIGN.LEFT, link=None):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = 0
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.add_run()
    run.text = value
    run.font.name = "DejaVu Sans"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    if link:
        run.hyperlink.address = link
    return shape


def line(slide, x1, y1, x2, y2, color, width=1.0):
    shape = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2)
    )
    shape.line.color.rgb = color
    shape.line.width = Pt(width)
    return shape


def chart(slide, title, code, y, low, high, color, signed=False):
    x, w, h = 0.59, 12.15, 2.37
    box(slide, x, y, w, h, PANEL, rounded=True)
    text(slide, title, x + 0.23, y + 0.12, 10.8, 0.31, 14, WHITE, True)
    plot_left, plot_right = x + 0.72, x + w - 0.37
    plot_top, plot_bottom = y + 0.82, y + 1.69
    for val in [low, (low + high) / 2, high]:
        yy = plot_bottom - (val - low) / (high - low) * (plot_bottom - plot_top)
        line(slide, plot_left, yy, plot_right, yy, GRID, 0.55)
    if low < 0 < high:
        zero = plot_bottom - (0 - low) / (high - low) * (plot_bottom - plot_top)
        line(slide, plot_left, zero, plot_right, zero, MUTED, 1)
    points = []
    for i, row in enumerate(DATA):
        val = float(row[code])
        xx = plot_left + i * (plot_right - plot_left) / 9
        yy = plot_bottom - (val - low) / (high - low) * (plot_bottom - plot_top)
        points.append((xx, yy, val))
    for (x1, y1, _), (x2, y2, _) in zip(points, points[1:]):
        line(slide, x1, y1, x2, y2, color, 2.3)
    for (xx, yy, val), row in zip(points, DATA):
        dot = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(xx - 0.052), Inches(yy - 0.052), Inches(0.104), Inches(0.104)
        )
        dot.fill.solid()
        dot.fill.fore_color.rgb = color
        dot.line.fill.background()
        label = f"{val:+.1f}" if signed else f"{val:.1f}"
        label_y = yy - 0.39 if val >= 0 else yy + 0.10
        text(slide, label.replace('.', ','), xx - 0.33, label_y, 0.66, 0.20, 9, WHITE, True, PP_ALIGN.CENTER)
        text(slide, row["year"], xx - 0.35, y + 2.05, 0.70, 0.20, 9, MUTED, align=PP_ALIGN.CENTER)


def base(title, n):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY
    box(slide, 0.58, 0.38, 0.08, 0.34, TEAL)
    text(slide, "РОССИЯ  /  2015–2024", 0.80, 0.35, 5.0, 0.26, 10, TEAL, True)
    text(slide, title, 0.65, 0.72, 12.0, 0.55, 25, WHITE, True)
    text(slide, f"{n} / 2", 11.95, 0.33, 0.76, 0.27, 10, MUTED, align=PP_ALIGN.RIGHT)
    return slide


slide = base("Цена кредита и внутренние сбережения", 1)
chart(slide, "Ставка по банковским кредитам частному сектору, % годовых", "lending_rate_pct", 1.43, 5, 19, ORANGE)
chart(slide, "Валовые внутренние сбережения, % ВВП", "gross_domestic_savings_gdp_pct", 3.99, 26, 39, TEAL)
text(slide, "Источник: Всемирный банк / WDI · FR.INR.LEND", 0.68, 6.80, 7.2, 0.23, 9, MUTED,
     link="https://data.worldbank.org/indicator/FR.INR.LEND?locations=RU")
text(slide, "NY.GDS.TOTL.ZS ↗", 9.27, 6.80, 3.34, 0.23, 9, MUTED, align=PP_ALIGN.RIGHT,
     link="https://data.worldbank.org/indicator/NY.GDS.TOTL.ZS?locations=RU")

slide = base("Потребление: доля и реальная динамика", 2)
chart(slide, "Потребление домохозяйств и НКООДХ, % ВВП", "household_consumption_gdp_pct", 1.43, 45, 56, TEAL)
chart(slide, "Реальный рост потребления, % к предыдущему году", "real_household_consumption_growth_pct", 3.99, -14, 15, ORANGE, signed=True)
text(slide, "Источник: Всемирный банк / WDI · NE.CON.PRVT.ZS", 0.68, 6.80, 8.1, 0.23, 9, MUTED,
     link="https://data.worldbank.org/indicator/NE.CON.PRVT.ZS?locations=RU")
text(slide, "NE.CON.PRVT.KD.ZG ↗", 9.22, 6.80, 3.39, 0.23, 9, MUTED, align=PP_ALIGN.RIGHT,
     link="https://data.worldbank.org/indicator/NE.CON.PRVT.KD.ZG?locations=RU")

prs.save(ROOT / "Ставки_сбережения_потребление_Россия_2015-2024.pptx")
