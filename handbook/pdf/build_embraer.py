#!/usr/bin/env python3
"""Собрать учебный PDF «Справочник типов ВС. Embraer»: глава на семейство."""

import json
from pathlib import Path

from weasyprint import HTML

ROOT = Path(__file__).resolve().parents[1]
HERE = Path(__file__).resolve().parent
PHOTO = HERE / "photos" / "embraer"
OUT = ROOT / "embraer.pdf"


def esc(s):
    return (
        str(s)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


MODEL_HEAD = [
    "Модель",
    "Производство",
    "Длина",
    "Размах крыла",
    "Высота",
    "Пассажировместимость",
    "Масса пустого",
    "Макс. масса топлива",
    "Макс. взлётная",
    "Крейсерская скорость",
    "Крейсерская высота",
    "Дальность",
]


# Масса пустого — опубликованная масса пустого снаряжённого. Топливо — масса
# из открытого описания. Прочерк — цифры нет. Годы — от первой поставки
# до последней.
FAMILIES = [
    {
        "id": "ejet",
        "h": "E-Jet",
        "title": "Embraer E-Jet",
        "photo": "ejet",
        "plates": [],
        "models": [
            ["E170", "2004–2018", "29,9 м", "26,0 м", "9,8 м", "66–78", "21,1 т", "9,3 т", "38,6 т", "800 км/ч", "11 000 м", "3 980 км"],
            ["E175", "2005–выпускается", "31,7 м", "28,7 м", "9,9 м", "76–88", "21,9 т", "9,3 т", "40,4 т", "800 км/ч", "11 000 м", "4 070 км"],
            ["E190", "2005–2022", "36,3 м", "28,7 м", "10,6 м", "96–114", "27,8 т", "13,0 т", "51,8 т", "830 км/ч", "11 000 м", "4 540 км"],
            ["E195", "2006–2019", "38,7 м", "28,7 м", "10,5 м", "100–124", "28,7 т", "13,0 т", "52,3 т", "830 км/ч", "11 000 м", "4 260 км"],
        ],
        "fleet": [
            ["E170", "191", "—"],
            ["E175", "840", "0"],
            ["E190", "568", "—"],
            ["E195", "172", "0"],
            ["E170 и E190", "—", "19"],
        ],
        "fleet_total": [1771, 19],
        "fleet_note": "В мире — поставки Embraer на 3-й квартал 2026. У E175 списания единичные, у E170, E190 и E195 часть бортов уже не летает. В РФ 19 — сертификаты 20 крупнейших авиакомпаний на 11 апреля 2025.",
        "engines": [
            ["E170, E175", "General Electric CF34-8E"],
            ["E190, E195", "General Electric CF34-10E"],
        ],
        "apu": [
            ["E170, E175, E190, E195", "Pratt & Whitney APS2300"],
        ],
        "gear": [
            ["E170, E175, носовая стойка", "2 колеса, 24×7.7R10"],
            ["E170, E175, основные стойки", "2 шт. по 2 колеса, H38×13.0R18"],
            ["E190, E195, носовая стойка", "2 колеса, 24×7.7R10"],
            ["E190, E195, основные стойки", "2 шт. по 2 колеса, H41×16.0R20"],
            ["Производитель стоек", "ELEB"],
        ],
    },
    {
        "id": "e2",
        "h": "E-Jet E2",
        "title": "Embraer E-Jet E2",
        "photo": "e2",
        "plates": [],
        "models": [
            ["E175-E2", "не выпускался", "32,4 м", "31,4 м", "10,0 м", "80–90", "—", "8,5 т", "44,6 т", "830 км/ч", "11 000 м", "3 700 км"],
            ["E190-E2", "2018–выпускается", "36,3 м", "33,7 м", "11,0 м", "97–114", "33,0 т", "13,7 т", "56,4 т", "830 км/ч", "11 000 м", "5 460 км"],
            ["E195-E2", "2019–выпускается", "41,5 м", "35,1 м", "10,9 м", "120–146", "35,7 т", "—", "62,5 т", "830 км/ч", "11 000 м", "5 560 км"],
        ],
        "fleet": [
            ["E175-E2", "0", "0"],
            ["E190-E2", "42", "0"],
            ["E195-E2", "182", "0"],
        ],
        "fleet_total": [224, 0],
        "fleet_note": "Поставки Embraer, которые находятся в коммерческой эксплуатации. Срез на октябрь 2026. E175-E2 авиакомпаниям не передавался. В РФ пассажирских E2 нет.",
        "engines": [
            ["E175-E2", "Pratt & Whitney PW1700G"],
            ["E190-E2, E195-E2", "Pratt & Whitney PW1900G"],
        ],
        "apu": [
            ["E175-E2, E190-E2, E195-E2", "Pratt & Whitney APS2600"],
        ],
        "gear": [
            ["E175-E2, носовая стойка", "2 колеса"],
            ["E175-E2, основные стойки", "2 шт. по 2 колеса"],
            ["E190-E2, E195-E2, носовая стойка", "2 колеса, H27×8.5R12"],
            ["E190-E2, E195-E2, основные стойки", "2 шт. по 2 колеса, H42×16.0R20"],
            ["Производитель стоек", "ELEB"],
        ],
    },
]


def html_table(headers, rows, klass):
    head = "".join(f"<th>{esc(h)}</th>" for h in headers)
    body = []
    for row in rows:
        body.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
    return f'<table class="{klass}"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table>'


def fmt_fleet(n):
    if n is None:
        return "—"
    return f"{n:,}".replace(",", " ")


def fleet_table(fam):
    world, rf = fam["fleet_total"]
    head = (
        "<tr><th>Модель</th><th>В мире</th><th>В РФ</th></tr>"
        f'<tr class="total"><th>{esc(fam["h"])}, всего</th>'
        f"<th>{esc(fmt_fleet(world))}</th><th>{esc(fmt_fleet(rf))}</th></tr>"
    )
    body = []
    for row in fam["fleet"]:
        body.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
    return f'<table class="fleet"><thead>{head}</thead><tbody>{"".join(body)}</tbody></table>'


def credit(meta, key):
    item = meta[key]
    artist = item.get("artist") or "автор не указан"
    license_name = item.get("license") or "лицензия см. на странице файла"
    return f"Фото: {artist}. {license_name}. Wikimedia Commons, файл «{item['file']}»."


def family_html(fam, meta):
    img = (PHOTO / f"{fam['photo']}.jpg").resolve().as_uri()
    return f"""
    <section class="family" id="{fam['id']}">
      <div class="sheet photo-only">
        <h2>{esc(fam['title'])}</h2>
        <figure class="shot">
          <img src="{img}" alt="{esc(fam['title'])}">
          <figcaption>{esc(credit(meta, fam['photo']))}</figcaption>
        </figure>
      </div>
      <p class="label">Технические характеристики</p>
      {html_table(MODEL_HEAD, fam["models"], "models")}
      <p class="label">Парк в эксплуатации</p>
      {fleet_table(fam)}
      <p class="platecap">{esc(fam["fleet_note"])}</p>
      <p class="label">Варианты двигателей</p>
      {html_table(["Модель", "Двигатели"], fam["engines"], "specs")}
      <p class="label">Варианты ВСУ</p>
      {html_table(["Серия", "Модель"], fam["apu"], "specs")}
      <p class="label">Шасси</p>
      {html_table(["Стойка", "Колёса"], fam["gear"], "specs")}
    </section>
    """


def toc():
    items = ['<li><a href="#embraer">Embraer</a></li>']
    for fam in FAMILIES:
        items.append(f'<li class="sub"><a href="#{fam["id"]}">{esc(fam["h"])}</a></li>')
    return "<ul class=\"toc\">" + "".join(items) + "</ul>"


CSS = """
@page {
  size: A4;
  margin: 12mm 10mm 14mm 10mm;
  @bottom-center { content: counter(page); font-family: "Liberation Sans", sans-serif; font-size: 9pt; color: #5c6b7a; }
}
html { font-family: "Liberation Serif", serif; color: #1c2b39; font-size: 10.5pt; line-height: 1.35; }
h1 { font-family: "Liberation Sans", sans-serif; font-size: 18pt; color: #16324f; margin: 0 0 6pt; bookmark-level: 1; bookmark-state: open; }
h1.mark { font-size: 0; line-height: 0; margin: 0; color: transparent; }
h2 { font-family: "Liberation Sans", sans-serif; font-size: 15pt; color: #16324f; margin: 0 0 6pt; bookmark-level: 2; bookmark-state: open; }
section.family { break-before: page; }
.sheet.photo-only { break-after: auto; }
.sheet h2 { margin-bottom: 2pt; }
.sheet figure { margin: 1pt 0 2pt; }
p { margin: 0 0 6pt; }
.kicker { font-family: "Liberation Sans", sans-serif; letter-spacing: 0.04em; color: #5c6b7a; font-size: 9pt; margin: 0 0 4pt; }
.cover { break-after: page; }
.lede { font-size: 12pt; }
ul.toc { list-style: none; padding: 0; margin: 8pt 0; }
ul.toc li { margin: 2pt 0; }
ul.toc a { color: #16324f; text-decoration: none; }
ul.toc a::after { content: leader(".") target-counter(attr(href), page); }
ul.toc li.sub { margin-left: 12pt; }
figure { margin: 6pt 0 8pt; }
figcaption { font-family: "Liberation Sans", sans-serif; font-size: 8pt; color: #5c6b7a; margin-top: 2pt; }
.sheet .shot img { width: 100%; height: auto; display: block; }
p.label { font-family: "Liberation Sans", sans-serif; font-weight: 600; font-size: 10pt; color: #16324f; margin: 8pt 0 2pt; break-after: avoid; }
table { width: 100%; border-collapse: collapse; font-family: "Liberation Sans", sans-serif; margin: 0 0 4pt; }
table th, table td { border: 0.4pt solid #c5d0db; padding: 1.6pt 1.6pt; vertical-align: top; text-align: left; }
table th { background: #e7eef5; color: #16324f; font-weight: 600; }
table.models { font-size: 5.7pt; }
table.models th { font-size: 5.4pt; }
table.specs { font-size: 8.5pt; break-inside: avoid; }
table.specs td:first-child, table.specs th:first-child { width: 42%; }
table.fleet { font-size: 8.5pt; }
table.fleet td:nth-child(2), table.fleet td:nth-child(3),
table.fleet th:nth-child(2), table.fleet th:nth-child(3) { text-align: right; width: 18%; }
table.fleet tr.total th { background: #d5e2ef; }
.sources { font-size: 9pt; color: #3d4d5c; }
"""


def build():
    meta = json.loads((PHOTO / "credits.json").read_text(encoding="utf-8"))
    blocks = [family_html(fam, meta) for fam in FAMILIES]
    blocks[0] = blocks[0].replace(
        "<h2>", '<h1 class="mark" id="embraer">Embraer</h1><h2>', 1
    )
    doc = f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>Справочник типов ВС. Embraer</title><style>{CSS}</style></head>
<body>
  <section class="cover">
    <p class="kicker">Учебное пособие · ноябрь 2026</p>
    <h1>Типы воздушных судов</h1>
    <p class="lede">Embraer. Два семейства пассажирских реактивных самолётов, которые сейчас широко летают: E-Jet и E-Jet E2. На первом листе главы фотография на всю ширину.</p>
    <p>Цифры типовые, из открытых описаний производителя. Это не РЛЭ. ERJ-135, ERJ-140, ERJ-145, EMB-120 и бизнес-джеты в этот том не входят.</p>
  </section>
  <section>
    <h1 id="how">Как читать</h1>
    <p>Каждое семейство начинается с нового листа. На этом листе фотография самолёта целиком. Свободно опубликованного чертежа с видами для этих семейств нет, поэтому чертежа на листе нет.</p>
    <p>В таблице длины, размах крыла и высота самолёта даны в метрах, массы в тоннах, скорость и дальность в километрах.</p>
    <p>Одна таблица собирает годы производства и технические характеристики. Годы — от первой поставки до последней. Масса пустого — опубликованная масса пустого снаряжённого для базовой компоновки. Максимальная масса топлива — опубликованная масса заправляемого топлива. Дальность — опубликованная типовая дальность. Прочерк стоит там, где открытое описание эту массу не приводит. «выпускается» значит, что линия на 17 ноября 2026 ещё открыта. «не выпускался» значит, что серийной поставки в авиакомпанию ещё не было. На эту дату так отмечен E175-E2: он летал на испытаниях, но в авиакомпанию не поставлялся.</p>
    <p>Дальше идут варианты двигателей, модели ВСУ и схема шасси. У двигателя и ВСУ указаны производитель и модель. В колонке «Колёса» после числа колёс стоит размер шины. Отдельной строкой указан производитель стойки. Из первого поколения на 17 ноября 2026 серийно выпускается E175. Последняя поставка E170 — 2018 год, E195 — 2019 год, E190 — 2022 год. Эти три модели остаются в эксплуатации.</p>
    <p>Таблица «Парк в эксплуатации» даёт число бортов в мире и у эксплуатантов в РФ. Вторая строка шапки — сумма по этому типу. Ноль значит, что пассажирских бортов в эксплуатации нет. Прочерк значит, что открытый срез не делит цифру по этой модели. Дата среза стоит в подписи под таблицей.</p>
  </section>
  <section>
    <h1 id="toc">Содержание</h1>
    {toc()}
    <p class="sources">Опора: открытые описания Embraer и сводные таблицы опубликованных характеристик. Парк: поставки Embraer на 3-й квартал 2026 и сертификаты 20 крупнейших российских авиакомпаний на 11 апреля 2025. Фотографии — Wikimedia Commons, автор и лицензия под каждым снимком.</p>
  </section>
  {''.join(blocks)}
</body>
</html>"""
    HTML(string=doc, base_url=str(HERE)).write_pdf(OUT)
    print(OUT, OUT.stat().st_size)


if __name__ == "__main__":
    build()
