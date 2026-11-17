#!/usr/bin/env python3
"""Собрать учебный PDF «Справочник типов ВС. Airbus»: глава на семейство."""

import json
from pathlib import Path

from weasyprint import HTML

ROOT = Path(__file__).resolve().parents[1]
HERE = Path(__file__).resolve().parent
PHOTO = HERE / "photos" / "airbus"
OUT = ROOT / "airbus.pdf"


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


# Масса пустого — опубликованная масса пустого снаряжённого. Топливо — только
# там, где источник даёт массу, а не объём. Прочерк — цифры в открытом
# описании нет. Скорость округлена до 10 км/ч из опубликованного числа Маха.
FAMILIES = [
    {
        "id": "a220",
        "h": "A220",
        "title": "Airbus A220",
        "photo": "a220",
        "plates": [],
        "models": [
            ["A220-100", "2016–выпускается", "35,0 м", "35,1 м", "11,5 м", "100–135", "35,2 т", "—", "63,7 т", "830 км/ч", "12 500 м", "6 700 км"],
            ["A220-300", "2016–выпускается", "38,7 м", "35,1 м", "11,5 м", "120–160", "37,1 т", "—", "70,9 т", "830 км/ч", "12 500 м", "6 300 км"],
        ],
        "engines": [
            ["A220-100, A220-300", "Pratt & Whitney PW1500G"],
        ],
        "apu": [
            ["A220-100, A220-300", "Honeywell GTCP131-9C"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["Основные стойки", "2 шт. по 2 колеса"],
        ],
    },
    {
        "id": "a320",
        "h": "A320",
        "title": "Airbus A320",
        "photo": "a320",
        "plates": [("a32x.png", "150mm")],
        "plate_note": "Размеры на чертеже: первая величина в метрах, после черты футы и дюймы. Слева боковые виды A318, A319, A320 и A321, справа общий вид сверху, сечение и вид спереди. Julien Scavini, CC BY-SA 3.0, файл «Airbus A32X family v1.0.png».",
        "models": [
            ["A318", "2003–2013", "31,4 м", "34,1 м", "12,6 м", "около 110", "39,5 т", "—", "68,0 т", "830 км/ч", "11 900 м", "5 700 км"],
            ["A319", "1996–2021", "33,8 м", "35,8 м", "11,8 м", "124–160", "40,8 т", "—", "75,5 т", "830 км/ч", "11 900 м", "6 940 км"],
            ["A320", "1988–2021", "37,6 м", "35,8 м", "11,8 м", "150–180", "42,6 т", "—", "78,0 т", "830 км/ч", "11 900 м", "6 200 км"],
            ["A321", "1994–2021", "44,5 м", "35,8 м", "11,8 м", "185–220", "48,5 т", "—", "93,5 т", "830 км/ч", "11 900 м", "5 900 км"],
            ["A319neo", "2019–выпускается", "33,8 м", "35,8 м", "11,8 м", "около 140", "—", "—", "75,5 т", "830 км/ч", "11 900 м", "6 850 км"],
            ["A320neo", "2016–выпускается", "37,6 м", "35,8 м", "11,8 м", "150–180", "—", "—", "79,0 т", "830 км/ч", "11 900 м", "6 300 км"],
            ["A321neo", "2017–выпускается", "44,5 м", "35,8 м", "11,8 м", "180–220", "—", "—", "97,0 т", "830 км/ч", "11 900 м", "7 400 км"],
            ["A321XLR", "2024–выпускается", "44,5 м", "35,8 м", "11,8 м", "около 180", "—", "—", "101 т", "830 км/ч", "11 900 м", "8 700 км"],
        ],
        "engines": [
            ["A318", "CFM International CFM56-5B или Pratt & Whitney PW6000"],
            ["A319, A320, A321", "CFM International CFM56-5B или IAE V2500"],
            ["A319neo, A320neo, A321neo, A321XLR", "CFM International LEAP-1A или Pratt & Whitney PW1100G"],
        ],
        "apu": [
            ["A318, A319, A320, A321", "Honeywell GTCP131-9A или Pratt & Whitney APS3200"],
            ["A319neo, A320neo, A321neo, A321XLR", "Honeywell GTCP131-9A"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["Основные стойки", "2 шт. по 2 колеса"],
        ],
    },
    {
        "id": "a330",
        "h": "A330",
        "title": "Airbus A330",
        "photo": "a330",
        "plates": [("a330-pair.png", "80mm")],
        "plate_note": "Размеры на чертеже: первая величина в метрах, после черты футы и дюймы. Слева A330-200, справа A330-300. Julien Scavini, CC BY-SA 3.0, файл «Airbus A330 family v1.1.png». Лист разрезан по пустому полю.",
        "models": [
            ["A330-200", "1998–2019", "58,8 м", "60,3 м", "17,4 м", "около 250", "120,6 т", "109,2 т", "242 т", "870 км/ч", "12 500 м", "13 450 км"],
            ["A330-300", "1994–2020", "63,7 м", "60,3 м", "16,8 м", "около 300", "129,4 т", "109,2 т", "242 т", "870 км/ч", "12 500 м", "11 750 км"],
            ["A330-800", "2020–выпускается", "58,8 м", "64,0 м", "17,4 м", "около 260", "—", "—", "251 т", "870 км/ч", "12 500 м", "15 000 км"],
            ["A330-900", "2018–выпускается", "63,7 м", "64,0 м", "16,8 м", "около 290", "—", "—", "251 т", "870 км/ч", "12 500 м", "13 610 км"],
        ],
        "engines": [
            ["A330-200, A330-300", "General Electric CF6-80E1, Pratt & Whitney PW4000 или Rolls-Royce Trent 700"],
            ["A330-800, A330-900", "Rolls-Royce Trent 7000"],
        ],
        "apu": [
            ["A330-200, A330-300, A330-800, A330-900", "Honeywell GTCP331-350"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["Основные стойки", "2 шт. по 4 колеса"],
        ],
    },
    {
        "id": "a340",
        "h": "A340",
        "title": "Airbus A340",
        "photo": "a340",
        "plates": [("a340-pair.png", "95mm")],
        "plate_note": "Размеры на чертеже: первая величина в метрах, после черты футы и дюймы. Слева A340-200, справа A340-300. Julien Scavini, CC BY-SA 3.0, файл «Airbus A340 family v1.1.png». Лист разрезан по пустому полю.",
        "models": [
            ["A340-200", "1993–1998", "59,4 м", "60,3 м", "17,0 м", "210–250", "118 т", "110,4 т", "275 т", "870 км/ч", "12 000 м", "13 500 км"],
            ["A340-300", "1993–2008", "63,7 м", "60,3 м", "17,0 м", "250–290", "131 т", "110,4 т", "276,5 т", "870 км/ч", "12 000 м", "12 408 км"],
            ["A340-500", "2003–2012", "67,9 м", "63,5 м", "17,5 м", "270–310", "168 т", "175,2 т", "380 т", "870 км/ч", "12 000 м", "16 670 км"],
            ["A340-600", "2002–2010", "75,4 м", "63,5 м", "17,9 м", "320–370", "174 т", "155,5 т", "380 т", "870 км/ч", "12 000 м", "14 450 км"],
        ],
        "engines": [
            ["A340-200, A340-300", "CFM International CFM56-5C"],
            ["A340-500", "Rolls-Royce Trent 553 или Trent 556"],
            ["A340-600", "Rolls-Royce Trent 556 или Trent 560"],
        ],
        "apu": [
            ["A340-200, A340-300", "Honeywell 331-350"],
            ["A340-500, A340-600", "Honeywell 331-600"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["A340-200, A340-300, основные стойки", "3 шт.: 2 по 4 колеса и 1 по 2 колеса"],
            ["A340-500, A340-600, основные стойки", "3 шт. по 4 колеса"],
        ],
    },
    {
        "id": "a350",
        "h": "A350",
        "title": "Airbus A350",
        "photo": "a350",
        "plates": [],
        "models": [
            ["A350-900", "2014–выпускается", "66,8 м", "64,8 м", "17,1 м", "около 325", "139,3 т", "108,3 т", "283 т", "900 км/ч", "12 500 м", "15 750 км"],
            ["A350-1000", "2018–выпускается", "73,8 м", "64,8 м", "17,1 м", "350–410", "—", "127,6 т", "322 т", "900 км/ч", "12 500 м", "16 670 км"],
        ],
        "engines": [
            ["A350-900", "Rolls-Royce Trent XWB-84"],
            ["A350-1000", "Rolls-Royce Trent XWB-97"],
        ],
        "apu": [
            ["A350-900, A350-1000", "Honeywell HGT1700"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["A350-900, основные стойки", "2 шт. по 4 колеса"],
            ["A350-1000, основные стойки", "2 шт. по 6 колёс"],
        ],
    },
    {
        "id": "a380",
        "h": "A380",
        "title": "Airbus A380",
        "photo": "a380",
        "plates": [("a380.png", "140mm")],
        "plate_note": "Размеры на чертеже: первая величина в метрах, после черты футы и дюймы. На листе A380-800. Julien Scavini, CC BY-SA 3.0, файл «Airbus A380-800 v1.0.png».",
        "models": [
            ["A380-800", "2007–2021", "72,7 м", "79,8 м", "24,1 м", "около 525", "285 т", "254,0 т", "575 т", "900 км/ч", "12 500 м", "14 800 км"],
        ],
        "engines": [
            ["A380-800", "Rolls-Royce Trent 900 или Engine Alliance GP7200"],
        ],
        "apu": [
            ["A380-800", "Pratt & Whitney Canada PW980"],
        ],
        "gear": [
            ["Носовая стойка", "2 колеса"],
            ["Основные стойки под крылом", "2 шт. по 4 колеса"],
            ["Основные стойки под фюзеляжем", "2 шт. по 6 колёс"],
        ],
    },
]


def html_table(headers, rows, klass):
    head = "".join(f"<th>{esc(h)}</th>" for h in headers)
    body = []
    for row in rows:
        body.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in row) + "</tr>")
    return f'<table class="{klass}"><thead><tr>{head}</tr></thead><tbody>{"".join(body)}</tbody></table>'


def credit(meta, key):
    item = meta[key]
    artist = item.get("artist") or "автор не указан"
    license_name = item.get("license") or "лицензия см. на странице файла"
    return f"Фото: {artist}. {license_name}. Wikimedia Commons, файл «{item['file']}»."


def family_html(fam, meta):
    img = (PHOTO / f"{fam['photo']}.jpg").resolve().as_uri()
    plates = []
    for name, height in fam["plates"]:
        src = (HERE / "plates" / name).resolve().as_uri()
        plates.append(
            f'<figure class="plate" style="--plate:{height}"><img src="{src}" alt="{esc(fam["title"])}"></figure>'
        )
    note = ""
    if fam.get("plate_note"):
        note = f'<p class="platecap">{esc(fam["plate_note"])}</p>'
    sheet = "sheet" if plates else "sheet photo-only"
    return f"""
    <section class="family" id="{fam['id']}">
      <div class="{sheet}">
        <h2>{esc(fam['title'])}</h2>
        <figure class="shot">
          <img src="{img}" alt="{esc(fam['title'])}">
          <figcaption>{esc(credit(meta, fam['photo']))}</figcaption>
        </figure>
        {''.join(plates)}
        {note}
      </div>
      <p class="label">Технические характеристики</p>
      {html_table(MODEL_HEAD, fam["models"], "models")}
      <p class="label">Варианты двигателей</p>
      {html_table(["Модель", "Двигатели"], fam["engines"], "specs")}
      <p class="label">Варианты ВСУ</p>
      {html_table(["Серия", "Модель"], fam["apu"], "specs")}
      <p class="label">Шасси</p>
      {html_table(fam.get("gear_head") or ["Стойка", "Колёса"], fam["gear"], "specs")}
    </section>
    """


def toc():
    items = ['<li><a href="#airbus">Airbus</a></li>']
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
.sheet { break-after: page; }
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
.sheet .plate { text-align: center; }
.sheet .plate img { width: auto; max-width: 100%; max-height: var(--plate, 100mm); height: auto; }
.platecap { font-family: "Liberation Sans", sans-serif; font-size: 8pt; color: #5c6b7a; margin: 0 0 4pt; }
p.label { font-family: "Liberation Sans", sans-serif; font-weight: 600; font-size: 10pt; color: #16324f; margin: 8pt 0 2pt; break-after: avoid; }
table { width: 100%; border-collapse: collapse; font-family: "Liberation Sans", sans-serif; margin: 0 0 4pt; }
table th, table td { border: 0.4pt solid #c5d0db; padding: 1.6pt 1.6pt; vertical-align: top; text-align: left; }
table th { background: #e7eef5; color: #16324f; font-weight: 600; }
table.models { font-size: 5.7pt; }
table.models th { font-size: 5.4pt; }
table.specs { font-size: 8.5pt; break-inside: avoid; }
table.specs td:first-child, table.specs th:first-child { width: 42%; }
.sources { font-size: 9pt; color: #3d4d5c; }
"""


def build():
    meta = json.loads((PHOTO / "credits.json").read_text(encoding="utf-8"))
    blocks = [family_html(fam, meta) for fam in FAMILIES]
    blocks[0] = blocks[0].replace(
        "<h2>", '<h1 class="mark" id="airbus">Airbus</h1><h2>', 1
    )
    doc = f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>Справочник типов ВС. Airbus</title><style>{CSS}</style></head>
<body>
  <section class="cover">
    <p class="kicker">Учебное пособие · ноябрь 2026</p>
    <h1>Типы воздушных судов</h1>
    <p class="lede">Airbus. Шесть семейств: A220, A320, A330, A340, A350 и A380. На первом листе главы фотография на всю ширину. У A320, A330, A340 и A380 рядом чертёж.</p>
    <p>Цифры типовые, из открытых описаний производителя и сертификационных листов. Это не РЛЭ. A300, A310, грузовые варианты и Beluga в этот том не входят.</p>
  </section>
  <section>
    <h1 id="how">Как читать</h1>
    <p>Каждое семейство начинается с нового листа. На этом листе фотография самолёта целиком. У A320, A330, A340 и A380 рядом чертёж с видами. Чертёж — рисунок автора листа, не заводской чертёж Airbus. У A220 и A350 свободно опубликованного чертежа с видами нет, поэтому на листе только фотография.</p>
    <p>На чертежах A320, A330, A340 и A380 первая величина в метрах, после черты — футы и дюймы. Лист A340 разрезан по пустому полю: слева A340-200, справа A340-300. В таблице длины, размах крыла и высота самолёта даны в метрах, массы в тоннах, скорость и дальность в километрах.</p>
    <p>Одна таблица собирает годы производства и технические характеристики. Масса пустого — опубликованная масса пустого снаряжённого для базовой компоновки. Максимальная масса топлива указана только там, где открытое описание даёт массу, а не объём баков. Дальность — опубликованная типовая дальность. Прочерк стоит там, где открытое описание эту массу не приводит. «выпускается» значит, что линия на 17 ноября 2026 ещё открыта.</p>
    <p>Дальше идут варианты двигателей, модели ВСУ и схема шасси. У двигателя и ВСУ указаны производитель и модель. У A340 три основные стойки. У A380 четыре основные стойки: две под крылом и две под фюзеляжем. У A350-1000 тележка основной стойки шестиколёсная, у A350-900 — четырёхколёсная.</p>
  </section>
  <section>
    <h1 id="toc">Содержание</h1>
    {toc()}
    <p class="sources">Опора: открытые описания Airbus и страницы семейств, сертификационные данные EASA и сводные таблицы опубликованных характеристик. Фотографии и чертежи — Wikimedia Commons, автор и лицензия под каждым снимком и в подписи к чертежу.</p>
  </section>
  {''.join(blocks)}
</body>
</html>"""
    HTML(string=doc, base_url=str(HERE)).write_pdf(OUT)
    print(OUT, OUT.stat().st_size)


if __name__ == "__main__":
    build()
