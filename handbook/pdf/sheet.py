"""Общая вёрстка томов справочника: справка о производителе и лист на семейство."""

from pathlib import Path

from weasyprint import HTML

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

CSS = """
@page {
  size: A4;
  margin: 12mm 10mm 14mm 10mm;
  @bottom-center { content: counter(page); font-family: "Liberation Sans", sans-serif; font-size: 9pt; color: #5c6b7a; }
}
html { font-family: "Liberation Serif", serif; color: #1c2b39; font-size: 10.5pt; line-height: 1.35; }
h1 { font-family: "Liberation Sans", sans-serif; font-size: 18pt; color: #16324f; margin: 0 0 8pt; bookmark-level: 1; bookmark-state: open; }
h2 { font-family: "Liberation Sans", sans-serif; font-size: 15pt; color: #16324f; margin: 0 0 4pt; bookmark-level: 2; bookmark-state: open; }
section.intro { break-after: page; }
section.intro p { margin: 0 0 8pt; }
section.family { break-before: page; }
p.familylead { font-size: 10pt; line-height: 1.28; margin: 0 0 3pt; }
.sheet { break-after: page; }
.sheet.photo-only { break-after: auto; }
.sheet h2 { margin-bottom: 2pt; }
.sheet figure { margin: 1pt 0 2pt; }
p { margin: 0 0 6pt; }
figure { margin: 6pt 0 8pt; }
figcaption { font-family: "Liberation Sans", sans-serif; font-size: 8pt; color: #5c6b7a; margin-top: 2pt; }
.sheet .shot img { width: 100%; height: auto; display: block; }
.platecap { font-family: "Liberation Sans", sans-serif; font-size: 8pt; color: #5c6b7a; margin: 0 0 4pt; }
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
"""


def esc(s):
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


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


def family_html(fam, meta, photo_dir):
    img = (photo_dir / f"{fam['photo']}.jpg").resolve().as_uri()
    return f"""
    <section class="family" id="{fam['id']}">
      <div class="sheet photo-only">
        <h2>{esc(fam['title'])}</h2>
        <p class="familylead">{esc(fam['lead'])}</p>
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


def write_book(title, intro_id, paragraphs, families, photo_dir, out):
    import json

    meta = json.loads((photo_dir / "credits.json").read_text(encoding="utf-8"))
    blocks = [family_html(fam, meta, photo_dir) for fam in families]
    body = "".join(f"<p>{esc(p)}</p>" for p in paragraphs)
    doc = f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="utf-8"><title>Справочник типов ВС. {esc(title)}</title><style>{CSS}</style></head>
<body>
  <section class="intro">
    <h1 id="{esc(intro_id)}">{esc(title)}</h1>
    {body}
  </section>
  {''.join(blocks)}
</body>
</html>"""
    HTML(string=doc, base_url=str(photo_dir)).write_pdf(out)
    print(out, out.stat().st_size)
