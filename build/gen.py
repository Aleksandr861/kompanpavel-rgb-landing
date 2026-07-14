# -*- coding: utf-8 -*-
"""Генератор лендинга kompanpavel.com/rgb из Excel-таблицы."""
import openpyxl, html, json, re, os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # корень репо
SRC = os.path.join(BASE, "Книги и журналы -2.xlsx")   # исходная таблица (не в репо)
OUT = os.path.join(BASE, "rgb", "index.html")

wb = openpyxl.load_workbook(SRC, data_only=True)
ws = wb["Книги"]

def clean_rec(s):
    s = (s or "").strip()
    s = s.replace("фотонАлексей", "фотограф Алексей")  # опечатка в таблице
    s = re.sub(r"\s+", " ", s)
    return s

ROLES = ["журналист", "продюсер", "фотограф", "визажист", "стилист", "основатель Lebigmag"]

def split_rec(s):
    for role in ROLES:
        if s.lower().startswith(role.lower()):
            return role, s[len(role):].strip()
    return "", s

items = []
for r in range(2, 38):
    name = str(ws.cell(r, 1).value or "").strip()
    if not name:
        continue
    desc = str(ws.cell(r, 2).value or "").strip()
    rec = clean_rec(str(ws.cell(r, 4).value or ""))
    if "\n" in desc:
        first, rest = desc.split("\n", 1)
        title = first.strip().rstrip(".")
        text = re.sub(r"\s*\n\s*", " ", rest).strip()
        kind = "book"
    else:
        title = name
        text = desc
        kind = "mag"
    role, person = split_rec(rec)
    idx = r - 1
    # ПЛЕЙСХОЛДЕРЫ шифров — заменить на реальные из РГБ
    code = f"ЕБШ{10876 + idx}"
    items.append(dict(idx=idx, title=title, text=text, code=code,
                      kind=kind, role=role, person=person))

experts = sorted({(i["person"], i["role"]) for i in items})

def card(i):
    kind_label = "Книга" if i["kind"] == "book" else "Журнал"
    t = html.escape(i["title"])
    return f'''
      <article class="card" data-kind="{i['kind']}" data-person="{html.escape(i['person'])}"
               data-search="{html.escape((i['title'] + ' ' + i['code']).lower())}" id="item-{i['idx']}">
        <div class="card-cover">
          <img src="assets/covers/{i['idx']:02d}.jpg" alt="{t} — обложка" loading="lazy">
          <span class="card-kind">{kind_label}</span>
        </div>
        <div class="card-body">
          <div class="card-code" title="Библиотечный шифр для заказа в РГБ">{i['code']}</div>
          <h3 class="card-title">{t}</h3>
          <p class="card-rec">Рекомендует <b>{html.escape(i['person'])}</b><span class="card-role">{html.escape(i['role'])}</span></p>
          <p class="card-text">{html.escape(i['text'])}</p>
          <button class="card-more" type="button" aria-expanded="false">Читать полностью</button>
        </div>
      </article>'''

cards_html = "\n".join(card(i) for i in items)

# лента обложек в hero: 14 изданий, дубль для бесшовной прокрутки
MARQUEE_PICKS = [1, 2, 3, 4, 5, 10, 12, 14, 20, 22, 24, 25, 28, 35]
by_idx = {i["idx"]: i for i in items}
mq_one = "\n".join(
    f'      <a class="mq-item" href="#item-{n}" title="{html.escape(by_idx[n]["title"])} — открыть в каталоге">'
    f'<img src="assets/covers/{n:02d}.jpg" alt="{html.escape(by_idx[n]["title"])}" decoding="async"></a>'
    for n in MARQUEE_PICKS if n in by_idx
)
marquee_html = mq_one + "\n" + mq_one

persons_options = "\n".join(
    f'          <option value="{html.escape(p)}">{html.escape(p)}</option>'
    for p, _ in experts
)
n_books = sum(1 for i in items if i["kind"] == "book")
n_mags = sum(1 for i in items if i["kind"] == "mag")

with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "template.html"), encoding="utf-8") as f:
    tpl = f.read()

page = (tpl.replace("__CARDS__", cards_html)
           .replace("__MARQUEE__", marquee_html)
           .replace("__PERSON_OPTIONS__", persons_options)
           .replace("__TOTAL__", str(len(items)))
           .replace("__NBOOKS__", str(n_books))
           .replace("__NMAGS__", str(n_mags)))

with open(OUT, "w", encoding="utf-8") as f:
    f.write(page)
print(f"OK: {len(items)} изданий ({n_books} книг, {n_mags} журналов) -> {OUT}")
