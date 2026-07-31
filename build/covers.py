# -*- coding: utf-8 -*-
"""Извлечение обложек из Excel-таблицы заказчика.

Картинки вшиты в таблицу и привязаны к строкам. Имя файла = порядковый номер
издания (idx = номер строки - 1), тот же, что использует gen.py в src обложек.
Порядок изданий в таблице меняется — поэтому пересобирать обложки надо целиком,
а не дописывать новые в конец.

Запуск (нужен системный питон, openpyxl стоит только в нём):
    /usr/bin/python3 build/covers.py
"""
import io
import os
import subprocess

import openpyxl

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(BASE, "Книги и журналы -2.xlsx")
DESTS = [os.path.join(BASE, d, "assets", "covers")
         for d in ("rgb", "rgb-light", "docs")]
MAX_SIDE = 700          # обложки в каталоге показываются мелко, больше не нужно
QUALITY = 82

wb = openpyxl.load_workbook(SRC)
ws = wb["Книги"]

for d in DESTS:
    os.makedirs(d, exist_ok=True)

saved = 0
for im in ws._images:
    frm = getattr(im.anchor, "_from", None)
    if frm is None:
        continue
    row = frm.row + 1            # _from.row нумеруется с нуля
    idx = row - 1                # тот же idx, что в gen.py
    name = f"{idx:02d}.jpg"

    data = im._data()
    tmp = os.path.join(DESTS[0], f".tmp-{name}")
    with open(tmp, "wb") as f:
        f.write(data)

    first = os.path.join(DESTS[0], name)
    subprocess.run(
        ["sips", "-Z", str(MAX_SIDE), tmp,
         "-s", "format", "jpeg", "-s", "formatOptions", str(QUALITY),
         "--out", first],
        check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    os.remove(tmp)

    with open(first, "rb") as f:
        blob = f.read()
    for d in DESTS[1:]:
        with open(os.path.join(d, name), "wb") as f:
            f.write(blob)
    saved += 1

print(f"OK: обложек сохранено {saved} -> " + ", ".join(os.path.relpath(d, BASE) for d in DESTS))
