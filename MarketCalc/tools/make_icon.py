# -*- coding: utf-8 -*-
"""Иконка приложения: сумка с процентом."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "assets"
OUT.mkdir(exist_ok=True)

S = 1024
img = Image.new("RGBA", (S, S), (0, 0, 0, 0))

bg = Image.new("RGB", (S, S))
pixels = bg.load()
for y in range(S):
    k = y / S
    r = int(22 + (10 - 22) * k)
    g = int(26 + (12 - 26) * k)
    b = int(34 + (17 - 34) * k)
    for x in range(S):
        pixels[x, y] = (r, g, b)

mask = Image.new("L", (S, S), 0)
ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), radius=int(S * 0.22), fill=255)
img.paste(bg, (0, 0), mask)
draw = ImageDraw.Draw(img)

yellow = (255, 212, 59, 255)
dark = (18, 20, 26, 255)

# ручка сумки
draw.arc((300, 190, 724, 560), start=180, end=360, fill=yellow, width=54)

# корпус сумки
draw.rounded_rectangle((214, 380, 810, 850), radius=86, fill=yellow)

# блик
draw.rounded_rectangle((260, 430, 300, 720), radius=20, fill=(255, 236, 150, 90))

# процент
font = None
for name in ("arialbd.ttf", "segoeuib.ttf", "Arial Bold.ttf"):
    try:
        font = ImageFont.truetype(name, 380)
        break
    except OSError:
        continue
if font:
    text = "%"
    bbox = draw.textbbox((0, 0), text, font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    draw.text(((S - w) / 2 - bbox[0], 520 - h / 2 - bbox[1] + 60), text, font=font, fill=dark)

img = img.resize((256, 256), Image.LANCZOS)
img.save(OUT / "icon.png")
img.save(OUT / "icon.ico", sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("saved", OUT / "icon.ico")
