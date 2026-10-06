#!/usr/bin/env python3
"""Build an n8n-style canvas using official n8n node icons + exact workflow names."""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

ROOT = Path("/Users/ahmedhamed/Documents/AI Agent/Technical/Graduation_Project/AdaptiveEdu")
ICONS = Path("/tmp/n8n_icons")
OUT = ROOT / "diagrams" / "adaptiveedu_n8n_flow.jpg"

W, H = 2480, 1260
BG = (244, 246, 248)
CARD = (255, 255, 255)
NAVY = (15, 39, 68)
TEAL = (14, 140, 126)
GRAY = (90, 102, 118)
LINE = (180, 190, 200)
TRUE = (46, 140, 90)
FALSE = (184, 92, 28)

# Exact names from AdaptiveEdu_Telegram_Bot.json
NODES = [
    # Exact names from AdaptiveEdu_Telegram_Bot.json
    ("Telegram Trigger", "telegram", 40, 540),
    ("Read message", "code", 310, 540),
    ("IF student ID", "if", 580, 540),
    ("Get students", "sheets", 880, 150),
    ("Get lessons", "sheets", 1160, 150),
    ("Pick next lesson", "code", 1440, 150),
    ("Save open lesson", "sheets", 1720, 150),
    ("Send lesson + quiz", "telegram", 2000, 150),
    ("Get students for quiz", "sheets", 880, 900),
    ("Get lessons for quiz", "sheets", 1160, 900),
    ("Grade quiz", "code", 1440, 900),
    ("Save score", "sheets", 1720, 790),
    ("Append result", "sheets", 1720, 1010),
    ("Send result", "telegram", 2000, 900),
]


def font(size, bold=False):
    for name in (
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
    ):
        p = Path(name)
        if p.exists():
            return ImageFont.truetype(str(p), size)
    return ImageFont.load_default()


def load_icon(key, size=52):
    im = Image.open(ICONS / f"{key}.png").convert("RGBA")
    im.thumbnail((size, size), Image.Resampling.LANCZOS)
    return im


def card(draw, img, name, icon_key, x, y, w=240, h=88):
    draw.rounded_rectangle([x, y, x + w, y + h], radius=16, fill=CARD, outline=(220, 226, 232), width=2)
    ic = load_icon(icon_key, 48)
    img.paste(ic, (x + 16, y + 20), ic)
    # wrap name
    f = font(18, True)
    words = name.split()
    lines, cur = [], ""
    for word in words:
        test = (cur + " " + word).strip()
        if draw.textlength(test, font=f) <= w - 86:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = word
    if cur:
        lines.append(cur)
    ty = y + (h - 22 * len(lines)) // 2
    for i, line in enumerate(lines[:2]):
        draw.text((x + 76, ty + i * 22), line, fill=NAVY, font=f)
    return (x + w // 2, y, x + w // 2, y + h, x, y + h // 2, x + w, y + h // 2)


def arrow(draw, x1, y1, x2, y2, color=LINE, label=""):
    draw.line([(x1, y1), (x2, y2)], fill=color, width=4)
    # simple head
    if abs(x2 - x1) >= abs(y2 - y1):
        hx, hy = x2, y2
        s = 10 if x2 >= x1 else -10
        draw.polygon([(hx, hy), (hx - s, hy - 7), (hx - s, hy + 7)], fill=color)
    else:
        s = 10 if y2 >= y1 else -10
        draw.polygon([(x2, y2), (x2 - 7, y2 - s), (x2 + 7, y2 - s)], fill=color)
    if label:
        mx, my = (x1 + x2) // 2, (y1 + y2) // 2
        f = font(14, True)
        tw = draw.textlength(label, font=f)
        draw.rounded_rectangle([mx - tw / 2 - 8, my - 14, mx + tw / 2 + 8, my + 12], radius=8, fill=(255, 255, 255))
        draw.text((mx - tw / 2, my - 10), label, fill=color, font=f)


def main():
    img = Image.new("RGB", (W, H), BG)
    d = ImageDraw.Draw(img)
    d.text((70, 28), "n8n workflow  ·  AdaptiveEdu — Telegram study bot", fill=NAVY, font=font(32, True))
    d.text(
        (70, 78),
        "Node titles below are the exact names inside AdaptiveEdu_Telegram_Bot.json",
        fill=GRAY,
        font=font(20),
    )

    boxes = {}
    for name, key, x, y in NODES:
        boxes[name] = card(d, img, name, key, x, y)

    def top(n):
        return boxes[n][0], boxes[n][1]

    def bot(n):
        return boxes[n][2], boxes[n][3]

    def left(n):
        return boxes[n][4], boxes[n][5]

    def right(n):
        return boxes[n][6], boxes[n][7]

    arrow(d, *right("Telegram Trigger"), *left("Read message"))
    arrow(d, *right("Read message"), *left("IF student ID"))
    # true branch up
    x1, y1 = boxes["IF student ID"][0], boxes["IF student ID"][1]
    x2, y2 = boxes["Get students"][4], boxes["Get students"][5]
    d.line([(x1, y1), (x1, y2), (x2, y2)], fill=TRUE, width=4)
    d.polygon([(x2, y2), (x2 - 10, y2 - 7), (x2 - 10, y2 + 7)], fill=TRUE)
    f = font(14, True)
    d.text((x1 + 10, (y1 + y2) // 2 - 8), "true  (ID)", fill=TRUE, font=f)
    arrow(d, *right("Get students"), *left("Get lessons"), TRUE)
    arrow(d, *right("Get lessons"), *left("Pick next lesson"), TRUE)
    arrow(d, *right("Pick next lesson"), *left("Save open lesson"), TRUE)
    arrow(d, *right("Save open lesson"), *left("Send lesson + quiz"), TRUE)

    # false branch down
    x1, y1 = boxes["IF student ID"][2], boxes["IF student ID"][3]
    x2, y2 = boxes["Get students for quiz"][4], boxes["Get students for quiz"][5]
    d.line([(x1, y1), (x1, y2), (x2, y2)], fill=FALSE, width=4)
    d.polygon([(x2, y2), (x2 - 10, y2 - 7), (x2 - 10, y2 + 7)], fill=FALSE)
    d.text((x1 + 10, (y1 + y2) // 2 - 8), "false  (quiz)", fill=FALSE, font=f)
    arrow(d, *right("Get students for quiz"), *left("Get lessons for quiz"), FALSE)
    arrow(d, *right("Get lessons for quiz"), *left("Grade quiz"), FALSE)
    # grade splits to save + append then send
    gx, gy = right("Grade quiz")
    sx, sy = left("Save score")
    ax, ay = left("Append result")
    tx, ty = left("Send result")
    d.line([(gx, gy), (gx + 40, gy), (gx + 40, sy), (sx, sy)], fill=FALSE, width=4)
    d.line([(gx + 40, gy), (gx + 40, ay), (ax, ay)], fill=FALSE, width=4)
    d.line([(boxes["Save score"][6], boxes["Save score"][7]), (tx - 20, boxes["Save score"][7]), (tx - 20, ty), (tx, ty)], fill=FALSE, width=4)
    d.line([(boxes["Append result"][6], boxes["Append result"][7]), (tx - 20, boxes["Append result"][7])], fill=FALSE, width=4)
    d.polygon([(tx, ty), (tx - 10, ty - 7), (tx - 10, ty + 7)], fill=FALSE)

    # legend
    d.rounded_rectangle([40, 1160, 2440, 1230], radius=12, fill=CARD)
    tg = load_icon("telegram", 28)
    sh = load_icon("sheets", 28)
    cd = load_icon("code", 28)
    iff = load_icon("if", 28)
    legend = [
        (70, tg, "Telegram  ·  n8n-nodes-base.telegramTrigger / telegram"),
        (720, sh, "Google Sheets  ·  n8n-nodes-base.googleSheets"),
        (1280, cd, "Code  ·  n8n-nodes-base.code"),
        (1780, iff, "If  ·  n8n-nodes-base.if"),
    ]
    for x, ic, label in legend:
        img.paste(ic, (x, 1182), ic)
        d.text((x + 40, 1186), label, fill=NAVY, font=font(18, True))

    img = img.convert("RGB")
    img.save(OUT, "JPEG", quality=92)
    print(OUT)


if __name__ == "__main__":
    main()
