#!/usr/bin/env python3
"""Build editable SVG scene boards, PNG previews, a Word board, and the web board."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import re
from xml.etree import ElementTree as ET

import cairosvg
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parent
SCENES_DIR = ROOT / "scenes"
PREVIEW_DIR = ROOT / "previews"
ASSETS_DIR = ROOT / "assets"
SVG_NS = "http://www.w3.org/2000/svg"
ET.register_namespace("", SVG_NS)
W, H = 1920, 1080

INK = "#263238"
BLUE = "#407BFF"
BLUE_DARK = "#2859BE"
BLUE_PALE = "#E6F0FC"
PAPER = "#F8F7F2"
CORAL = "#ED8067"
MINT = "#B9D8CA"
LILAC = "#DADFFC"
WHITE = "#FFFFFF"
MUTED = "#697985"
PLUM = "#3D2645"
SAGE = "#7FB89A"
TERRACOTTA = "#E4572E"
MUSTARD = "#F2B134"
PEACH = "#FFC4A3"
CREAM = "#FFF6EC"
ESPRESSO = "#2A1F2D"
CASE_TINTS = ["#FFE3D3", "#DCEFE4", "#FCE9C2"]
PALETTE = [PLUM, SAGE, TERRACOTTA, MUSTARD, PEACH, CREAM, ESPRESSO]
PALETTE_NAMES = ["deep plum", "sage", "terracotta", "mustard", "peach", "cream", "espresso"]

ASSETS = {
    "feed": {"file": REPO / "Feed-rafiki.svg", "title": "Feed", "url": "https://storyset.com/illustration/feed/rafiki"},
    "texting": {"file": REPO / "Texting-rafiki.svg", "title": "Texting", "url": "https://storyset.com/illustration/texting/rafiki"},
    "couple": {"file": REPO / "Couple stress-rafiki.svg", "title": "Couple Stress", "url": "https://storyset.com/illustration/couple-stress/rafiki"},
    "problem": {"file": REPO / "Problem solving-rafiki.svg", "title": "Problem Solving", "url": "https://storyset.com/illustration/problem-solving/rafiki"},
    "document": {"file": REPO / "Online document-rafiki.svg", "title": "Online Document", "url": "https://storyset.com/illustration/online-document/rafiki"},
    "research": {"file": REPO / "Research paper-rafiki.svg", "title": "Research Paper", "url": "https://storyset.com/illustration/research-paper/rafiki"},
    "todo": {"file": ASSETS_DIR / "To-do-list-rafiki.svg", "title": "To do list", "url": "https://storyset.com/illustration/to-do-list/rafiki"},
    "social": {"file": REPO / "Social interaction-rafiki.svg", "title": "Social Interaction", "url": "https://storyset.com/illustration/social-interaction/rafiki"},
}

SCENE2_COMMENTS = [
    {
        "id": "comment-01", "handle": "@raraa", "accent": PLUM,
        "lines": [
            [("Pacar gue tiap ada masalah suka ngilang.", INK, False)],
            [("Gue chat malam, pagi baru dibales", INK, False)],
            [("‘ketiduran’. Fix ", INK, False), ("avoidant", BLUE_DARK, True), (" sih.", INK, False)],
        ],
    },
    {
        "id": "comment-02", "handle": "@anakrantau", "accent": MUSTARD,
        "lines": [
            [("Gue kemarin dimarahin temen.", INK, False)],
            [("Padahal dia yang salah, malah marah", INK, False)],
            [("ke gue. ", INK, False), ("Toxic", CORAL, True), (" banget, asli.", INK, False)],
        ],
    },
    {
        "id": "comment-03", "handle": "@halokak", "accent": PEACH,
        "lines": [
            [("Ih kak, itu ", INK, False), ("gaslight", BLUE_DARK, True), (" dah kayaknya.", INK, False)],
            [("Padahal kakaknya cuma cerita soal", INK, False)],
            [("apa yang dia rasain.", INK, False)],
        ],
    },
    {
        "id": "comment-04", "handle": "@ncy", "accent": PLUM,
        "lines": [
            [("Kak, dia murung terus akhir-akhir ini.", INK, False)],
            [("Kayaknya ", INK, False), ("depresi", CORAL, True), (" deh.", INK, False)],
        ],
    },
    {
        "id": "comment-05", "handle": "@vanyaa", "accent": PEACH,
        "lines": [
            [("Baru kenal seminggu udah dikirimin", INK, False)],
            [("hadiah tiap hari. ", INK, False), ("Love bombing", CORAL, True), (" nggak sih?", INK, False)],
        ],
    },
    {
        "id": "comment-06", "handle": "@dindap", "accent": PLUM,
        "lines": [
            [("Gue bilang butuh waktu sendiri, malah", INK, False)],
            [("dibilang nggak punya ", INK, False), ("boundaries.", BLUE_DARK, True), (" Lah?", INK, False)],
        ],
    },
    {
        "id": "comment-07", "handle": "@mikaela", "accent": MUSTARD,
        "lines": [
            [("Ih kak, dia cuma beda pendapat, kok", INK, False)],
            [("langsung dibilang ", INK, False), ("narcissist?", BLUE_DARK, True)],
        ],
    },
    {
        "id": "comment-08", "handle": "@firaaja", "accent": PEACH,
        "lines": [
            [("Nggak bales chat seharian dibilang", INK, False)],
            [
                ("ghosting.", CORAL, True),
                (" Kadang emang lagi sibuk, bestie.", INK, False),
            ],
        ],
    },
    {
        "id": "comment-09", "handle": "@rereaja", "accent": PEACH,
        "lines": [
            [("Baru putus seminggu, dibilang ", INK, False), ("trauma.", BLUE_DARK, True)],
            [("Padahal gue cuma lagi sedih, wkwk.", INK, False)],
        ],
    },
    {
        "id": "comment-10", "handle": "@eka", "accent": PLUM,
        "lines": [
            [("Ih, liburan aja dibilang ", INK, False), ("healing.", CORAL, True)],
            [("Istilahnya dipakai buat semua hal.", INK, False)],
        ],
    },
]

SCENES = [
    {
        "no": 1, "slug": "the-hook", "time": "00:00–00:30", "title": "The Hook",
        "visual": "Live presenter in medium close-up, surrounded by fast social-comment bubbles; the frame then zooms toward a smartphone.",
        "screen": "Are we using psychology terms correctly?",
        "motion": "Komentar masuk cepat mengitari presenter. Presenter menunjuk kamera; garis zoom mengantar ke layar smartphone.",
        "voice": "Guys, pernah nggak sih, cuma karena seseorang nggak membalas chat, kita langsung bilang, ‘Dia avoidant!’ Teman nggak jadi datang, kita sebut toxic. Lagi sedih beberapa hari, langsung merasa depresi. Atau ketika seseorang nggak sesuai ekspektasi, tiba-tiba kita bilang, ‘That's narcissistic!’ Wait. Are we actually understanding psychology, or are we just throwing psychological terms everywhere? Welcome to Therapy Speak Gone Wrong. Let's talk about it!",
        "assets": [], "roles": ["Frame presenter live dan antarmuka smartphone dibuat sebagai vektor SVG yang bisa diedit."],
        "credits": [], "background": "Warm paper, a blue footage frame, and small coral comment accents.",
    },
    {
        "no": 2, "slug": "what-is-therapy-speak", "time": "00:30–01:05", "title": "What Is Therapy Speak?",
        "visual": "Wide cream living-room frame: TV and wall art on the left, a comment feed on an espresso phone at center, and the reclining man on a sage sofa with colored shelves at right.",
        "screen": "02A: intact comments with only the term highlighted. 02B: #Toxic · #Narcissist · #Gaslighting · #Trauma · #Avoidant · #Depressed. 02C: Understanding psychology = Good! · Misusing psychological terms = Problematic.",
        "motion": "02A · 00:30–00:45: comments scroll through the centered phone. 02B · 00:45–00:58: selected terms lift from the feed into a loose, full-screen hashtag stack. 02C · 00:58–01:05: show the two centered on-screen statements.",
        "voice": "Di era social media, istilah psikologi semakin sering kita dengar. Mulai dari toxic relationship, gaslighting, trauma, sampai narcissistic behavior. Fenomena ini sering disebut sebagai therapy speak, yaitu penggunaan istilah yang berasal dari dunia psikologi atau terapi dalam percakapan sehari-hari. Actually, there's nothing wrong with learning psychological terms. Justru, memahami psikologi bisa membantu kita mengenali emosi, memahami perilaku, dan membangun hubungan yang lebih sehat. But here's the problem: ketika istilah-istilah tersebut digunakan tanpa memahami maknanya, kita bisa salah menilai diri sendiri maupun orang lain.",
        "assets": ["feed"], "roles": ["Feed-rafiki.svg: karakter dan sofa dipisahkan dari seluruh perangkat latar; ukurannya disesuaikan dengan komposisi ruang yang lebih lebar.", "Rak buku, bunga, TV, lukisan, dan latar ruang biru adalah vektor editable.", "Scene-02-komponen.svg: enam komentar tambahan sebagai grup SVG terpisah untuk animasi scroll."],
        "credits": ["feed"], "background": "02A: cream room, espresso phone, plum UI details, and restrained warm accents. 02B: cream field with overlapping hashtag cards in terracotta, mustard, and peach. 02C: centered sage and terracotta statements.",
        "frames": [
            {"code": "02A", "slug": "02a-comment-scroll", "title": "Phone Zoom + Comment Scroll", "time": "00:30–00:45"},
            {"code": "02B", "slug": "02b-hashtag-cloud", "title": "Hashtag Cloud Transition", "time": "00:45–00:58"},
            {"code": "02C", "slug": "02c-on-screen-text", "title": "Understanding Psychology · On-Screen Text", "time": "00:58–01:05"},
        ],
    },
    {
        "no": 3, "slug": "normal-emotions-become-labels", "time": "01:05–01:55", "title": "When Normal Emotions Become Labels",
        "visual": "Three separate full-screen examples: an unanswered message, a student slumped over a failed exam with a crossed-out 100 and a red 45 remedial score, and a disagreement. A final frame brings the examples together behind centered on-screen text.",
        "screen": "They're avoidant! · I'm depressed! (100 dicoret → 45 · REMEDI) · You're gaslighting me!\nOne behavior ≠ A psychological diagnosis.",
        "motion": "03A–03C: tiap contoh tampil sendiri; bubble terracotta muncul lebih dulu, lalu koreksi sage. 03B: nilai 100 dicoret, nilai 45 dan tanda REMEDI muncul di atas siswa yang lesu memegang hasil ujian. 03D: ketiga contoh bertemu dan kalimat penutup tampil di tengah.",
        "voice": "Contoh pertama. Teman kita belum membalas chat selama beberapa jam. Lalu kita langsung menyimpulkan bahwa dia punya avoidant attachment. Padahal, mungkin dia sedang sibuk, kelelahan, atau memang belum sempat membalas. One situation doesn't define someone's entire personality. Contoh kedua. Kita merasa sedih setelah gagal mencapai target. Kemudian langsung menyebut diri sendiri mengalami depresi. Padahal, sedih merupakan emosi manusia yang normal. Kesedihan setelah mengalami kegagalan tidak otomatis menunjukkan adanya gangguan depresi. Contoh ketiga. Seseorang tidak setuju dengan pendapat kita, lalu langsung dituduh melakukan gaslighting. Padahal, disagreement atau perbedaan pendapat tidak otomatis termasuk gaslighting. Kita perlu melihat konteks, pola perilaku, dan bagaimana tindakan tersebut terjadi.",
        "assets": ["texting", "couple"], "roles": ["Texting-rafiki.svg: pesan belum dibalas di 03A.", "03B dan thumbnail 03D: siswa lesu serta hasil ujian digambar sebagai vektor SVG editable.", "Couple stress-rafiki.svg: konteks perbedaan pendapat."],
        "credits": ["texting", "couple"], "background": "Case 1 peach tint #FFE3D3, Case 2 sage tint #DCEFE4, and Case 3 mustard tint #FCE9C2; terracotta labels and sage corrections use the PDF's contrast rules.",
        "frames": [
            {"code": "03A", "slug": "03a-case-01-avoidant", "title": "Case 1 · Unanswered Message", "time": "01:05–01:20"},
            {"code": "03B", "slug": "03b-case-02-depressed", "title": "Case 2 · Sadness After Failure", "time": "01:20–01:35"},
            {"code": "03C", "slug": "03c-case-03-gaslighting", "title": "Case 3 · Disagreement", "time": "01:35–01:50"},
            {"code": "03D", "slug": "03d-combined-on-screen-text", "title": "Combined Examples · On-Screen Text", "time": "01:50–01:55"},
        ],
    },
    {
        "no": 4, "slug": "why-it-matters", "time": "01:55–02:35", "title": "Why Does It Matter?",
        "visual": "A person facing a maze anchors a four-step consequence chain.",
        "screen": "Salah memahami diri · Salah menilai orang lain · Konflik hubungan · Stigma gangguan mental",
        "motion": "Kartu label memicu rangkaian empat kartu dampak; urutannya dibaca kiri ke kanan.",
        "voice": "Now, why should we care? First, misusing psychological terms can make us misunderstand our own emotions. Normal feelings like sadness, fear, or disappointment can feel like something is seriously wrong with us. Second, it can create unnecessary conflict. Imagine labeling someone as narcissistic just because they disagree with you. And third, it can reinforce stigma. Mental health conditions are complex experiences, not simply labels we can attach to someone based on one behavior. Remember, psychology is more than just a collection of viral terms. It is a scientific discipline that requires context, evidence, and careful understanding.",
        "assets": ["problem"], "roles": ["Problem solving-rafiki.svg: metafora mencari konteks; empat kartu dampak digambar sebagai vektor."],
        "credits": ["problem"], "background": "Cool pale-blue field with coral domino accents.",
    },
    {
        "no": 5, "slug": "ethics", "time": "02:35–03:20", "title": "The Ethical Side of Psychology",
        "visual": "A digital-document composition supports four concise article cards.",
        "screen": "Pasal 9 · 15 · 28 · 31",
        "motion": "Kartu pasal masuk satu per satu. Pita dasar ilmiah dan kehati-hatian menyatukan keempatnya.",
        "voice": "That's also why ethical responsibility matters when we talk about psychology. Dalam Kode Etik Psikologi Indonesia, ada beberapa prinsip yang relevan dengan fenomena ini. Pasal 9 menekankan bahwa keputusan psikolog dan ilmuwan psikologi harus didasarkan pada pengetahuan ilmiah dan sikap profesional yang telah teruji. Artinya, psychological claims should not be based on assumptions or viral trends alone. Pasal 15 membahas penghindaran dampak buruk. Jadi, informasi psikologi perlu disampaikan dengan mempertimbangkan kemungkinan dampaknya terhadap orang lain. Pasal 28 mengingatkan bahwa pernyataan publik harus disampaikan secara bijaksana, jujur, teliti, dan hati-hati, serta berpedoman pada dasar ilmiah. Sementara itu, Pasal 31 menekankan bahwa pernyataan melalui media harus konsisten dengan kode etik dan didasarkan pada pengetahuan serta praktik psikologi yang tepat. So, sharing psychology is not just about being relatable. It's also about being responsible.",
        "assets": ["document", "research"], "roles": ["Online document-rafiki.svg: halaman digital.", "Research paper-rafiki.svg: konteks dasar ilmiah.", "Empat kartu pasal dibuat sebagai vektor."],
        "credits": ["document", "research"], "background": "Warm white document page on a pale-lilac digital canvas.",
    },
    {
        "no": 6, "slug": "use-therapy-speak-wisely", "time": "03:20–03:50", "title": "How to Use Therapy Speak Wisely",
        "visual": "A checklist illustration sits beside three practical steps and short Indonesian explanations.",
        "screen": "Understand before you label · Describe, don’t diagnose · Learn, don’t self-diagnose",
        "motion": "Tiga baris checklist dan ikon sederhana muncul satu per satu.",
        "voice": "So, how can we use psychological terms more wisely? Here's your three-step reminder. Number one, understand before you label. Cari tahu makna suatu istilah sebelum menggunakannya. Number two, describe behavior, not diagnose people. Daripada langsung menyebut seseorang narcissist, lebih baik jelaskan perilaku spesifik yang membuat kita merasa tidak nyaman. Number three, know your limits. Konten media sosial bisa menjadi sumber edukasi awal, but it cannot replace professional assessment. And most importantly, allow yourself and others to experience normal human emotions without immediately turning them into disorders.",
        "assets": ["todo"], "roles": ["To-do list-rafiki.svg: ilustrasi checklist; tiga langkah dan ikon pendamping disusun sebagai vektor."],
        "credits": ["todo"], "background": "Light mint backdrop with blue checklist rows.",
    },
    {
        "no": 7, "slug": "closing", "time": "03:50–04:00", "title": "Closing",
        "visual": "A group scene sits beside four everyday emotion icons and the closing title card.",
        "screen": "THINK BEFORE YOU LABEL.\nUnderstand psychology. Use it responsibly.",
        "motion": "Istilah memudar; ekspresi dan judul penutup muncul. Sisakan slot identitas pembuat.",
        "voice": "Psychology should help us understand people, not reduce them to labels. So, before you say it, make sure you understand it. Because not every emotion is a disorder, and not every behavior needs a psychological label. Think before you label. Stay curious, stay kind!",
        "assets": ["social"], "roles": ["Social interaction-rafiki.svg: kelompok dengan ekspresi beragam.", "Empat ikon emosi dan end card dibuat sebagai vektor."],
        "credits": ["social"], "background": "Soft peach paper, muted fading labels, and a clear final card.",
    },
]


def tag(name: str) -> str:
    return f"{{{SVG_NS}}}{name}"


def add(parent, name: str, **attrs):
    return ET.SubElement(parent, tag(name), {k.replace("_", "-"): str(v) for k, v in attrs.items()})


def rect(parent, x, y, w, h, fill, rx=0, stroke=None, sw=1, opacity=None):
    attrs = {"x": x, "y": y, "width": w, "height": h, "fill": fill, "rx": rx}
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw})
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "rect", **attrs)


def circle(parent, cx, cy, r, fill, opacity=None, stroke=None, sw=1):
    attrs = {"cx": cx, "cy": cy, "r": r, "fill": fill}
    if opacity is not None:
        attrs["opacity"] = opacity
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw})
    return add(parent, "circle", **attrs)


def line(parent, x1, y1, x2, y2, stroke, sw=2, dash=None, opacity=None):
    attrs = {"x1": x1, "y1": y1, "x2": x2, "y2": y2, "stroke": stroke, "stroke-width": sw, "stroke-linecap": "round"}
    if dash:
        attrs["stroke-dasharray"] = dash
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "line", **attrs)


def path(parent, d, fill="none", stroke=None, sw=2, opacity=None):
    attrs = {"d": d, "fill": fill}
    if stroke:
        attrs.update({"stroke": stroke, "stroke-width": sw, "stroke-linecap": "round", "stroke-linejoin": "round"})
    if opacity is not None:
        attrs["opacity"] = opacity
    return add(parent, "path", **attrs)


def txt(parent, x, y, value, size=24, fill=INK, weight=400, anchor="start", letter=0, opacity=None):
    node = add(parent, "text", x=x, y=y, fill=fill, font_family="Arial, Helvetica, sans-serif", font_size=size,
               font_weight=weight, text_anchor=anchor, letter_spacing=letter)
    if opacity is not None:
        node.set("opacity", str(opacity))
    for i, row in enumerate(str(value).split("\n")):
        span = add(node, "tspan", x=x, dy=0 if i == 0 else size * 1.25)
        span.text = row
    return node


def inline_txt(parent, x, y, pieces, size=18, anchor="start"):
    node = add(parent, "text", x=x, y=y, fill=INK, font_family="Arial, Helvetica, sans-serif", font_size=size,
               font_weight=500, text_anchor=anchor)
    for value, color, bold in pieces:
        color = ESPRESSO if color == INK else TERRACOTTA if color in (BLUE, BLUE_DARK, CORAL) else color
        span = add(node, "tspan", fill=color, font_weight=700 if bold else 500)
        span.text = value
    return node


def draw_social_comment(parent, x, y, w, h, comment_id, handle, lines, accent, size=17):
    group = add(parent, "g", id=comment_id, role="group", aria_label=f"Comment by {handle}")
    rect(group, x, y, w, h, CREAM, rx=17, stroke=PEACH, sw=2)
    circle(group, x + 27, y + 27, 15, accent)
    txt(group, x + 27, y + 32, handle[1:2].upper(), 12, WHITE if accent in (PLUM, TERRACOTTA) else ESPRESSO, 700, "middle")
    txt(group, x + 52, y + 30, handle, 15, ESPRESSO, 700)
    for row, pieces in enumerate(lines):
        inline_txt(group, x + 52, y + 66 + row * 23, pieces, size)
    return group


def pill(parent, x, y, w, h, label, fill=WHITE, color=INK, size=22, stroke=None, weight=700, opacity=None):
    rect(parent, x, y, w, h, fill, rx=h / 2, stroke=stroke, sw=2, opacity=opacity)
    txt(parent, x + w / 2, y + h * 0.66, label, size, color, weight, "middle", opacity=opacity)


def card(parent, x, y, w, h, fill=WHITE, border="#E3E8ED", radius=22):
    rect(parent, x + 7, y + 10, w, h, ESPRESSO, rx=radius, opacity=0.12)
    return rect(parent, x, y, w, h, fill, rx=radius, stroke=border, sw=2)


def embed_asset(parent, asset_key, x, y, w, h, omit_ids=(), include_ids=None, color_map=None,
                view_box=None, sad_expression=False):
    info = ASSETS[asset_key]
    source = ET.parse(info["file"]).getroot()
    vb = view_box or source.get("viewBox", "0 0 500 500")
    nested = add(parent, "svg", x=x, y=y, width=w, height=h, viewBox=vb, preserveAspectRatio="xMidYMid meet", overflow="hidden", aria_label=info["title"] + " illustration, Storyset Rafiki")
    copied = [
        deepcopy(child)
        for child in list(source)
        if (child.get("id") in include_ids if include_ids is not None else child.get("id") not in omit_ids)
    ]
    mapping = {}
    prefix = "s" + re.sub(r"[^a-z0-9]", "", asset_key.lower()) + "-"
    for child in copied:
        for elem in child.iter():
            old = elem.get("id")
            if old:
                mapping[old] = prefix + old
    for child in copied:
        for elem in child.iter():
            if sad_expression:
                d = elem.get("d", "")
                if d.startswith("M255.56,297.55"):
                    elem.set("d", "M253.8,299.2 Q257.1,294.5 260.5,299.2 L260.5,301.1 Q257.1,297.0 253.8,301.1 Z")
                elif d.startswith("M256.38,297.31") or d.startswith("M257.57,300.16"):
                    elem.set("style", "fill:none")
                elif d.startswith("M257.36,282.6"):
                    elem.set("d", "M257.36,282.9 Q259.2,280.8 261.1,282.3")
                elif d.startswith("M251.2,283.29"):
                    elem.set("d", "M247.9,284.8 Q249.8,282.1 252.2,282.7")
            for key, value in list(elem.attrib.items()):
                for old, new in mapping.items():
                    value = value.replace("url(#" + old + ")", "url(#" + new + ")")
                    value = value.replace("#" + old, "#" + new)
                for old, new in (color_map or {}).items():
                    value = re.sub(re.escape(old), new, value, flags=re.IGNORECASE)
                elem.set(key, value)
            if elem.text:
                for old, new in mapping.items():
                    elem.text = elem.text.replace("#" + old, "#" + new)
            if elem.get("id") in mapping:
                elem.set("id", mapping[elem.get("id")])
        nested.append(child)
    return nested


def start_scene(scene, bg=CREAM, accent=PLUM, dark=False):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "scene-title scene-desc"})
    foreground = CREAM if dark else ESPRESSO
    secondary = PEACH if dark else PLUM
    divider = PLUM if dark else PEACH
    title_node = add(root, "title", id="scene-title")
    title_node.text = f"Scene {scene['no']}: {scene['title']}"
    desc_node = add(root, "desc", id="scene-desc")
    desc_node.text = scene["visual"]
    metadata = add(root, "metadata")
    metadata.text = f"Scene {scene['no']} | {scene['title']} | Editable vector storyboard. Existing Storyset art is embedded as SVG; all other visuals are hand-edited vector shapes, not generated images."
    rect(root, 0, 0, W, H, bg)
    # Shared editorial frame and recurring dots keep the seven boards related.
    circle(root, 1790, 165, 165, accent, opacity=0.075)
    circle(root, 120, 900, 110, accent, opacity=0.045)
    circle(root, 1780, 920, 8, accent, opacity=0.55)
    circle(root, 1812, 920, 8, TERRACOTTA, opacity=0.75)
    circle(root, 1844, 920, 8, foreground, opacity=0.60)
    txt(root, 92, 66, "THERAPY SPEAK GONE WRONG", 20, foreground, 700, letter=1.2)
    txt(root, 92, 94, "PRODUCTION BOARD  /  EDITABLE VECTOR", 13, secondary, 700, letter=1.4)
    pill(root, 1575, 43, 250, 48, scene["time"], CREAM, PLUM, 19, stroke=PEACH)
    txt(root, 92, 168, scene["title"], 49 if len(scene["title"]) < 35 else 43, foreground, 700)
    line(root, 92, 195, 1828, 195, divider, 2)
    line(root, 92, 1015, 1828, 1015, divider, 2)
    txt(root, 92, 1045, "VISUAL + TEKS LAYAR + GERAK", 13, secondary, 700, letter=1.1)
    txt(root, 1828, 1045, f"{scene['no']:02d} / 07", 15, foreground, 700, "end")
    return root


def build_scene_1(scene):
    root = start_scene(scene, PLUM, MUSTARD, dark=True)
    # Presenter frame with a neutral editable stand-in, not a generated portrait.
    card(root, 92, 230, 1062, 595, CREAM, PEACH, 24)
    rect(root, 92, 230, 1062, 60, ESPRESSO, rx=24)
    rect(root, 92, 267, 1062, 23, ESPRESSO)
    circle(root, 125, 260, 9, TERRACOTTA)
    circle(root, 155, 260, 9, MUSTARD)
    circle(root, 185, 260, 9, SAGE)
    txt(root, 220, 268, "LIVE FOOTAGE  ·  MEDIUM CLOSE-UP", 17, WHITE, 700, letter=1)
    rect(root, 112, 306, 1022, 492, PEACH, rx=14)
    # Camera frame marks and a simple placeholder silhouette.
    for x, y, sx, sy in [(145, 336, 1, 1), (1100, 336, -1, 1), (145, 770, 1, -1), (1100, 770, -1, -1)]:
        path(root, f"M{x},{y+32*sy} L{x},{y} L{x+32*sx},{y}", stroke=MUSTARD, sw=5)
    circle(root, 620, 463, 69, PEACH)
    path(root, "M551,451 C553,373 690,374 691,457 L678,440 C648,456 613,421 579,451 L562,479 Z", fill=ESPRESSO)
    path(root, "M470,716 C486,614 530,555 620,555 C710,555 755,614 770,716 Z", fill=PLUM)
    path(root, "M538,589 C574,612 667,612 702,589 L725,716 L516,716 Z", fill=PEACH, opacity=0.45)
    rect(root, 130, 741, 986, 42, CREAM, rx=12, opacity=0.9)
    txt(root, 623, 770, "PLACE PRESENTER VIDEO HERE", 17, ESPRESSO, 700, "middle", letter=1)
    # Floating labels arrive from all sides; the words remain clearly framed as quoted reactions.
    pill(root, 154, 201, 230, 54, "“avoidant?”", MUSTARD, ESPRESSO, 23, stroke=CREAM)
    pill(root, 748, 210, 180, 54, "“toxic”", TERRACOTTA, WHITE, 23, stroke=CREAM)
    pill(root, 234, 795, 235, 52, "“depressed?”", PEACH, ESPRESSO, 22, stroke=CREAM)
    pill(root, 790, 794, 260, 52, "“narcissistic!”", CREAM, PLUM, 21, stroke=PEACH)
    # Smartphone transition is a custom vector object.
    path(root, "M1175,500 C1218,500 1217,500 1252,500", stroke=PEACH, sw=8)
    path(root, "M1240,484 L1262,500 L1240,516 Z", fill=PEACH)
    card(root, 1274, 276, 352, 500, ESPRESSO, ESPRESSO, 34)
    rect(root, 1290, 302, 320, 438, CREAM, rx=22)
    rect(root, 1412, 286, 76, 8, PLUM, rx=4)
    txt(root, 1322, 346, "SOCIAL FEED", 16, PLUM, 700, letter=1)
    for i, (word, col) in enumerate([("avoidant", PLUM), ("toxic", TERRACOTTA), ("sad?", MUSTARD)]):
        y = 374 + i * 96
        rect(root, 1318, y, 264, 76, CREAM, rx=16, stroke=PEACH)
        circle(root, 1344, y + 28, 11, col)
        rect(root, 1370, y + 19, 172, 12, PEACH, rx=6)
        txt(root, 1370, y + 53, word, 18, ESPRESSO, 700)
    txt(root, 1450, 812, "ZOOM → SMARTPHONE", 16, MUSTARD, 700, "middle", letter=0.8)
    rect(root, 92, 876, 1736, 93, PLUM, rx=20, stroke=PEACH)
    txt(root, 132, 932, "ON-SCREEN", 14, CREAM, 700, letter=1.2)
    screen_text = add(root, "text", x=1320, y=936, fill=CREAM, font_family="Arial, Helvetica, sans-serif", font_size=27, font_weight=700, text_anchor="middle")
    for value, color in [("Are we using psychology ", CREAM), ("terms correctly?", MUSTARD)]:
        span = add(screen_text, "tspan", fill=color)
        span.text = value
    return root


def build_scene_2(scene):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "scene-title scene-desc"})
    title_node = add(root, "title", id="scene-title")
    title_node.text = "Scene 2A: room, phone, and scrolling comments"
    desc_node = add(root, "desc", id="scene-desc")
    desc_node.text = "A cream living room with a TV and framed art at left, an espresso phone with a scrolling comment feed in the center, and a reclining man on a sage sofa beneath colored shelves at right."
    defs = add(root, "defs")
    tv_gradient = add(defs, "linearGradient", id="tv-screen", x1="0%", y1="0%", x2="100%", y2="100%")
    add(tv_gradient, "stop", offset="0%", stop_color=PLUM)
    add(tv_gradient, "stop", offset="100%", stop_color=ESPRESSO)
    rect(root, 0, 0, W, H, CREAM)
    rect(root, 0, 862, W, 218, PEACH, opacity=0.38)
    line(root, 0, 862, W, 862, PEACH, 3, opacity=0.8)
    circle(root, 175, 170, 185, MUSTARD, opacity=0.10)
    circle(root, 1040, 845, 280, WHITE, opacity=0.34)
    circle(root, 1860, 540, 205, SAGE, opacity=0.10)

    # Left wall art and television form a simple, colored living-room setting.
    rect(root, 115, 130, 358, 226, ESPRESSO, rx=17)
    rect(root, 132, 147, 324, 192, PEACH, rx=8)
    circle(root, 370, 196, 32, MUSTARD)
    path(root, "M145,307 L242,218 L310,282 L365,238 L447,314 L447,339 L145,339 Z", fill=SAGE)
    path(root, "M145,319 L225,270 L281,315 L350,266 L447,326 L447,339 L145,339 Z", fill=PLUM)
    path(root, "M145,322 L216,293 L276,324 L345,291 L447,329 L447,339 L145,339 Z", fill=TERRACOTTA)

    rect(root, 67, 425, 460, 307, ESPRESSO, rx=24)
    rect(root, 83, 441, 428, 270, "url(#tv-screen)", rx=13)
    circle(root, 494, 454, 3, MUSTARD, opacity=0.8)
    path(root, "M99,675 C215,591 284,614 371,527 C427,471 485,493 510,462 L510,711 L99,711 Z", fill=PLUM, opacity=0.55)
    path(root, "M250,732 L385,732 L405,757 L230,757 Z", fill=ESPRESSO)
    rect(root, 29, 758, 582, 91, PLUM, rx=13)
    rect(root, 53, 778, 534, 9, PEACH, rx=4)
    rect(root, 64, 796, 154, 39, MUSTARD, rx=7)
    rect(root, 239, 796, 154, 39, SAGE, rx=7)
    rect(root, 414, 796, 154, 39, TERRACOTTA, rx=7)
    rect(root, 72, 849, 18, 26, ESPRESSO, rx=4)
    rect(root, 548, 849, 18, 26, ESPRESSO, rx=4)
    # A small plant and a few colored books keep the console from reading as a UI panel.
    rect(root, 540, 708, 60, 48, TERRACOTTA, rx=8)
    path(root, "M570,709 C557,687 555,672 563,656 C577,674 579,690 570,709 Z", fill=SAGE)
    path(root, "M568,708 C582,687 595,679 607,682 C600,700 587,710 568,708 Z", fill=SAGE)
    path(root, "M570,707 C560,689 543,683 533,689 C541,704 553,711 570,707 Z", fill=SAGE)
    rect(root, 229, 726, 20, 31, TERRACOTTA, rx=3)
    rect(root, 253, 717, 23, 40, MUSTARD, rx=3)
    rect(root, 280, 723, 22, 34, PLUM, rx=3)

    # Right-side shelves: colored books, a vase, and flowers sit above the original character.
    rect(root, 1287, 250, 538, 20, PLUM, rx=7)
    rect(root, 1314, 382, 500, 19, PLUM, rx=7)
    rect(root, 1326, 213, 10, 38, ESPRESSO, rx=4)
    rect(root, 1777, 213, 10, 38, ESPRESSO, rx=4)
    rect(root, 1345, 168, 26, 81, TERRACOTTA, rx=5)
    rect(root, 1376, 181, 23, 68, MUSTARD, rx=5)
    rect(root, 1405, 157, 31, 92, SAGE, rx=5)
    rect(root, 1442, 176, 22, 73, PEACH, rx=5)
    rect(root, 1470, 189, 27, 60, ESPRESSO, rx=5)
    rect(root, 1714, 207, 63, 42, MUSTARD, rx=8)
    path(root, "M1745,207 C1736,183 1739,162 1751,144 C1765,166 1764,187 1745,207 Z", fill=SAGE)
    path(root, "M1749,207 C1762,183 1780,174 1793,181 C1781,199 1767,207 1749,207 Z", fill=SAGE)
    circle(root, 1753, 145, 9, TERRACOTTA)
    circle(root, 1745, 151, 8, MUSTARD)
    circle(root, 1761, 154, 8, PEACH)
    circle(root, 1754, 158, 7, CREAM)
    rect(root, 1342, 342, 20, 38, SAGE, rx=4)
    rect(root, 1367, 324, 24, 56, TERRACOTTA, rx=4)
    rect(root, 1396, 335, 28, 45, PEACH, rx=4)
    rect(root, 1430, 314, 22, 66, MUSTARD, rx=4)
    rect(root, 1460, 330, 24, 50, SAGE, rx=4)
    rect(root, 1705, 336, 69, 43, PLUM, rx=9)
    path(root, "M1739,336 C1729,316 1730,299 1740,282 C1756,304 1754,322 1739,336 Z", fill=SAGE)
    path(root, "M1741,335 C1753,316 1770,310 1781,316 C1771,332 1759,338 1741,335 Z", fill=SAGE)

    # Keep the Storyset man and sofa at the same display size; widen the room around them.
    add(root, "ellipse", cx=1630, cy=866, rx=346, ry=47, fill=ESPRESSO, opacity=0.15)
    shared_map = {"#263238": ESPRESSO, "#B97964": PEACH, "#A24E3F": TERRACOTTA,
                  "#E0E0E0": CREAM, "#E6E6E6": PEACH, "#EBEBEB": CREAM,
                  "#F0F0F0": CREAM, "#F5F5F5": CREAM}
    embed_asset(root, "feed", 1080, 150, 850, 850,
                include_ids={"freepik--Sofa--inject-2"}, color_map={**shared_map, "#407BFF": SAGE})
    embed_asset(root, "feed", 1080, 150, 850, 850,
                include_ids={"freepik--Character--inject-2"}, color_map={**shared_map, "#407BFF": PLUM})

    # The centered phone is the only place where the scrollable social feed appears.
    rect(root, 664, 79, 622, 934, ESPRESSO, rx=58, opacity=0.22)
    rect(root, 646, 61, 622, 934, ESPRESSO, rx=58)
    rect(root, 666, 81, 582, 894, CREAM, rx=42)
    rect(root, 889, 94, 136, 11, PLUM, rx=6)
    circle(root, 1172, 129, 5, PLUM)
    circle(root, 1187, 129, 5, PLUM)
    rect(root, 1201, 121, 23, 15, "none", rx=3, stroke=PLUM, sw=2)
    rect(root, 1206, 125, 13, 7, SAGE, rx=2)
    txt(root, 704, 179, "Beranda", 23, ESPRESSO, 700)
    circle(root, 1200, 167, 9, "none", stroke=PLUM, sw=2)
    path(root, "M1207,174 L1214,181", stroke=PLUM, sw=2)
    line(root, 686, 193, 1228, 193, PEACH, 2)
    for i, comment in enumerate(SCENE2_COMMENTS[:4]):
        draw_social_comment(root, 686, 209 + i * 166, 542, 153, comment["id"], comment["handle"], comment["lines"], comment["accent"], size=18)
    # A narrow track cues the scroll animation without adding a storyboard caption.
    rect(root, 1236, 218, 4, 624, PEACH, rx=2)
    rect(root, 1232, 310, 12, 154, PLUM, rx=6)
    line(root, 686, 889, 1228, 889, PEACH, 2)
    circle(root, 760, 926, 9, PLUM)
    circle(root, 905, 926, 9, MUSTARD)
    circle(root, 1050, 926, 9, MUSTARD)
    circle(root, 1195, 926, 9, MUSTARD)
    return root


def build_scene_2_components():
    root = ET.Element(tag("svg"), {"width": "560", "height": "1080", "viewBox": "0 0 560 1080", "role": "img", "aria-labelledby": "components-title components-desc"})
    title_node = add(root, "title", id="components-title")
    title_node.text = "Scene 2 scroll components: six additional social comments"
    desc_node = add(root, "desc", id="components-desc")
    desc_node.text = "Six separate editable comment groups, comment-05 through comment-10, ready to move through the phone during the scroll animation."
    for i, comment in enumerate(SCENE2_COMMENTS[4:]):
        draw_social_comment(root, 20, 14 + i * 176, 520, 153, comment["id"], comment["handle"], comment["lines"], comment["accent"], size=17)
    return root


def build_scene_2b(scene):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "scene-title scene-desc"})
    title_node = add(root, "title", id="scene-title")
    title_node.text = "Scene 2B: full-screen hashtag cloud"
    desc_node = add(root, "desc", id="scene-desc")
    desc_node.text = "A loose stack of overlapping hashtag word cards in terracotta, mustard, and peach fills a cream frame."
    rect(root, 0, 0, W, H, CREAM)
    circle(root, 140, 230, 200, MUSTARD, opacity=0.10)
    circle(root, 1780, 870, 245, PEACH, opacity=0.20)
    # Each hashtag has its own colored backing; offsets and slight rotations make a layered word cloud.
    txt(root, 165, 384, "#", 88, PLUM, 700, "middle", opacity=0.14)
    txt(root, 1760, 405, "#", 92, TERRACOTTA, 700, "middle", opacity=0.18)
    txt(root, 170, 798, "#", 78, MUSTARD, 700, "middle", opacity=0.24)
    txt(root, 1752, 788, "#", 84, PLUM, 700, "middle", opacity=0.15)
    tags = [
        ("#psychologytok", 100, 258, 245, 88, BLUE_DARK, WHITE, -3, 22),
        ("#toxicrelationship", 355, 258, 335, 88, CORAL, WHITE, 3, 23),
        ("#gaslighting", 700, 258, 270, 88, BLUE, WHITE, -2, 26),
        ("#trauma", 980, 258, 180, 88, WHITE, INK, 4, 27),
        ("#avoidant", 1170, 258, 215, 88, LILAC, INK, -3, 24),
        ("#redflag", 1395, 258, 185, 88, INK, WHITE, 3, 24),
        ("#depressed", 1590, 258, 220, 88, MINT, INK, -3, 23),
        ("#narcissist", 120, 342, 250, 88, BLUE_DARK, WHITE, 4, 24),
        ("#healing", 380, 342, 185, 88, MINT, INK, -4, 24),
        ("#therapy", 575, 342, 180, 88, WHITE, BLUE_DARK, 3, 24),
        ("#ghosting", 765, 342, 200, 88, CORAL, WHITE, -4, 24),
        ("#innerchild", 975, 342, 220, 88, LILAC, INK, 3, 22),
        ("#boundaries", 1205, 342, 245, 88, BLUE, WHITE, -2, 23),
        ("#selfdiagnose", 1460, 342, 250, 88, WHITE, BLUE_DARK, 4, 22),
        ("#triggered", 1720, 342, 170, 88, CORAL, WHITE, -3, 21),
        ("#attachmentstyle", 105, 426, 300, 88, LILAC, INK, -2, 21),
        ("#situationship", 415, 426, 245, 88, BLUE, WHITE, 3, 22),
        ("#lovebombing", 670, 426, 220, 88, WHITE, CORAL, -3, 22),
        ("#emotionalabuse", 900, 426, 275, 88, CORAL, WHITE, 3, 21),
        ("#mentalhealth", 1185, 426, 250, 88, BLUE_DARK, WHITE, -3, 22),
        ("#relationshipgoals", 1445, 426, 350, 88, MINT, INK, 3, 20),
        ("#toxic", 95, 510, 170, 88, CORAL, WHITE, 3, 26),
        ("#narcissistic", 275, 510, 245, 88, WHITE, INK, -3, 21),
        ("#traumaresponse", 530, 510, 280, 88, BLUE_DARK, WHITE, 3, 21),
        ("#selflove", 820, 510, 190, 88, MINT, INK, -2, 23),
        ("#healthyrelationships", 1020, 510, 340, 88, LILAC, INK, 3, 19),
        ("#viralterms", 1370, 510, 200, 88, BLUE, WHITE, -3, 22),
        ("#context", 1580, 510, 205, 88, WHITE, BLUE_DARK, 3, 23),
        ("#redflag", 110, 594, 190, 88, INK, WHITE, -3, 24),
        ("#healingjourney", 310, 594, 290, 88, WHITE, BLUE_DARK, 3, 21),
        ("#mentalhealth", 610, 594, 255, 88, BLUE, WHITE, -3, 22),
        ("#avoidant", 875, 594, 210, 88, MINT, INK, 4, 24),
        ("#depressed", 1095, 594, 210, 88, CORAL, WHITE, -3, 23),
        ("#innerchild", 1315, 594, 225, 88, LILAC, INK, 3, 22),
        ("#selfdiagnose", 1550, 594, 250, 88, BLUE_DARK, WHITE, -3, 21),
        ("#therapy", 105, 678, 185, 88, BLUE, WHITE, -3, 24),
        ("#boundaries", 300, 678, 240, 88, CORAL, WHITE, 3, 22),
        ("#ghosting", 550, 678, 205, 88, WHITE, BLUE_DARK, -3, 23),
        ("#attachmentstyle", 765, 678, 300, 88, MINT, INK, 3, 21),
        ("#gaslighting", 1075, 678, 270, 88, BLUE_DARK, WHITE, -3, 24),
        ("#psychologytok", 1355, 678, 255, 88, INK, WHITE, 3, 21),
        ("#selflove", 1620, 678, 210, 88, LILAC, INK, -3, 23),
        ("#toxicrelationship", 120, 762, 330, 88, CORAL, WHITE, 3, 21),
        ("#situationship", 460, 762, 245, 88, BLUE, WHITE, -3, 22),
        ("#lovebombing", 715, 762, 230, 88, WHITE, CORAL, 3, 22),
        ("#narcissist", 955, 762, 250, 88, BLUE_DARK, WHITE, -3, 23),
        ("#redflag", 1215, 762, 185, 88, MINT, INK, 3, 24),
        ("#healing", 1410, 762, 180, 88, WHITE, BLUE_DARK, -3, 23),
        ("#mentalhealth", 1600, 762, 250, 88, INK, WHITE, 3, 21),
    ]
    backplates = [TERRACOTTA, MUSTARD, PEACH]
    for index, (word, x, y, w, h, fill, color, angle, size) in enumerate(tags):
        y += ((index * 7) % 13) - 6
        y = 72 + (y - 258) * 1.6
        fill = backplates[index % len(backplates)]
        color = WHITE if fill == TERRACOTTA else ESPRESSO
        group = add(root, "g", transform=f"rotate({angle} {x + w / 2} {y + h / 2})")
        rect(group, x + 5, y + 6, w, h, ESPRESSO, rx=17, opacity=0.12)
        rect(group, x, y, w, h, fill, rx=17, stroke=CREAM, sw=2)
        txt(group, x + w / 2, y + h * 0.65, word, size, color, 700, "middle")
    return root


def build_scene_2c(scene):
    root = clean_frame(
        "Scene 2C: understanding psychology and misusing terms",
        "Centered sage pill: Understanding psychology = Good! Centered terracotta pill: Misusing psychological terms = Problematic.",
        CREAM,
    )
    rect(root, 385, 304, 1150, 148, SAGE, rx=30)
    txt(root, 960, 398, "Understanding psychology = Good!", 48, ESPRESSO, 700, "middle")
    rect(root, 270, 552, 1380, 176, TERRACOTTA, rx=30)
    txt(root, 960, 660, "Misusing psychological terms = Problematic.", 46, WHITE, 700, "middle")
    return root


def clean_frame(title, description, background):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "frame-title frame-desc"})
    title_node = add(root, "title", id="frame-title")
    title_node.text = title
    desc_node = add(root, "desc", id="frame-desc")
    desc_node.text = description
    metadata = add(root, "metadata")
    metadata.text = "Editable animation frame. Existing Storyset vector groups are recolored in SVG; all added elements are hand-edited vector shapes. No generated imagery."
    rect(root, 0, 0, W, H, background)
    return root


SCENE3_CASES = [
    {
        "key": "texting", "background": CASE_TINTS[0], "quote": "They're avoidant!",
        "correction": ["Mungkin sibuk, lelah, atau", "belum sempat."],
        "colors": {"#2E353A": ESPRESSO, "#407BFF": PLUM, "#DE5753": TERRACOTTA, "#E4897B": PEACH,
                   "#E0E0E0": CREAM, "#E9E9E9": CREAM, "#EBEBEB": PEACH, "#F4F4F4": CREAM, "#F5F5F5": CASE_TINTS[0]},
    },
    {
        "key": "feeling", "background": CASE_TINTS[1], "quote": "I'm depressed!",
        "correction": ["Sedih setelah gagal itu", "emosi yang normal."],
    },
    {
        "key": "couple", "background": CASE_TINTS[2], "quote": "You're gaslighting me!",
        "correction": ["Beda pendapat belum tentu", "gaslighting."],
        "colors": {"#263238": ESPRESSO, "#407BFF": PLUM, "#D58745": MUSTARD, "#D9877C": PEACH,
                   "#DE5753": TERRACOTTA, "#EBB376": PEACH, "#E0E0E0": CREAM, "#EBEBEB": CREAM,
                   "#F0F0F0": CASE_TINTS[2], "#F5F5F5": CREAM, "#FAFAFA": CASE_TINTS[2]},
    },
]


def draw_case_frame(case_index):
    case = SCENE3_CASES[case_index]
    number = case_index + 1
    description = (
        "A slumped student at a desk holds an exam paper. Above him, a dark 100 is crossed out in terracotta and replaced by a red 45 marked REMEDI. The terracotta label says: I'm depressed! The sage correction says: Sedih setelah gagal itu emosi yang normal."
        if case["key"] == "feeling"
        else f"{case['quote']} A terracotta assumption appears beside the illustration, then a sage correction reads: {' '.join(case['correction'])}"
    )
    root = clean_frame(
        f"Scene 3 case {number}: {case['quote']}",
        description,
        case["background"],
    )
    if case["key"] == "feeling":
        root.find(tag("metadata")).text = "Hand-drawn, editable SVG vectors for the student, desk, and exam sheet. No generated imagery."
    color_map = case.get("colors", {})
    omit_bg = {f"freepik--background-complete--inject-2"}
    if case["key"] == "texting":
        embed_asset(root, "texting", 58, 78, 900, 900,
                    omit_ids=omit_bg | {"freepik--Chat--inject-2", "freepik--character-1--inject-2"},
                    color_map=color_map, sad_expression=True)
        # The recipient avatar belongs in the chat header; both outgoing messages share one side.
        rect(root, 245, 142, 570, 82, CREAM, rx=24, stroke=PEACH, sw=3)
        embed_asset(root, "texting", 260, 153, 58, 58,
                    include_ids={"freepik--character-1--inject-2"}, color_map=color_map,
                    view_box="75 111 58 64")
        txt(root, 338, 193, "Teman", 23, ESPRESSO, 700)
        rect(root, 335, 244, 480, 82, PLUM, rx=24)
        txt(root, 575, 296, "Hei, ada waktu?", 26, WHITE, 700, "middle")
        rect(root, 255, 344, 560, 88, PLUM, rx=24)
        txt(root, 535, 400, "Balas dong :( ", 26, WHITE, 700, "middle")
        txt(root, 800, 468, "Dibaca · 3 jam", 18, ESPRESSO, 700, "end")
    elif case["key"] == "feeling":
        draw_student_failure(root)
    else:
        embed_asset(root, "couple", 70, 76, 910, 910, omit_ids=omit_bg, color_map=color_map)

    # Only the mistaken label and its explanatory correction remain as text on this frame.
    rect(root, 1054, 298, 770, 146, TERRACOTTA, rx=28)
    txt(root, 1439, 390, case["quote"], 50 if number < 3 else 43, WHITE, 700, "middle")
    rect(root, 1054, 574, 770, 178, SAGE, rx=28)
    txt(root, 1112, 635, "REALITANYA", 20, ESPRESSO, 700, letter=1.4)
    txt(root, 1439, 690, case["correction"][0], 30, ESPRESSO, 600, "middle")
    txt(root, 1439, 731, case["correction"][1], 30, ESPRESSO, 600, "middle")
    return root


def draw_student_failure(root):
    """Editable, hand-drawn exam-failure vignette for scene 03B."""
    # A marked score card makes the setback readable before the emotional label appears.
    card(root, 108, 112, 864, 292, CREAM, border=PEACH, radius=26)
    txt(root, 158, 172, "HASIL ULANGAN", 22, ESPRESSO, 700, letter=1.3)
    txt(root, 365, 328, "100", 112, ESPRESSO, 700, "middle")
    path(root, "M292,245 L432,344", stroke=TERRACOTTA, sw=13)
    path(root, "M304,344 L427,251", stroke=TERRACOTTA, sw=9)
    path(root, "M495,292 C536,258 584,257 625,288", stroke=TERRACOTTA, sw=8)
    path(root, "M610,272 L635,291 L603,299 Z", fill=TERRACOTTA)
    txt(root, 770, 325, "45", 116, TERRACOTTA, 700, "middle")
    txt(root, 770, 367, "REMEDI", 22, TERRACOTTA, 700, "middle", letter=1.5)

    # Chair, legs, and desk establish a schoolwork setting rather than an abstract target.
    add(root, "ellipse", cx=552, cy=917, rx=390, ry=30, fill=ESPRESSO, opacity=0.10)
    rect(root, 650, 484, 205, 235, MUSTARD, rx=34, stroke=ESPRESSO, sw=8)
    rect(root, 490, 686, 390, 38, ESPRESSO, rx=18)
    path(root, "M533,720 L512,900 M835,720 L857,900", stroke=ESPRESSO, sw=18)
    path(root, "M500,897 L545,897 M839,897 L881,897", stroke=ESPRESSO, sw=15)

    # Bent legs and slumped torso: the shoulders tip forward and the head drops toward the test.
    path(root, "M493,677 C510,729 477,783 416,826 L390,853 Q417,873 459,858 L552,803 Q585,784 598,749 L650,839 Q673,858 711,844 L676,740 L647,681 Z", fill=PLUM)
    path(root, "M390,851 Q421,862 462,847 L476,864 Q464,879 424,884 L382,876 Z", fill=ESPRESSO)
    path(root, "M674,837 Q695,850 714,833 L732,850 Q719,870 682,871 L659,859 Z", fill=ESPRESSO)
    path(root, "M481,544 C506,500 576,488 640,529 Q672,552 676,631 L654,715 L475,715 Q453,648 458,592 Z", fill=PLUM)
    path(root, "M509,563 Q556,590 608,557", stroke=PEACH, sw=7, opacity=0.65)

    # One arm droops across the desk; the other reaches for the exam sheet.
    path(root, "M493,563 C468,581 449,606 431,636", stroke=PLUM, sw=47)
    path(root, "M431,636 Q410,650 396,665", stroke=PEACH, sw=27)
    path(root, "M635,559 C670,587 685,620 699,660", stroke=PLUM, sw=47)
    path(root, "M699,660 Q704,674 695,686", stroke=PEACH, sw=28)

    # Drooped face with low, inward brows, downcast eyes, and a frown.
    add(root, "ellipse", cx=589, cy=490, rx=61, ry=75, fill=PEACH, transform="rotate(-13 589 490)")
    path(root, "M528,475 C512,431 538,404 565,405 C588,380 629,406 640,439 L633,477 Q612,452 591,451 Q559,447 541,481 Z", fill=ESPRESSO)
    path(root, "M541,475 Q553,465 565,473 M601,468 Q613,460 625,467", stroke=ESPRESSO, sw=6)
    path(root, "M545,487 Q554,494 563,489 M604,482 Q613,489 622,483", stroke=ESPRESSO, sw=4)
    path(root, "M583,490 Q573,507 589,510", stroke=ESPRESSO, sw=4)
    path(root, "M570,526 Q587,512 604,524", stroke=ESPRESSO, sw=5)
    circle(root, 632, 495, 13, PEACH)
    path(root, "M640,495 Q649,497 642,506", stroke=ESPRESSO, sw=3)

    # Paper held up from the desk, with the actual low grade and a remedial mark.
    sheet = add(root, "g", transform="rotate(-5 388 622)")
    card(sheet, 280, 526, 230, 174, CREAM, border=PEACH, radius=12)
    txt(sheet, 302, 557, "ULANGAN", 14, ESPRESSO, 700, letter=0.8)
    line(sheet, 302, 570, 486, 570, PEACH, 3)
    txt(sheet, 395, 631, "45", 62, TERRACOTTA, 700, "middle")
    txt(sheet, 395, 663, "REMEDI", 14, TERRACOTTA, 700, "middle", letter=0.8)
    line(sheet, 304, 682, 479, 682, "#E8DACF", 3)
    # Hand wraps over the upper-right corner so the student is visibly holding the page.
    hand = path(root, "M431,636 Q448,620 463,606 Q474,598 483,606 L489,617 Q485,630 470,641 L450,654", fill=PEACH)

    # Desk is foregrounded to anchor the slumped figure and the marked exam.
    rect(root, 132, 704, 828, 40, MUSTARD, rx=13, stroke=ESPRESSO, sw=7)
    rect(root, 161, 744, 770, 28, PLUM, rx=7)
    path(root, "M194,770 L183,904 M892,770 L906,904", stroke=ESPRESSO, sw=18)
    path(root, "M166,903 L207,903 M886,903 L927,903", stroke=ESPRESSO, sw=15)
    # The sheet sits in front of the desktop so its grade stays legible in the animation frame.
    root.append(sheet)
    root.append(hand)


def build_scene_3_case_1(scene):
    return draw_case_frame(0)


def build_scene_3_case_2(scene):
    return draw_case_frame(1)


def build_scene_3_case_3(scene):
    return draw_case_frame(2)


def build_scene_3_combined(scene):
    root = clean_frame(
        "Scene 3 combined examples: one behavior does not define a diagnosis",
        "The unanswered message, sadness, and disagreement examples meet behind the centered on-screen statement: One behavior does not equal a psychological diagnosis.",
        CREAM,
    )
    panel_w = 640
    case_positions = [(0, CASE_TINTS[0]), (640, CASE_TINTS[1]), (1280, CASE_TINTS[2])]
    for i, (x, tint) in enumerate(case_positions):
        rect(root, x, 0, panel_w, H, tint)
        case = SCENE3_CASES[i]
        if case["key"] == "feeling":
            # Reuse the exam-failure vignette so the combined frame matches 03B.
            miniature = add(root, "g", transform="translate(722 0) scale(0.44)")
            draw_student_failure(miniature)
        else:
            omit_bg = {"freepik--background-complete--inject-2"}
            embed_asset(root, case["key"], x + 70, 24, 500, 410, omit_ids=omit_bg, color_map=case["colors"])
    rect(root, 95, 433, 1730, 214, PLUM, rx=34)
    txt(root, 548, 564, "One behavior", 58, MUSTARD, 700, "middle")
    line(root, 925, 520, 987, 520, TERRACOTTA, 9)
    line(root, 925, 552, 987, 552, TERRACOTTA, 9)
    path(root, "M940,573 L972,501", stroke=TERRACOTTA, sw=9)
    txt(root, 1346, 562, "A psychological diagnosis.", 48, WHITE, 700, "middle")
    return root


def build_scene_3(scene):
    return build_scene_3_combined(scene)


def build_scene_4(scene):
    root = start_scene(scene, CREAM, PLUM)
    card(root, 80, 250, 650, 620, CREAM, PEACH, 24)
    problem_colors = {"#263238": ESPRESSO, "#407BFF": SAGE, "#FF5652": TERRACOTTA,
                      "#FF8B7B": PEACH, "#E6E6E6": PEACH, "#EBEBEB": CREAM,
                      "#F0F0F0": CREAM, "#F5F5F5": CREAM, "#FAFAFA": CREAM}
    embed_asset(root, "problem", 100, 268, 610, 580, color_map=problem_colors)
    txt(root, 405, 856, "MENCARI KONTEKS", 15, PLUM, 700, "middle", letter=1)
    pill(root, 790, 244, 1020, 82, "LABEL TANPA KONTEKS", TERRACOTTA, WHITE, 26)
    path(root, "M1300,328 C1300,365 1300,365 1300,390", stroke=PEACH, sw=7)
    path(root, "M1287,377 L1300,398 L1313,377 Z", fill=PEACH)
    outcomes = [
        ("01", "Salah memahami\ndiri sendiri", MUSTARD),
        ("02", "Salah menilai\norang lain", TERRACOTTA),
        ("03", "Konflik dalam\nhubungan", PLUM),
        ("04", "Stigma terhadap\ngangguan mental", ESPRESSO),
    ]
    x0, y, cw, ch, gap = 780, 430, 238, 320, 22
    for i, (number, label, accent) in enumerate(outcomes):
        x = x0 + i * (cw + gap)
        rect(root, x + 7, y + 10, cw, ch, ESPRESSO, rx=20, opacity=0.10)
        rect(root, x, y, cw, ch, accent, rx=20, stroke=accent, sw=2)
        label_color = ESPRESSO if accent == MUSTARD else WHITE
        circle(root, x + 46, y + 75, 25, accent)
        txt(root, x + 46, y + 82, number, 15, label_color, 700, "middle")
        for row, text_line in enumerate(label.split("\n")):
            txt(root, x + 26, y + 151 + row * 34, text_line, 23, label_color, 700)
        txt(root, x + 26, y + 270, "DAMPAK", 14, label_color, 700, letter=1)
        if i < 3:
            line(root, x + cw + 4, y + 170, x + cw + gap - 3, y + 170, ESPRESSO, 5)
            path(root, f"M{x+cw+gap-11},{y+160} L{x+cw+gap-1},{y+170} L{x+cw+gap-11},{y+180} Z", fill=ESPRESSO)
    return root


def build_scene_5(scene):
    root = start_scene(scene, CREAM, PLUM)
    card(root, 72, 240, 680, 638, CREAM, PEACH, 26)
    doc_colors = {"#263238": ESPRESSO, "#407BFF": PLUM, "#D58745": MUSTARD, "#EBB376": PEACH,
                  "#E0E0E0": PEACH, "#E6E6E6": CREAM, "#EBEBEB": CREAM, "#F0F0F0": CREAM,
                  "#F5F5F5": CREAM, "#FAFAFA": CREAM}
    research_colors = {**doc_colors, "#2E353A": ESPRESSO, "#577AFA": PLUM, "#D4827D": TERRACOTTA,
                       "#EBC0BB": PEACH, "#EEC1BB": PEACH}
    embed_asset(root, "document", 92, 250, 595, 580, color_map=doc_colors)
    embed_asset(root, "research", 390, 516, 326, 314, color_map=research_colors)
    pill(root, 125, 832, 555, 42, "KODE ETIK PSIKOLOGI INDONESIA", PLUM, MUSTARD, 15, stroke=PEACH)
    articles = [
        ("9", "Dasar ilmiah & profesional", "Keputusan psikologi berlandaskan pengetahuan yang teruji."),
        ("15", "Menghindari dampak buruk", "Sampaikan informasi dengan mempertimbangkan dampaknya."),
        ("28", "Pernyataan publik yang cermat", "Bijaksana, jujur, teliti, hati-hati, dan berbasis ilmiah."),
        ("31", "Komunikasi media bertanggung jawab", "Selaras dengan kode etik dan praktik psikologi yang tepat."),
    ]
    for i, (num, title, desc) in enumerate(articles):
        y = 230 + i * 164
        rect(root, 799, y + 8, 1032, 145, ESPRESSO, rx=18, opacity=0.10)
        rect(root, 792, y, 1032, 145, PLUM, rx=18, stroke=PEACH, sw=2)
        circle(root, 858, y + 72, 39, MUSTARD)
        txt(root, 858, y + 82, num, 27, ESPRESSO, 700, "middle")
        txt(root, 924, y + 60, "PASAL " + num, 14, MUSTARD, 700, letter=1)
        txt(root, 924, y + 91, title, 24, CREAM, 700)
        txt(root, 924, y + 121, desc, 17, CREAM, 400)
    pill(root, 792, 906, 1032, 53, "SCIENTIFIC BASIS  +  PROFESSIONAL CARE", PLUM, CREAM, 19, stroke=PEACH)
    return root


def build_scene_6(scene):
    root = start_scene(scene, CREAM, SAGE)
    card(root, 72, 235, 660, 660, CREAM, PEACH, 24)
    todo_colors = {"#263238": ESPRESSO, "#407BFF": SAGE, "#E6E6E6": CREAM, "#EBEBEB": CREAM,
                   "#F0F0F0": CREAM, "#F5F5F5": CREAM, "#FAFAFA": CREAM,
                   "#ED847E": PEACH, "#FFC3BD": PEACH}
    embed_asset(root, "todo", 90, 250, 620, 615, color_map=todo_colors)
    steps = [
        ("1", "Understand before you label", "Cari tahu makna istilah sebelum memakainya."),
        ("2", "Describe, don’t diagnose", "Jelaskan perilaku spesifik yang membuat tidak nyaman."),
        ("3", "Learn, don’t self-diagnose", "Konten edukasi bukan pengganti asesmen profesional."),
    ]
    for i, (num, title, desc) in enumerate(steps):
        y = 245 + i * 216
        card(root, 770, y, 1070, 190, CREAM, PEACH, 20)
        circle(root, 844, y + 75, 37, MUSTARD)
        txt(root, 844, y + 85, num, 26, ESPRESSO, 700, "middle")
        title_node = add(root, "text", x=910, y=y + 70, fill=PLUM, font_family="Arial, Helvetica, sans-serif", font_size=27, font_weight=700)
        split = [("Understand before you ", "label"), ("Describe, don’t ", "diagnose"), ("Learn, don’t ", "self-diagnose")][i]
        for value, color in [(split[0], PLUM), (split[1], TERRACOTTA)]:
            span = add(title_node, "tspan", fill=color)
            span.text = value
        txt(root, 910, y + 116, desc, 20, ESPRESSO, 400)
        # Small, editable line icons: learn, describe, and know the limit.
        if i == 0:
            circle(root, 1740, y + 73, 25, "none", stroke=SAGE, sw=5)
            line(root, 1758, y + 91, 1779, y + 112, SAGE, 6)
        elif i == 1:
            rect(root, 1700, y + 45, 93, 58, PEACH, rx=12, stroke=SAGE, sw=3)
            path(root, f"M1720,{y+72} L1734,{y+86} L1769,{y+58}", stroke=SAGE, sw=5)
        else:
            circle(root, 1744, y + 75, 29, CREAM, stroke=SAGE, sw=3)
            path(root, f"M1728,{y+76} L1740,{y+88} L1762,{y+61}", stroke=SAGE, sw=5)
    return root


def face(parent, cx, cy, emotion, color):
    circle(parent, cx, cy, 38, color, stroke=CREAM, sw=4)
    circle(parent, cx - 12, cy - 8, 3.5, ESPRESSO)
    circle(parent, cx + 12, cy - 8, 3.5, ESPRESSO)
    if emotion == "joy":
        path(parent, f"M{cx-15},{cy+7} Q{cx},{cy+25} {cx+15},{cy+7}", stroke=ESPRESSO, sw=3)
    elif emotion == "sadness":
        path(parent, f"M{cx-14},{cy+19} Q{cx},{cy+4} {cx+14},{cy+19}", stroke=ESPRESSO, sw=3)
        path(parent, f"M{cx+22},{cy-1} Q{cx+29},{cy+8} {cx+22},{cy+14} Q{cx+16},{cy+8} {cx+22},{cy-1}", fill=PEACH)
    elif emotion == "fear":
        circle(parent, cx, cy + 13, 5, "none", stroke=ESPRESSO, sw=3)
    else:
        line(parent, cx - 15, cy + 12, cx + 15, cy + 12, ESPRESSO, 3)


def build_scene_7(scene):
    root = start_scene(scene, PLUM, TERRACOTTA, dark=True)
    social_colors = {"#263238": ESPRESSO, "#407BFF": SAGE, "#630F0F": ESPRESSO,
                     "#7F3E3B": PEACH, "#A02724": PEACH, "#B55B52": PEACH,
                     "#E0E0E0": PLUM, "#E6E6E6": PLUM, "#EBEBEB": PLUM,
                     "#F0F0F0": PLUM, "#F5F5F5": PLUM, "#FAFAFA": PLUM,
                     "#ED847E": PEACH, "#FFC3BD": PEACH}
    embed_asset(root, "social", 88, 250, 770, 620,
                omit_ids={"freepik--background-complete--inject-2"}, color_map=social_colors)
    # Viral labels fade from terracotta toward the plum closing field.
    pill(root, 114, 258, 180, 40, "toxic", TERRACOTTA, WHITE, 15, opacity=0.58)
    pill(root, 300, 258, 204, 40, "trauma", TERRACOTTA, WHITE, 15, opacity=0.42)
    pill(root, 512, 258, 190, 40, "avoidant", TERRACOTTA, WHITE, 15, opacity=0.28)
    txt(root, 1373, 290, "NORMAL HUMAN EMOTIONS", 15, CREAM, 700, "middle", letter=1.2)
    emotions = [("joy", "Gembira", PEACH), ("sadness", "Sedih", SAGE), ("fear", "Takut", MUSTARD), ("calm", "Tenang", CREAM)]
    for i, (kind, label, color) in enumerate(emotions):
        x = 960 + i * 210
        face(root, x + 95, 370, kind, color)
        txt(root, x + 95, 444, label, 17, CREAM, 700, "middle")
    line(root, 980, 500, 1764, 500, ESPRESSO, 2)
    txt(root, 1250, 638, "THINK BEFORE YOU", 40, CREAM, 700, "middle", letter=0.6)
    txt(root, 1594, 638, "LABEL.", 40, MUSTARD, 700, "middle", letter=0.6)
    txt(root, 1373, 692, "Understand psychology. Use it responsibly.", 19, CREAM, 400, "middle")
    pill(root, 1088, 748, 570, 55, "PSYCHOLOGY HELPS US UNDERSTAND PEOPLE", MUSTARD, ESPRESSO, 16)
    txt(root, 1373, 850, "END CARD  ·  VIDEO TITLE  ·  CREATOR IDENTITY", 14, PEACH, 700, "middle", letter=0.7)
    rect(root, 92, 926, 1736, 47, ESPRESSO, rx=15)
    txt(root, 960, 958, "Stay curious. Stay kind.", 23, CREAM, 700, "middle")
    return root


DRAW = [build_scene_1, build_scene_2, build_scene_3, build_scene_4, build_scene_5, build_scene_6, build_scene_7]


def create_previews():
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    SCENES_DIR.mkdir(parents=True, exist_ok=True)
    for stale in [
        SCENES_DIR / "scene-02-what-is-therapy-speak.svg",
        PREVIEW_DIR / "scene-02-what-is-therapy-speak.png",
        SCENES_DIR / "scene-03-normal-emotions-become-labels.svg",
        PREVIEW_DIR / "scene-03-normal-emotions-become-labels.png",
    ]:
        if stale.exists():
            stale.unlink()
    thumbs = []
    for scene, draw_scene in zip(SCENES, DRAW):
        if scene["no"] == 2:
            frames = [
                ("02A", "02a-comment-scroll", "Phone Zoom + Comment Scroll", "00:30–00:45", build_scene_2),
                ("02B", "02b-hashtag-cloud", "Hashtag Cloud Transition", "00:45–00:58", build_scene_2b),
                ("02C", "02c-on-screen-text", "Understanding Psychology · On-Screen Text", "00:58–01:05", build_scene_2c),
            ]
        elif scene["no"] == 3:
            frames = [
                ("03A", "03a-case-01-avoidant", "Case 1 · Unanswered Message", "01:05–01:20", build_scene_3_case_1),
                ("03B", "03b-case-02-depressed", "Case 2 · Sadness After Failure", "01:20–01:35", build_scene_3_case_2),
                ("03C", "03c-case-03-gaslighting", "Case 3 · Disagreement", "01:35–01:50", build_scene_3_case_3),
                ("03D", "03d-combined-on-screen-text", "Combined Examples · On-Screen Text", "01:50–01:55", build_scene_3_combined),
            ]
        else:
            frames = [(f"{scene['no']:02d}", f"{scene['no']:02d}-{scene['slug']}", scene["title"], scene["time"], draw_scene)]
        for code, slug, title, time, renderer in frames:
            frame_scene = dict(scene, title=title, time=time)
            root = renderer(frame_scene)
            ET.indent(root, space="  ")
            svg_path = SCENES_DIR / f"scene-{slug}.svg"
            svg_path.write_bytes(ET.tostring(root, encoding="utf-8", xml_declaration=True))
            png_path = PREVIEW_DIR / f"scene-{slug}.png"
            cairosvg.svg2png(url=str(svg_path), write_to=str(png_path), output_width=1440, output_height=810)
            thumbs.append((code, title, png_path))
    components = build_scene_2_components()
    ET.indent(components, space="  ")
    component_svg = SCENES_DIR / "scene-02-komponen.svg"
    component_svg.write_bytes(ET.tostring(components, encoding="utf-8", xml_declaration=True))
    component_png = PREVIEW_DIR / "scene-02-komponen.png"
    cairosvg.svg2png(url=str(component_svg), write_to=str(component_png), output_width=560, output_height=1080)
    transparent_preview = Image.open(component_png).convert("RGBA")
    preview_canvas = Image.new("RGB", transparent_preview.size, CREAM)
    preview_canvas.paste(transparent_preview, (0, 0), transparent_preview.getchannel("A"))
    preview_canvas.save(component_png)
    # Three columns make a compact visual audit sheet while retaining the 16:9 frames.
    card_w, card_h, gap, top = 420, 236, 26, 56
    cols, rows = 4, (len(thumbs) + 3) // 4
    sheet = Image.new("RGB", (cols * card_w + (cols + 1) * gap, rows * (card_h + 48) + (rows + 1) * gap + top), "#F1F4F8")
    draw = ImageDraw.Draw(sheet)
    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
    label_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    draw.text((gap, 15), f"STORYBOARD C · VISUAL AUDIT · 7 SCENES / {len(thumbs)} FRAMES", font=title_font, fill=ESPRESSO)
    for i, (code, title, png_path) in enumerate(thumbs):
        col, row = i % cols, i // cols
        x = gap + col * (card_w + gap)
        y = top + gap + row * (card_h + 48 + gap)
        im = Image.open(png_path).convert("RGB").resize((card_w, card_h), Image.Resampling.LANCZOS)
        sheet.paste(im, (x, y))
        draw.text((x, y + card_h + 9), f"{code}  ·  {title}", font=label_font, fill=ESPRESSO)
    contact = PREVIEW_DIR / "contact-sheet.png"
    sheet.save(contact, quality=94)
    return contact


def shade(cell, fill):
    props = cell._tc.get_or_add_tcPr()
    node = OxmlElement("w:shd")
    node.set(qn("w:fill"), fill.lstrip("#"))
    props.append(node)


def set_cell_margins(cell, top=90, start=120, bottom=90, end=120):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = OxmlElement("w:tcMar")
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement("w:" + name)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        margins.append(node)
    tc_pr.append(margins)


def add_hyperlink(paragraph, label, url):
    part = paragraph.part
    rid = part.relate_to(url, RT.HYPERLINK, is_external=True)
    link = OxmlElement("w:hyperlink")
    link.set(qn("r:id"), rid)
    run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), PLUM.lstrip("#"))
    rpr.append(color)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    rpr.append(underline)
    run.append(rpr)
    text = OxmlElement("w:t")
    text.text = label
    run.append(text)
    link.append(run)
    paragraph._p.append(link)
    return link


def set_run(run, size=10, color=INK, bold=False):
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor.from_string(color.lstrip("#"))


def add_field(cell, label, content, size=10):
    p = cell.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(label.upper() + "\n")
    set_run(r, 8, BLUE_DARK, True)
    r = p.add_run(content)
    set_run(r, size, INK)
    return p


def make_docx(contact):
    doc = Document()
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Inches(11.7), Inches(8.3)
    sec.top_margin = Inches(0.45)
    sec.bottom_margin = Inches(0.45)
    sec.left_margin = Inches(0.5)
    sec.right_margin = Inches(0.5)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.font.color.rgb = RGBColor.from_string(ESPRESSO.lstrip("#"))
    normal.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("THERAPY SPEAK GONE WRONG")
    set_run(r, 28, INK, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run("Storyboard produksi · Layout C · 7 scene · 12 keyframe · 00:00–04:00")
    set_run(r, 15, PLUM, True)
    p = doc.add_paragraph("Setiap frame tersedia sebagai SVG mandiri yang editable. Scene 2 memiliki tiga frame: guliran komentar, hashtag, lalu dua pernyataan di tengah. Scene 3 memiliki tiga case terpisah dan satu frame gabungan dengan teks penutup di tengah. Ilustrasi Storyset ditanam sebagai vektor dan diwarnai ulang sebagai SVG; komposisi tambahan juga digambar sebagai vektor. Tidak ada gambar generatif.")
    p.paragraph_format.space_after = Pt(10)
    for r in p.runs:
        set_run(r, 10, MUTED)

    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(4)
    r = heading.add_run("AUDIT VISUAL")
    set_run(r, 14, INK, True)
    audit = [
        "Scene 2 menggunakan cream, espresso, dan plum sebagai warna utama; istilah diberi chip terracotta, mustard, atau peach sesuai pedoman Opsi B.",
        "Scene 3 memakai tint peach, sage, dan mustard per case. Label keliru memakai terracotta dengan teks putih; koreksi memakai sage dengan teks espresso.",
        "Frame 02A–02C dan 03A–03D adalah komposisi animasi bersih tanpa judul board, waktu, footer, atau catatan produksi di dalam frame.",
        "Ilustrasi Storyset tetap vektor dan diwarnai ulang lewat SVG. Frame presenter tetap slot footage; tidak ada gambar yang dihasilkan.",
    ]
    for item in audit:
        p = doc.add_paragraph(style="List Bullet")
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(item)
        set_run(r, 9, INK)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    r = p.add_run("Warna sistem:   ")
    set_run(r, 9, INK, True)
    for name, color in zip(PALETTE_NAMES, PALETTE):
        r = p.add_run(f"   {name} {color}   ")
        set_run(r, 8, ESPRESSO, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("Kredit Storyset: Freepik License, free with attribution. Tautan unduh per ilustrasi ada di scene terkait. Periksa ketentuan Storyset saat penggunaan final.")
    set_run(r, 9, MUTED)
    add_hyperlink(p, "Storyset terms", "https://storyset.com/terms")

    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("CONTACT SHEET · AUDIT 7 SCENE / 12 FRAME")
    set_run(r, 20, INK, True)
    doc.add_picture(str(contact), width=Inches(10.65))
    p = doc.add_paragraph("Ritme visual bergerak dari presenter → feed → tiga contoh → dampak → etika → checklist → closing. Semua frame memakai palet dan grid yang sama.")
    for r in p.runs:
        set_run(r, 9, MUTED)

    for scene in SCENES:
        doc.add_page_break()
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(1)
        r = p.add_run(f"SCENE {scene['no']:02d}  ·  {scene['time']}")
        set_run(r, 11, PLUM, True)
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(scene["title"])
        set_run(r, 21, INK, True)
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Inches(6.5)
        table.columns[1].width = Inches(4.05)
        left, right = table.rows[0].cells
        left.width, right.width = Inches(6.5), Inches(4.05)
        left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        shade(left, CREAM)
        shade(right, CREAM)
        set_cell_margins(left, 100, 100, 100, 100)
        set_cell_margins(right, 100, 140, 100, 140)
        if scene["no"] == 2:
            left.paragraphs[0].text = ""
            inner = left.add_table(rows=1, cols=3)
            inner.autofit = False
            for column in inner.columns:
                column.width = Inches(2.05)
            frame_data = [
                ("02A · 00:30–00:45", "scene-02a-comment-scroll.png"),
                ("02B · 00:45–00:58", "scene-02b-hashtag-cloud.png"),
                ("02C · 00:58–01:05", "scene-02c-on-screen-text.png"),
            ]
            for cell, (label, filename) in zip(inner.rows[0].cells, frame_data):
                cell.width = Inches(2.05)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                caption = cell.paragraphs[0]
                caption.alignment = 1
                caption.paragraph_format.space_after = Pt(5)
                run = caption.add_run(label)
                set_run(run, 7, PLUM, True)
                image_para = cell.add_paragraph()
                image_para.alignment = 1
                image_para.add_run().add_picture(str(PREVIEW_DIR / filename), width=Inches(1.96))
        elif scene["no"] == 3:
            left.paragraphs[0].text = ""
            inner = left.add_table(rows=2, cols=2)
            inner.autofit = False
            for column in inner.columns:
                column.width = Inches(3.08)
            frame_data = [
                ("03A · 01:05–01:20", "scene-03a-case-01-avoidant.png"),
                ("03B · 01:20–01:35", "scene-03b-case-02-depressed.png"),
                ("03C · 01:35–01:50", "scene-03c-case-03-gaslighting.png"),
                ("03D · 01:50–01:55", "scene-03d-combined-on-screen-text.png"),
            ]
            for cell, (label, filename) in zip([c for row in inner.rows for c in row.cells], frame_data):
                cell.width = Inches(3.08)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                caption = cell.paragraphs[0]
                caption.alignment = 1
                caption.paragraph_format.space_after = Pt(3)
                run = caption.add_run(label)
                set_run(run, 7, PLUM, True)
                image_para = cell.add_paragraph()
                image_para.alignment = 1
                image_para.add_run().add_picture(str(PREVIEW_DIR / filename), width=Inches(2.94))
        else:
            image = PREVIEW_DIR / f"scene-{scene['no']:02d}-{scene['slug']}.png"
            p = left.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(str(image), width=Inches(6.2))
        p = right.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("VISUAL")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["visual"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ON-SCREEN TEXT")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["screen"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK, True)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("GERAK / KAMERA")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["motion"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("SCRIPT / VOICE-OVER")
        set_run(r, 8, PLUM, True)
        p = right.add_paragraph(scene["voice"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ASET & PERAN")
        set_run(r, 8, PLUM, True)
        for role in scene["roles"]:
            p = right.add_paragraph(role, style="List Bullet")
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs: set_run(r, 8, INK)
        for asset_key in scene["credits"]:
            info = ASSETS[asset_key]
            p = right.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            add_hyperlink(p, "Unduh " + info["title"] + " · Storyset Rafiki", info["url"])
            for r in p.runs: set_run(r, 8, PLUM)
        p = right.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run("BACKGROUND: ")
        set_run(r, 8, PLUM, True)
        r = p.add_run(scene["background"])
        set_run(r, 8, MUTED)

    out = ROOT / "Storyboard_Therapy_Speak_Gone_Wrong.docx"
    doc.save(out)
    return out


def make_website():
    rows = []
    for scene in SCENES:
        asset_links = []
        for key in scene["credits"]:
            info = ASSETS[key]
            asset_links.append(f'<a href="{info["url"]}" target="_blank" rel="noreferrer">{info["title"]} ↗</a>')
        role_text = "<br>".join(scene["roles"])
        if scene["no"] == 2:
            frame_markup = '''<div class="frame-stack">
              <section class="frame-card"><div class="frame-label"><b>02A · PHONE + COMMENT SCROLL</b><span>00:30–00:45</span></div><a href="scenes/scene-02a-comment-scroll.svg" target="_blank" aria-label="Open editable SVG for Scene 2A"><img src="scenes/scene-02a-comment-scroll.svg" alt="Cream living room with TV and art at left, an espresso phone in the center, and the reclining man on a sage sofa at right"></a><div class="file-links"><a href="scenes/scene-02a-comment-scroll.svg" download>Download editable SVG</a><a href="previews/scene-02a-comment-scroll.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02B · HASHTAG CLOUD</b><span>00:45–00:58</span></div><a href="scenes/scene-02b-hashtag-cloud.svg" target="_blank" aria-label="Open editable SVG for Scene 2B"><img src="scenes/scene-02b-hashtag-cloud.svg" alt="Overlapping hashtag cards fill the cream frame"></a><div class="file-links"><a href="scenes/scene-02b-hashtag-cloud.svg" download>Download editable SVG</a><a href="previews/scene-02b-hashtag-cloud.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02C · ON-SCREEN TEXT</b><span>00:58–01:05</span></div><a href="scenes/scene-02c-on-screen-text.svg" target="_blank" aria-label="Open editable SVG for Scene 2C"><img src="scenes/scene-02c-on-screen-text.svg" alt="Centered sage and terracotta statements on a cream background"></a><div class="file-links"><a href="scenes/scene-02c-on-screen-text.svg" download>Download editable SVG</a><a href="previews/scene-02c-on-screen-text.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02 · SCROLL COMPONENTS</b><span>6 extra comments</span></div><p class="component-copy">The six separate SVG groups use IDs comment-05 through comment-10, ready to animate inside the phone feed.</p><div class="file-links"><a href="scenes/scene-02-komponen.svg" target="_blank">Open SVG components</a><a href="scenes/scene-02-komponen.svg" download>Download SVG</a><a href="previews/scene-02-komponen.png" target="_blank">Preview sheet</a></div></section>
            </div>'''
        elif scene["no"] == 3:
            frame_cards = []
            for frame in scene["frames"]:
                svg_name = f"scene-{frame['slug']}.svg"
                png_name = f"scene-{frame['slug']}.png"
                frame_cards.append(f'''<section class="frame-card"><div class="frame-label"><b>{frame['code']} · {frame['title']}</b><span>{frame['time']}</span></div><a href="scenes/{svg_name}" target="_blank" aria-label="Open editable SVG for {frame['code']}"><img src="scenes/{svg_name}" alt="{frame['title']} animation frame"></a><div class="file-links"><a href="scenes/{svg_name}" download>Download editable SVG</a><a href="previews/{png_name}" download>Download PNG preview</a></div></section>''')
            frame_markup = '<div class="case-frame-grid">' + "".join(frame_cards) + '</div>'
        else:
            svg_name = f"scene-{scene['no']:02d}-{scene['slug']}.svg"
            png_name = f"scene-{scene['no']:02d}-{scene['slug']}.png"
            frame_markup = f'<a href="scenes/{svg_name}" target="_blank" aria-label="Open editable SVG for scene {scene["no"]}"><img src="scenes/{svg_name}" alt="Editable scene illustration, scene {scene["no"]}"></a><div class="file-links"><a href="scenes/{svg_name}" download>Download editable SVG</a><a href="previews/{png_name}" download>Download PNG preview</a></div>'
        rows.append(f'''<article class="board-row">
          <header class="row-meta"><span class="scene-number">{scene['no']:02d}</span><span class="time">{scene['time']}</span><h2>{scene['title']}</h2><p>{scene['visual']}</p></header>
          <div class="row-frame">{frame_markup}</div>
          <dl class="row-notes"><div><dt>On-screen text</dt><dd>{scene['screen'].replace(chr(10), '<br>')}</dd></div><div><dt>Motion / camera</dt><dd>{scene['motion']}</dd></div><div><dt>Script / voice-over</dt><dd>{scene['voice']}</dd></div><div><dt>Assets & roles</dt><dd>{role_text}</dd><p class="asset-sources">{' · '.join(asset_links) if asset_links else 'Custom editable SVG vectors; no external illustration.'}</p></div></dl>
        </article>''')
    html = '''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Editable production board with twelve SVG animation frames for Therapy Speak Gone Wrong.">
<title>Therapy Speak Gone Wrong · Production Board</title>
<style>
:root{--ink:#2a1f2d;--plum:#3d2645;--sage:#7fb89a;--terracotta:#e4572e;--mustard:#f2b134;--peach:#ffc4a3;--cream:#fff6ec;--blue:#3d2645;--blue-dark:#3d2645;--paper:#fff6ec;--line:#e8dacf;--muted:#665967;--coral:#e4572e;--white:#fff;--shadow:0 12px 30px rgba(42,31,45,.09)}
*{box-sizing:border-box}html{background:var(--paper);color:var(--ink);font-family:Arial,Helvetica,sans-serif}body{margin:0;padding:42px 28px 92px}.shell{max-width:1560px;margin:0 auto}.hero{display:grid;grid-template-columns:1fr auto;gap:32px;align-items:end;margin:8px 0 32px}.eyebrow{color:var(--blue-dark);font-size:12px;font-weight:700;letter-spacing:.14em;text-transform:uppercase}.hero h1{max-width:920px;margin:12px 0 12px;font-size:clamp(38px,5.5vw,76px);line-height:.98;letter-spacing:-.055em}.hero p{max-width:760px;margin:0;color:var(--muted);font-size:16px;line-height:1.55}.hero-actions{display:flex;flex-direction:column;gap:9px;align-items:stretch}.action{display:inline-block;padding:12px 16px;border:1px solid var(--line);border-radius:999px;background:var(--white);color:var(--ink);font-size:13px;font-weight:700;text-align:center;text-decoration:none}.action.primary{border-color:var(--blue);background:var(--blue);color:var(--white)}.style-strip{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:28px 0 18px;padding:15px 18px;border:1px solid var(--line);border-radius:14px;background:var(--white);color:var(--muted);font-size:12px}.swatch{display:inline-flex;align-items:center;gap:6px}.swatch i{width:15px;height:15px;border:1px solid #d5dce2;border-radius:50%}.board{display:grid;gap:13px}.board-row{display:grid;grid-template-columns:minmax(0,1fr);gap:18px;align-items:stretch;padding:15px;border:1px solid var(--line);border-radius:17px;background:var(--white);box-shadow:var(--shadow)}.row-meta{display:flex;flex-direction:column;align-items:flex-start;padding:5px 0}.scene-number{display:grid;place-items:center;width:34px;height:34px;border-radius:50%;background:var(--ink);color:white;font-size:12px;font-weight:700}.time{margin:12px 0 9px;color:var(--blue-dark);font-size:11px;font-weight:700;letter-spacing:.08em}.row-meta h2{margin:0;font-size:19px;line-height:1.12;letter-spacing:-.02em}.row-meta p{margin:12px 0 0;color:var(--muted);font-size:12px;line-height:1.45}.row-frame{min-width:0;display:flex;flex-direction:column;justify-content:center}.row-frame img{display:block;width:100%;aspect-ratio:16/9;height:auto;object-fit:contain;object-position:center;border-radius:10px;background:#fff6ec}.frame-stack,.case-frame-grid{display:grid;grid-template-columns:minmax(0,1fr);gap:22px}.frame-card{padding:9px;border:1px solid var(--line);border-radius:12px;background:#fffaf5}.frame-label{display:flex;justify-content:space-between;gap:8px;margin-bottom:7px;color:var(--blue-dark);font-size:10px;letter-spacing:.04em}.frame-label span{color:var(--muted);font-weight:400}.component-copy{margin:0;color:var(--muted);font-size:11px;line-height:1.45}.file-links{display:flex;flex-wrap:wrap;justify-content:space-between;gap:10px;padding:9px 2px 0}.file-links a,.asset-sources a{color:var(--blue-dark);font-size:11px;font-weight:700;text-decoration:none}.file-links a:hover,.asset-sources a:hover{text-decoration:underline}.row-notes{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px;align-content:center;margin:0;padding:2px 4px}.row-notes div{min-width:0}.row-notes dt{margin-bottom:4px;color:var(--blue-dark);font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase}.row-notes dd{margin:0;color:#50616d;font-size:11px;line-height:1.4}.asset-sources{margin:7px 0 0}.asset-sources a{margin-right:6px}.audit{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:28px 0;padding:20px;border:1px solid var(--line);border-radius:16px;background:#fff0e6}.audit h2{margin:0 0 8px;font-size:16px}.audit p{margin:0;color:#566974;font-size:12px;line-height:1.5}.footer{margin-top:24px;padding-top:16px;border-top:1px solid var(--line);color:var(--muted);font-size:11px;line-height:1.5}.footer a{color:var(--blue-dark)}
@media(max-width:1100px){.board-row{grid-template-columns:minmax(0,1fr)}.row-notes{grid-column:auto;grid-template-columns:1fr 1fr}.hero{grid-template-columns:1fr}.hero-actions{flex-direction:row;justify-content:flex-start}.action{min-width:180px}}
@media(max-width:700px){body{padding:24px 12px 60px}.hero h1{font-size:42px}.board-row{grid-template-columns:1fr;gap:12px;padding:13px}.row-meta{display:grid;grid-template-columns:40px 1fr;column-gap:8px;align-items:center}.scene-number{grid-row:span 2}.time{margin:0 0 4px}.row-meta h2{grid-column:2}.row-meta p{grid-column:1/-1;margin-top:10px}.row-notes{grid-column:auto;grid-template-columns:1fr}.hero-actions{flex-wrap:wrap}.action{min-width:0}.audit{grid-template-columns:1fr}}
</style>
</head>
<body><main class="shell">
<header class="hero"><div><div class="eyebrow">Editable production board · Layout C</div><h1>Therapy Speak Gone Wrong</h1><p>Seven scenes from the supplied script, arranged as twelve animation keyframes. Scene 2 has three frames; Scene 3 separates each case before its combined closing frame. Every frame is an editable SVG.</p></div><div class="hero-actions"><a class="action primary" href="Storyboard_Therapy_Speak_Gone_Wrong.docx" download>Download Word storyboard</a><a class="action" href="Storyboard_Editable_Source.zip" download>Download editable source bundle</a></div></header>
<div class="style-strip"><strong>Opsi B · Warm Peach &amp; Mustard</strong><span class="swatch"><i style="background:#3d2645"></i>deep plum</span><span class="swatch"><i style="background:#7fb89a"></i>sage</span><span class="swatch"><i style="background:#e4572e"></i>terracotta</span><span class="swatch"><i style="background:#f2b134"></i>mustard</span><span class="swatch"><i style="background:#ffc4a3"></i>peach</span><span class="swatch"><i style="background:#fff6ec"></i>cream</span><span class="swatch"><i style="background:#2a1f2d"></i>espresso</span><span>· terracotta for mistaken labels; sage for corrections</span></div>
<section class="board" aria-label="Seven scene production board">
''' + "\n".join(rows) + '''
</section>
<section class="audit"><div><h2>Visual audit</h2><p>Scene 2 follows Opsi B with cream, espresso, and plum as its dominant colors; hashtag cards use terracotta, mustard, and peach. Scene 3 uses the case tints peach #FFE3D3, sage #DCEFE4, and mustard #FCE9C2. Incorrect labels use terracotta; corrections use sage with espresso text. The animation frames were rendered and checked visually.</p></div><div><h2>Editing & attribution</h2><p>Open or download each SVG to edit its vector layers and text. The source bundle contains the HTML board and all scene SVGs. Storyset art is free with attribution; use the per-scene links and review <a href="https://storyset.com/terms" target="_blank" rel="noreferrer">Storyset terms</a> for final use.</p></div></section>
<footer class="footer">The presenter frame is a footage slot. Custom scene backgrounds, chips, labels, article cards, faces, and icons are hand-built SVG shapes; no images were generated. Storyset illustrations are embedded as SVG groups and remain vector artwork.</footer>
</main></body></html>'''
    output = ROOT / "index.html"
    output.write_text(html, encoding="utf-8")
    return output


def make_zip():
    import zipfile
    out = ROOT / "Storyboard_Editable_Source.zip"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.write(ROOT / "index.html", "index.html")
        archive.write(ROOT / "build_storyboard.py", "build_storyboard.py")
        archive.write(ROOT / "Storyboard_Therapy_Speak_Gone_Wrong.docx", "Storyboard_Therapy_Speak_Gone_Wrong.docx")
        readme = ROOT / "README.md"
        if readme.exists():
            archive.write(readme, "README.md")
        for path in sorted(SCENES_DIR.glob("*.svg")):
            archive.write(path, "scenes/" + path.name)
        for path in sorted(ASSETS_DIR.glob("*.svg")):
            archive.write(path, "assets/" + path.name)
        for path in sorted(PREVIEW_DIR.glob("*.png")):
            archive.write(path, "previews/" + path.name)
    return out


def main():
    SCENES_DIR.mkdir(parents=True, exist_ok=True)
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    contact = create_previews()
    docx = make_docx(contact)
    website = make_website()
    source = make_zip()
    print("Generated:")
    for path in [contact, docx, website, source]:
        print(path.relative_to(ROOT), path.stat().st_size)
    print("Scene SVGs:", len(list(SCENES_DIR.glob("*.svg"))))


if __name__ == "__main__":
    main()
