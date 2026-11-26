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
PALETTE = [PAPER, BLUE_PALE, BLUE, INK, CORAL, MINT]

ASSETS = {
    "feed": {"file": REPO / "Feed-rafiki.svg", "title": "Feed", "url": "https://storyset.com/illustration/feed/rafiki"},
    "texting": {"file": REPO / "Texting-rafiki.svg", "title": "Texting", "url": "https://storyset.com/illustration/texting/rafiki"},
    "feeling": {"file": REPO / "Feeling Blue-rafiki.svg", "title": "Feeling Blue", "url": "https://storyset.com/illustration/feeling-blue/rafiki"},
    "couple": {"file": REPO / "Couple stress-rafiki.svg", "title": "Couple Stress", "url": "https://storyset.com/illustration/couple-stress/rafiki"},
    "problem": {"file": REPO / "Problem solving-rafiki.svg", "title": "Problem Solving", "url": "https://storyset.com/illustration/problem-solving/rafiki"},
    "document": {"file": REPO / "Online document-rafiki.svg", "title": "Online Document", "url": "https://storyset.com/illustration/online-document/rafiki"},
    "research": {"file": REPO / "Research paper-rafiki.svg", "title": "Research Paper", "url": "https://storyset.com/illustration/research-paper/rafiki"},
    "todo": {"file": ASSETS_DIR / "To-do-list-rafiki.svg", "title": "To do list", "url": "https://storyset.com/illustration/to-do-list/rafiki"},
    "social": {"file": REPO / "Social interaction-rafiki.svg", "title": "Social Interaction", "url": "https://storyset.com/illustration/social-interaction/rafiki"},
}

SCENES = [
    {
        "no": 1, "slug": "the-hook", "time": "00:00–00:30", "title": "The Hook",
        "visual": "Live presenter in medium close-up, surrounded by fast social-comment bubbles; the frame then zooms toward a smartphone.",
        "screen": "Are we using psychology terms correctly?",
        "motion": "Komentar masuk cepat mengitari presenter. Presenter menunjuk kamera; garis zoom mengantar ke layar smartphone.",
        "voice": "Pembuka mempertanyakan kebiasaan memberi label dari satu momen sehari-hari.",
        "assets": [], "roles": ["Frame presenter live dan antarmuka smartphone dibuat sebagai vektor SVG yang bisa diedit."],
        "credits": [], "background": "Warm paper, a blue footage frame, and small coral comment accents.",
    },
    {
        "no": 2, "slug": "what-is-therapy-speak", "time": "00:30–01:05", "title": "What Is Therapy Speak?",
        "visual": "Two keyframes: a seated person scrolls complete social-media comments as the phone moves to center; a separate transition gathers the highlighted terms into a full-screen hashtag cloud.",
        "screen": "Comment sentences stay intact; highlight only the term. Transition to #Toxic · #Narcissist · #Gaslighting · #Trauma · #Avoidant · #Depressed.",
        "motion": "02A · 00:30–00:48: push the phone to center and scroll comment sentences, with only the psychology word highlighted. 02B · 00:48–01:05: words lift from the comments and gather into a centered hashtag cloud that fills the frame.",
        "voice": "Belajar istilah psikologi membantu; menggunakannya tanpa memahami makna dapat membuat kita salah menilai.",
        "assets": ["feed"], "roles": ["Feed-rafiki.svg: konteks orang duduk menggunakan ponsel.", "Ponsel tengah, komentar lengkap, highlight kata, dan cloud hashtag adalah vektor editable."],
        "credits": ["feed"], "background": "02A: pale blue social-feed canvas. 02B: separate pale-blue word-cloud transition with a soft cloud-shaped backdrop.",
        "frames": [
            {"code": "02A", "slug": "02a-comment-scroll", "title": "Phone Zoom + Comment Scroll", "time": "00:30–00:48"},
            {"code": "02B", "slug": "02b-hashtag-cloud", "title": "Hashtag Cloud Transition", "time": "00:48–01:05"},
        ],
    },
    {
        "no": 3, "slug": "normal-emotions-become-labels", "time": "01:05–01:55", "title": "When Normal Emotions Become Labels",
        "visual": "Three equal panels: an unanswered message, sadness after a setback, and a disagreement.",
        "screen": "They’re avoidant! · I’m depressed! · You’re gaslighting me!\nOne behavior ≠ a psychological diagnosis.",
        "motion": "Tiga konteks muncul bergantian. Label tampil sebagai kutipan asumsi, lalu banner konteks menutup triptych.",
        "voice": "Satu situasi tidak menentukan kepribadian; sedih setelah gagal tidak otomatis depresi; beda pendapat tidak otomatis gaslighting.",
        "assets": ["texting", "feeling", "couple"], "roles": ["Texting-rafiki.svg: pesan belum dibalas.", "Feeling Blue-rafiki.svg: emosi sedih.", "Couple stress-rafiki.svg: konteks perbedaan pendapat."],
        "credits": ["texting", "feeling", "couple"], "background": "Three softly tinted panels in the shared blue, cream, and coral palette.",
    },
    {
        "no": 4, "slug": "why-it-matters", "time": "01:55–02:35", "title": "Why Does It Matter?",
        "visual": "A person facing a maze anchors a four-step consequence chain.",
        "screen": "Salah memahami diri · Salah menilai orang lain · Konflik hubungan · Stigma gangguan mental",
        "motion": "Kartu label memicu rangkaian empat kartu dampak; urutannya dibaca kiri ke kanan.",
        "voice": "Label yang keliru dapat mengaburkan emosi sendiri, memicu konflik, dan memperkuat stigma.",
        "assets": ["problem"], "roles": ["Problem solving-rafiki.svg: metafora mencari konteks; empat kartu dampak digambar sebagai vektor."],
        "credits": ["problem"], "background": "Cool pale-blue field with coral domino accents.",
    },
    {
        "no": 5, "slug": "ethics", "time": "02:35–03:20", "title": "The Ethical Side of Psychology",
        "visual": "A digital-document composition supports four concise article cards.",
        "screen": "Pasal 9 · 15 · 28 · 31",
        "motion": "Kartu pasal masuk satu per satu. Pita dasar ilmiah dan kehati-hatian menyatukan keempatnya.",
        "voice": "Pasal 9: dasar ilmiah dan profesional. Pasal 15: hindari dampak buruk. Pasal 28 dan 31: komunikasi publik dan media yang bertanggung jawab.",
        "assets": ["document", "research"], "roles": ["Online document-rafiki.svg: halaman digital.", "Research paper-rafiki.svg: konteks dasar ilmiah.", "Empat kartu pasal dibuat sebagai vektor."],
        "credits": ["document", "research"], "background": "Warm white document page on a pale-lilac digital canvas.",
    },
    {
        "no": 6, "slug": "use-therapy-speak-wisely", "time": "03:20–03:50", "title": "How to Use Therapy Speak Wisely",
        "visual": "A checklist illustration sits beside three practical steps and short Indonesian explanations.",
        "screen": "Understand before you label · Describe, don’t diagnose · Learn, don’t self-diagnose",
        "motion": "Tiga baris checklist dan ikon sederhana muncul satu per satu.",
        "voice": "Pahami istilah, jelaskan perilaku spesifik, dan ingat bahwa konten sosial bukan pengganti asesmen profesional.",
        "assets": ["todo"], "roles": ["To-do list-rafiki.svg: ilustrasi checklist; tiga langkah dan ikon pendamping disusun sebagai vektor."],
        "credits": ["todo"], "background": "Light mint backdrop with blue checklist rows.",
    },
    {
        "no": 7, "slug": "closing", "time": "03:50–04:00", "title": "Closing",
        "visual": "A group scene sits beside four everyday emotion icons and the closing title card.",
        "screen": "THINK BEFORE YOU LABEL.\nUnderstand psychology. Use it responsibly.",
        "motion": "Istilah memudar; ekspresi dan judul penutup muncul. Sisakan slot identitas pembuat.",
        "voice": "Psychology should help us understand people, not reduce them to labels. Stay curious, stay kind.",
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
        span = add(node, "tspan", fill=color, font_weight=700 if bold else 500)
        span.text = value
    return node


def pill(parent, x, y, w, h, label, fill=WHITE, color=INK, size=22, stroke=None, weight=700, opacity=None):
    rect(parent, x, y, w, h, fill, rx=h / 2, stroke=stroke, sw=2, opacity=opacity)
    txt(parent, x + w / 2, y + h * 0.66, label, size, color, weight, "middle", opacity=opacity)


def card(parent, x, y, w, h, fill=WHITE, border="#E3E8ED", radius=22):
    rect(parent, x + 7, y + 10, w, h, "#D7E0EA", rx=radius, opacity=0.42)
    return rect(parent, x, y, w, h, fill, rx=radius, stroke=border, sw=2)


def embed_asset(parent, asset_key, x, y, w, h):
    info = ASSETS[asset_key]
    source = ET.parse(info["file"]).getroot()
    vb = source.get("viewBox", "0 0 500 500")
    nested = add(parent, "svg", x=x, y=y, width=w, height=h, viewBox=vb, preserveAspectRatio="xMidYMid meet", overflow="hidden", aria_label=info["title"] + " illustration, Storyset Rafiki")
    copied = [deepcopy(child) for child in list(source)]
    mapping = {}
    prefix = "s" + re.sub(r"[^a-z0-9]", "", asset_key.lower()) + "-"
    for child in copied:
        for elem in child.iter():
            old = elem.get("id")
            if old:
                mapping[old] = prefix + old
    for child in copied:
        for elem in child.iter():
            for key, value in list(elem.attrib.items()):
                for old, new in mapping.items():
                    value = value.replace("url(#" + old + ")", "url(#" + new + ")")
                    value = value.replace("#" + old, "#" + new)
                elem.set(key, value)
            if elem.text:
                for old, new in mapping.items():
                    elem.text = elem.text.replace("#" + old, "#" + new)
            if elem.get("id") in mapping:
                elem.set("id", mapping[elem.get("id")])
        nested.append(child)
    return nested


def start_scene(scene, bg, accent=BLUE):
    root = ET.Element(tag("svg"), {"width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}", "role": "img", "aria-labelledby": "scene-title scene-desc"})
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
    circle(root, 1812, 920, 8, CORAL, opacity=0.75)
    circle(root, 1844, 920, 8, INK, opacity=0.36)
    txt(root, 92, 66, "THERAPY SPEAK GONE WRONG", 20, INK, 700, letter=1.2)
    txt(root, 92, 94, "PRODUCTION BOARD  /  EDITABLE VECTOR", 13, MUTED, 700, letter=1.4)
    pill(root, 1575, 43, 250, 48, scene["time"], WHITE, BLUE_DARK, 19, stroke="#D8E3F1")
    txt(root, 92, 168, scene["title"], 49 if len(scene["title"]) < 35 else 43, INK, 700)
    line(root, 92, 195, 1828, 195, "#D8E0E7", 2)
    line(root, 92, 1015, 1828, 1015, "#D8E0E7", 2)
    txt(root, 92, 1045, "VISUAL + TEKS LAYAR + GERAK", 13, MUTED, 700, letter=1.1)
    txt(root, 1828, 1045, f"{scene['no']:02d} / 07", 15, INK, 700, "end")
    return root


def build_scene_1(scene):
    root = start_scene(scene, PAPER, BLUE)
    # Presenter frame with a neutral editable stand-in, not a generated portrait.
    card(root, 92, 230, 1062, 595, WHITE, "#C8D4E0", 24)
    rect(root, 92, 230, 1062, 60, INK, rx=24)
    rect(root, 92, 267, 1062, 23, INK)
    circle(root, 125, 260, 9, CORAL)
    circle(root, 155, 260, 9, "#F4C66A")
    circle(root, 185, 260, 9, MINT)
    txt(root, 220, 268, "LIVE FOOTAGE  ·  MEDIUM CLOSE-UP", 17, WHITE, 700, letter=1)
    rect(root, 112, 306, 1022, 492, "#E7F0F8", rx=14)
    # Camera frame marks and a simple placeholder silhouette.
    for x, y, sx, sy in [(145, 336, 1, 1), (1100, 336, -1, 1), (145, 770, 1, -1), (1100, 770, -1, -1)]:
        path(root, f"M{x},{y+32*sy} L{x},{y} L{x+32*sx},{y}", stroke=BLUE, sw=5)
    circle(root, 620, 463, 69, "#F2B89F")
    path(root, "M551,451 C553,373 690,374 691,457 L678,440 C648,456 613,421 579,451 L562,479 Z", fill=INK)
    path(root, "M470,716 C486,614 530,555 620,555 C710,555 755,614 770,716 Z", fill=BLUE)
    path(root, "M538,589 C574,612 667,612 702,589 L725,716 L516,716 Z", fill=BLUE_DARK, opacity=0.35)
    rect(root, 130, 741, 986, 42, WHITE, rx=12, opacity=0.9)
    txt(root, 623, 770, "PLACE PRESENTER VIDEO HERE", 17, MUTED, 700, "middle", letter=1)
    # Floating labels arrive from all sides; the words remain clearly framed as quoted reactions.
    pill(root, 154, 201, 230, 54, "“avoidant?”", WHITE, BLUE_DARK, 23, stroke="#DCE5EF")
    pill(root, 748, 210, 180, 54, "“toxic”", "#FFF1EC", "#A54E3D", 23, stroke="#F5D8CF")
    pill(root, 234, 795, 235, 52, "“depressed?”", "#EEF2FF", BLUE_DARK, 22, stroke="#D7DDF8")
    pill(root, 790, 794, 260, 52, "“narcissistic!”", WHITE, INK, 21, stroke="#DCE5EF")
    # Smartphone transition is a custom vector object.
    path(root, "M1175,500 C1218,500 1217,500 1252,500", stroke=CORAL, sw=8)
    path(root, "M1240,484 L1262,500 L1240,516 Z", fill=CORAL)
    card(root, 1274, 276, 352, 500, INK, INK, 34)
    rect(root, 1290, 302, 320, 438, "#F3F7FD", rx=22)
    rect(root, 1412, 286, 76, 8, "#89939B", rx=4)
    txt(root, 1322, 346, "SOCIAL FEED", 16, MUTED, 700, letter=1)
    for i, (word, col) in enumerate([("avoidant", BLUE), ("toxic", CORAL), ("sad?", BLUE_DARK)]):
        y = 374 + i * 96
        rect(root, 1318, y, 264, 76, WHITE, rx=16, stroke="#E0E6ED")
        circle(root, 1344, y + 28, 11, col)
        rect(root, 1370, y + 19, 172, 12, "#CAD5E0", rx=6)
        txt(root, 1370, y + 53, word, 18, INK, 700)
    txt(root, 1450, 812, "ZOOM → SMARTPHONE", 16, BLUE_DARK, 700, "middle", letter=0.8)
    rect(root, 92, 876, 1736, 93, WHITE, rx=20, stroke="#E1E6EB")
    txt(root, 132, 932, "ON-SCREEN", 14, BLUE_DARK, 700, letter=1.2)
    txt(root, 1320, 936, scene["screen"], 27, INK, 700, "middle")
    return root


def build_scene_2(scene):
    root = start_scene(scene, "#EDF4FC", BLUE)
    txt(root, 118, 253, "ESTABLISHING  ·  SEATED SCROLL", 14, MUTED, 700, letter=1)
    card(root, 82, 270, 560, 575, WHITE, "#D8E3EE", 22)
    embed_asset(root, "feed", 94, 282, 536, 548)
    # The original small phone in the illustration leads into the center close-up.
    path(root, "M625,565 C650,548 664,548 684,548", stroke=CORAL, sw=7)
    path(root, "M671,534 L691,548 L671,562 Z", fill=CORAL)

    # Phone is now centered; complete comments scroll inside the screen.
    card(root, 685, 235, 570, 650, INK, INK, 38)
    rect(root, 704, 266, 532, 586, "#F5F8FC", rx=25)
    rect(root, 899, 246, 142, 8, "#89939B", rx=4)
    txt(root, 740, 310, "KOMENTAR NETIZEN", 16, MUTED, 700, letter=1)
    circle(root, 1195, 304, 5, CORAL)
    comments = [
        ("Chat 3 jam nggak dibales,", [("avoidant", BLUE_DARK, True), (" banget ya?", INK, False)]),
        ("Baru batal sekali,", [("kok udah ", INK, False), ("toxic", CORAL, True), ("?", INK, False)]),
        ("Beda pendapat dikit,", [("langsung ", INK, False), ("gaslighting?", BLUE_DARK, True)]),
        ("Sedih dikit habis gagal,", [("dibilang ", INK, False), ("depressed?", CORAL, True)]),
    ]
    for i, (first_line, second_line) in enumerate(comments):
        y = 335 + i * 119
        rect(root, 724, y, 490, 106, WHITE, rx=15, stroke="#E0E6ED")
        circle(root, 751, y + 30, 11, [BLUE, CORAL, "#8878D2", "#589681"][i])
        rect(root, 777, y + 22, 95, 8, "#CAD5E0", rx=4)
        rect(root, 777, y + 36, 141, 6, "#E1E6EB", rx=3)
        txt(root, 746, y + 68, first_line, 17, INK, 500)
        inline_txt(root, 746, y + 91, second_line, 17)
    # Scroll track and handle make the movement legible in a still frame.
    rect(root, 1221, 374, 5, 390, "#D9E1E9", rx=3)
    rect(root, 1218, 466, 11, 122, BLUE, rx=5)
    txt(root, 969, 838, "SCROLL", 12, MUTED, 700, "middle", letter=1.1)

    card(root, 1304, 306, 510, 498, WHITE, "#D8E3EE", 22)
    pill(root, 1340, 344, 300, 40, "KATA KUNCI DALAM KONTEKS", "#EAF1FF", BLUE_DARK, 13)
    txt(root, 1340, 435, "Komentarnya tetap utuh.", 22, INK, 700)
    txt(root, 1340, 473, "Yang disorot hanya istilahnya.", 18, MUTED, 400)
    rect(root, 1340, 516, 430, 142, "#F4F7FB", rx=15)
    txt(root, 1363, 560, "Beda pendapat dikit,", 20, INK, 500)
    inline_txt(root, 1363, 595, [("langsung ", INK, False), ("gaslighting?", BLUE_DARK, True)], 20)
    line(root, 1364, 620, 1510, 620, BLUE, 4)
    txt(root, 1340, 712, "Kutipan ini contoh percakapan,", 16, MUTED, 400)
    txt(root, 1340, 739, "bukan kesimpulan atau diagnosis.", 16, MUTED, 400)
    rect(root, 92, 919, 1736, 57, INK, rx=17)
    txt(root, 960, 955, "02A  ·  PHONE PUSH-IN  +  FULL-SENTENCE COMMENT SCROLL", 18, WHITE, 700, "middle", letter=0.6)
    return root


def build_scene_2b(scene):
    root = start_scene(scene, "#EAF1FB", BLUE)
    # Each hashtag has its own colored backing; offsets and slight rotations make a layered word cloud.
    txt(root, 165, 384, "#", 88, BLUE, 700, "middle", opacity=0.14)
    txt(root, 1760, 405, "#", 92, CORAL, 700, "middle", opacity=0.18)
    txt(root, 170, 798, "#", 78, "#8878D2", 700, "middle", opacity=0.16)
    txt(root, 1752, 788, "#", 84, BLUE_DARK, 700, "middle", opacity=0.15)
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
    for index, (word, x, y, w, h, fill, color, angle, size) in enumerate(tags):
        y += ((index * 7) % 13) - 6
        group = add(root, "g", transform=f"rotate({angle} {x + w / 2} {y + h / 2})")
        rect(group, x + 5, y + 6, w, h, "#AEBFD1", rx=17, opacity=0.3)
        rect(group, x, y, w, h, fill, rx=17, stroke="#FFFFFF", sw=2)
        txt(group, x + w / 2, y + h * 0.65, word, size, color, 700, "middle")
    rect(root, 92, 919, 1736, 57, INK, rx=17)
    txt(root, 960, 955, "02B  ·  SEPARATE TRANSITION  →  HASHTAGS GATHER AND FILL THE SCREEN", 18, WHITE, 700, "middle", letter=0.3)
    return root


def build_scene_3(scene):
    root = start_scene(scene, "#F7F5EF", CORAL)
    cases = [
        ("texting", "CASE 01  ·  UNANSWERED MESSAGE", "“They’re avoidant!”", "Belum dibalas beberapa jam"),
        ("feeling", "CASE 02  ·  SADNESS AFTER FAILURE", "“I’m depressed!”", "Sedih setelah gagal mencapai target"),
        ("couple", "CASE 03  ·  DISAGREEMENT", "“You’re gaslighting me!”", "Tidak setuju belum tentu gaslighting"),
    ]
    x_positions = [88, 682, 1276]
    accents = [BLUE, CORAL, "#8B79D6"]
    pale = ["#EAF2FE", "#FFF0EA", "#F0EDFF"]
    for i, (key, kicker, quote, context) in enumerate(cases):
        x = x_positions[i]
        card(root, x, 225, 556, 650, WHITE, "#E1E4E8", 22)
        rect(root, x, 225, 556, 58, pale[i], rx=22)
        rect(root, x, 262, 556, 21, pale[i])
        txt(root, x + 26, 262, kicker, 15, accents[i], 700, letter=0.8)
        embed_asset(root, key, x + 58, 282, 440, 360)
        pill(root, x + 34, 650, 488, 70, quote, pale[i], INK, 23, stroke="#E5E8EC")
        txt(root, x + 278, 763, context, 18, MUTED, 400, "middle")
        pill(root, x + 169, 798, 218, 40, "LABEL · ASSUMPTION", WHITE, accents[i], 13, stroke="#E5E8EC")
    rect(root, 88, 900, 1744, 72, INK, rx=18)
    txt(root, 960, 945, "One behavior ≠ a psychological diagnosis.", 29, WHITE, 700, "middle")
    return root


def build_scene_4(scene):
    root = start_scene(scene, "#EAF1F9", BLUE)
    card(root, 80, 250, 650, 620, WHITE, "#D8E2ED", 24)
    embed_asset(root, "problem", 100, 268, 610, 580)
    txt(root, 405, 856, "MENCARI KONTEKS", 15, BLUE_DARK, 700, "middle", letter=1)
    pill(root, 790, 244, 1020, 82, "LABEL TANPA KONTEKS", INK, WHITE, 26)
    path(root, "M1300,328 C1300,365 1300,365 1300,390", stroke=CORAL, sw=7)
    path(root, "M1287,377 L1300,398 L1313,377 Z", fill=CORAL)
    outcomes = [
        ("01", "Salah memahami\ndiri sendiri", BLUE),
        ("02", "Salah menilai\norang lain", "#5B78D6"),
        ("03", "Konflik dalam\nhubungan", CORAL),
        ("04", "Stigma terhadap\ngangguan mental", "#C76A5B"),
    ]
    x0, y, cw, ch, gap = 780, 430, 238, 320, 22
    for i, (number, label, accent) in enumerate(outcomes):
        x = x0 + i * (cw + gap)
        card(root, x, y, cw, ch, WHITE, "#DCE4EC", 20)
        rect(root, x, y, cw, 18, accent, rx=9)
        circle(root, x + 46, y + 75, 25, accent)
        txt(root, x + 46, y + 82, number, 15, WHITE, 700, "middle")
        txt(root, x + 26, y + 151, label, 25, INK, 700)
        txt(root, x + 26, y + 270, "DAMPAK", 14, MUTED, 700, letter=1)
        if i < 3:
            line(root, x + cw + 4, y + 170, x + cw + gap - 3, y + 170, CORAL, 5)
            path(root, f"M{x+cw+gap-11},{y+160} L{x+cw+gap-1},{y+170} L{x+cw+gap-11},{y+180} Z", fill=CORAL)
    return root


def build_scene_5(scene):
    root = start_scene(scene, "#F2F1FA", "#8878D2")
    card(root, 72, 240, 680, 638, "#FBFBFD", "#DDDDED", 26)
    embed_asset(root, "document", 92, 250, 595, 580)
    embed_asset(root, "research", 390, 516, 326, 314)
    pill(root, 125, 832, 555, 42, "KODE ETIK PSIKOLOGI INDONESIA", WHITE, BLUE_DARK, 15, stroke="#E0E4EC")
    articles = [
        ("9", "Dasar ilmiah & profesional", "Keputusan psikologi berlandaskan pengetahuan yang teruji."),
        ("15", "Menghindari dampak buruk", "Sampaikan informasi dengan mempertimbangkan dampaknya."),
        ("28", "Pernyataan publik yang cermat", "Bijaksana, jujur, teliti, hati-hati, dan berbasis ilmiah."),
        ("31", "Komunikasi media bertanggung jawab", "Selaras dengan kode etik dan praktik psikologi yang tepat."),
    ]
    colors = [BLUE, CORAL, "#8878D2", "#589681"]
    for i, (num, title, desc) in enumerate(articles):
        y = 230 + i * 164
        card(root, 792, y, 1032, 145, WHITE, "#E0E3EA", 18)
        circle(root, 858, y + 72, 39, colors[i])
        txt(root, 858, y + 82, num, 27, WHITE, 700, "middle")
        txt(root, 924, y + 60, "PASAL " + num, 14, colors[i], 700, letter=1)
        txt(root, 924, y + 91, title, 24, INK, 700)
        txt(root, 924, y + 121, desc, 17, MUTED, 400)
    pill(root, 792, 906, 1032, 53, "SCIENTIFIC BASIS  +  PROFESSIONAL CARE", INK, WHITE, 19)
    return root


def build_scene_6(scene):
    root = start_scene(scene, "#EFF6F0", "#589681")
    card(root, 72, 235, 660, 660, WHITE, "#DDE7E0", 24)
    embed_asset(root, "todo", 90, 250, 620, 615)
    steps = [
        ("1", "Understand before you label", "Cari tahu makna istilah sebelum memakainya.", BLUE),
        ("2", "Describe, don’t diagnose", "Jelaskan perilaku spesifik yang membuat tidak nyaman.", CORAL),
        ("3", "Learn, don’t self-diagnose", "Konten edukasi bukan pengganti asesmen profesional.", "#589681"),
    ]
    for i, (num, title, desc, accent) in enumerate(steps):
        y = 245 + i * 216
        card(root, 770, y, 1070, 190, WHITE, "#DDE7E0", 20)
        circle(root, 844, y + 75, 37, accent)
        txt(root, 844, y + 85, num, 26, WHITE, 700, "middle")
        txt(root, 910, y + 70, title, 27, INK, 700)
        txt(root, 910, y + 116, desc, 20, MUTED, 400)
        # Small, editable line icons: learn, describe, and know the limit.
        if i == 0:
            circle(root, 1740, y + 73, 25, "none", stroke=accent, sw=5)
            line(root, 1758, y + 91, 1779, y + 112, accent, 6)
        elif i == 1:
            rect(root, 1700, y + 45, 93, 58, "#FFF2EC", rx=12, stroke=accent, sw=3)
            path(root, f"M1720,{y+72} L1734,{y+86} L1769,{y+58}", stroke=accent, sw=5)
        else:
            circle(root, 1744, y + 75, 29, "#E8F5ED", stroke=accent, sw=3)
            path(root, f"M1728,{y+76} L1740,{y+88} L1762,{y+61}", stroke=accent, sw=5)
    return root


def face(parent, cx, cy, emotion, color):
    circle(parent, cx, cy, 38, "#FFF9F3", stroke=color, sw=4)
    circle(parent, cx - 12, cy - 8, 3.5, INK)
    circle(parent, cx + 12, cy - 8, 3.5, INK)
    if emotion == "joy":
        path(parent, f"M{cx-15},{cy+7} Q{cx},{cy+25} {cx+15},{cy+7}", stroke=INK, sw=3)
    elif emotion == "sadness":
        path(parent, f"M{cx-14},{cy+19} Q{cx},{cy+4} {cx+14},{cy+19}", stroke=INK, sw=3)
        path(parent, f"M{cx+22},{cy-1} Q{cx+29},{cy+8} {cx+22},{cy+14} Q{cx+16},{cy+8} {cx+22},{cy-1}", fill=BLUE)
    elif emotion == "fear":
        circle(parent, cx, cy + 13, 5, "none", stroke=INK, sw=3)
    else:
        line(parent, cx - 15, cy + 12, cx + 15, cy + 12, INK, 3)


def build_scene_7(scene):
    root = start_scene(scene, "#FAF1EC", CORAL)
    card(root, 72, 238, 798, 660, WHITE, "#EEDFD8", 25)
    embed_asset(root, "social", 88, 250, 770, 620)
    # Previous viral labels fade out before the human expressions take focus.
    pill(root, 114, 258, 180, 40, "toxic", WHITE, MUTED, 15, opacity=0.6)
    pill(root, 300, 258, 204, 40, "trauma", WHITE, MUTED, 15, opacity=0.6)
    pill(root, 512, 258, 190, 40, "avoidant", WHITE, MUTED, 15, opacity=0.6)
    rect(root, 918, 238, 910, 660, WHITE, rx=26, stroke="#EEDFD8", sw=2)
    txt(root, 1373, 290, "NORMAL HUMAN EMOTIONS", 15, MUTED, 700, "middle", letter=1.2)
    emotions = [("joy", "Gembira", BLUE), ("sadness", "Sedih", CORAL), ("fear", "Takut", "#8878D2"), ("calm", "Tenang", "#589681")]
    for i, (kind, label, color) in enumerate(emotions):
        x = 960 + i * 210
        card(root, x, 316, 190, 150, "#FFFDFC", "#E8E4E0", 16)
        face(root, x + 95, 370, kind, color)
        txt(root, x + 95, 444, label, 17, INK, 700, "middle")
    line(root, 980, 500, 1764, 500, "#E8E4E0", 2)
    txt(root, 1373, 573, "THINK BEFORE", 23, BLUE_DARK, 700, "middle", letter=1.2)
    txt(root, 1373, 638, "YOU LABEL.", 46, INK, 700, "middle", letter=1)
    txt(root, 1373, 692, "Understand psychology. Use it responsibly.", 19, MUTED, 400, "middle")
    pill(root, 1088, 748, 570, 55, "PSYCHOLOGY HELPS US UNDERSTAND PEOPLE", "#EDF3FF", BLUE_DARK, 16)
    txt(root, 1373, 850, "END CARD  ·  VIDEO TITLE  ·  CREATOR IDENTITY", 14, MUTED, 700, "middle", letter=0.7)
    rect(root, 92, 926, 1736, 47, INK, rx=15)
    txt(root, 960, 958, "Stay curious. Stay kind.", 23, WHITE, 700, "middle")
    return root


DRAW = [build_scene_1, build_scene_2, build_scene_3, build_scene_4, build_scene_5, build_scene_6, build_scene_7]


def create_previews():
    PREVIEW_DIR.mkdir(parents=True, exist_ok=True)
    SCENES_DIR.mkdir(parents=True, exist_ok=True)
    for stale in [
        SCENES_DIR / "scene-02-what-is-therapy-speak.svg",
        PREVIEW_DIR / "scene-02-what-is-therapy-speak.png",
    ]:
        if stale.exists():
            stale.unlink()
    thumbs = []
    for scene, draw_scene in zip(SCENES, DRAW):
        if scene["no"] == 2:
            frames = [
                ("02A", "02a-comment-scroll", "Phone Zoom + Comment Scroll", "00:30–00:48", build_scene_2),
                ("02B", "02b-hashtag-cloud", "Hashtag Cloud Transition", "00:48–01:05", build_scene_2b),
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
    # Three columns make a compact visual audit sheet while retaining the 16:9 frames.
    card_w, card_h, gap, top = 420, 236, 26, 56
    cols, rows = 4, 2
    sheet = Image.new("RGB", (cols * card_w + (cols + 1) * gap, rows * (card_h + 48) + (rows + 1) * gap + top), "#F1F4F8")
    draw = ImageDraw.Draw(sheet)
    title_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf", 22)
    label_font = ImageFont.truetype("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", 13)
    draw.text((gap, 15), "STORYBOARD C · VISUAL AUDIT · 7 SCENES / 8 FRAMES", font=title_font, fill="#263238")
    for i, (code, title, png_path) in enumerate(thumbs):
        col, row = i % cols, i // cols
        x = gap + col * (card_w + gap)
        y = top + gap + row * (card_h + 48 + gap)
        im = Image.open(png_path).convert("RGB").resize((card_w, card_h), Image.Resampling.LANCZOS)
        sheet.paste(im, (x, y))
        draw.text((x, y + card_h + 9), f"{code}  ·  {title}", font=label_font, fill="#263238")
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
    color.set(qn("w:val"), "2859BE")
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
    normal.font.color.rgb = RGBColor(38, 50, 56)
    normal.paragraph_format.space_after = Pt(4)

    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("THERAPY SPEAK GONE WRONG")
    set_run(r, 28, INK, True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(7)
    r = p.add_run("Storyboard produksi · Layout C · 7 scene · 00:00–04:00")
    set_run(r, 15, BLUE_DARK, True)
    p = doc.add_paragraph("Setiap frame tersedia sebagai SVG mandiri yang editable. Scene 2 memakai dua keyframe: guliran komentar utuh lalu transisi cloud hashtag. Ilustrasi Storyset ditanam sebagai vektor; background, label, chip, kartu pasal, dan anotasi digambar langsung sebagai bentuk SVG. Tidak ada gambar generatif.")
    p.paragraph_format.space_after = Pt(10)
    for r in p.runs:
        set_run(r, 10, MUTED)

    heading = doc.add_paragraph()
    heading.paragraph_format.space_before = Pt(4)
    r = heading.add_run("AUDIT VISUAL")
    set_run(r, 14, INK, True)
    audit = [
        "Sembilan ilustrasi yang digunakan mempertahankan bahasa flat vector, garis tipis, bayangan lembut, warna biru #407BFF dan tinta gelap #263238.",
        "Warna latar berubah menurut fungsi scene, tetapi tiap frame mengulang grid, tipografi sans-serif, kapsul waktu, dan aksen biru-koral agar tetap satu rangkaian.",
        "Caption pada scene 3 dibingkai sebagai kutipan asumsi; ilustrasi karakter tidak ditandai sebagai diagnosis atau pihak yang bersalah.",
        "Frame presenter adalah slot footage, bukan potret sintetis. Seluruh elemen baru berupa geometri dan ikon SVG yang dapat diedit.",
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
    for color, name in zip(PALETTE, ["paper", "pale blue", "Rafiki blue", "ink", "coral", "mint"]):
        r = p.add_run("   " + name + "   ")
        set_run(r, 9, color, True)
        r.font.highlight_color = None
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    r = p.add_run("Kredit Storyset: Freepik License, free with attribution. Tautan unduh per ilustrasi ada di scene terkait. Periksa ketentuan Storyset saat penggunaan final.")
    set_run(r, 9, MUTED)
    add_hyperlink(p, "Storyset terms", "https://storyset.com/terms")

    doc.add_page_break()
    p = doc.add_paragraph()
    r = p.add_run("CONTACT SHEET · AUDIT 7 SCENE / 8 FRAME")
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
        set_run(r, 11, BLUE_DARK, True)
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
        shade(left, "F1F4F8")
        shade(right, "F8FAFC")
        set_cell_margins(left, 100, 100, 100, 100)
        set_cell_margins(right, 100, 140, 100, 140)
        if scene["no"] == 2:
            left.paragraphs[0].text = ""
            inner = left.add_table(rows=1, cols=2)
            inner.autofit = False
            inner.columns[0].width = Inches(3.08)
            inner.columns[1].width = Inches(3.08)
            frame_data = [
                ("02A · 00:30–00:48", "scene-02a-comment-scroll.png"),
                ("02B · 00:48–01:05", "scene-02b-hashtag-cloud.png"),
            ]
            for cell, (label, filename) in zip(inner.rows[0].cells, frame_data):
                cell.width = Inches(3.08)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                caption = cell.paragraphs[0]
                caption.alignment = 1
                caption.paragraph_format.space_after = Pt(5)
                run = caption.add_run(label)
                set_run(run, 8, BLUE_DARK, True)
                image_para = cell.add_paragraph()
                image_para.alignment = 1
                image_para.add_run().add_picture(str(PREVIEW_DIR / filename), width=Inches(2.92))
        else:
            image = PREVIEW_DIR / f"scene-{scene['no']:02d}-{scene['slug']}.png"
            p = left.paragraphs[0]
            p.alignment = 1
            p.add_run().add_picture(str(image), width=Inches(6.2))
        p = right.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("VISUAL")
        set_run(r, 8, BLUE_DARK, True)
        p = right.add_paragraph(scene["visual"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ON-SCREEN TEXT")
        set_run(r, 8, BLUE_DARK, True)
        p = right.add_paragraph(scene["screen"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK, True)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("GERAK / KAMERA")
        set_run(r, 8, BLUE_DARK, True)
        p = right.add_paragraph(scene["motion"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("VOICE-OVER CUE")
        set_run(r, 8, BLUE_DARK, True)
        p = right.add_paragraph(scene["voice"])
        p.paragraph_format.space_after = Pt(5)
        for r in p.runs: set_run(r, 9, INK)
        p = right.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run("ASET & PERAN")
        set_run(r, 8, BLUE_DARK, True)
        for role in scene["roles"]:
            p = right.add_paragraph(role, style="List Bullet")
            p.paragraph_format.space_after = Pt(1)
            for r in p.runs: set_run(r, 8, INK)
        for asset_key in scene["credits"]:
            info = ASSETS[asset_key]
            p = right.add_paragraph()
            p.paragraph_format.space_after = Pt(1)
            add_hyperlink(p, "Unduh " + info["title"] + " · Storyset Rafiki", info["url"])
            for r in p.runs: set_run(r, 8, BLUE_DARK)
        p = right.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        r = p.add_run("BACKGROUND: ")
        set_run(r, 8, BLUE_DARK, True)
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
              <section class="frame-card"><div class="frame-label"><b>02A · PHONE + COMMENT SCROLL</b><span>00:30–00:48</span></div><a href="scenes/scene-02a-comment-scroll.svg" target="_blank" aria-label="Open editable SVG for Scene 2A"><img src="scenes/scene-02a-comment-scroll.svg" alt="Person scrolling complete social-media comments; highlighted psychology terms remain inside each sentence"></a><div class="file-links"><a href="scenes/scene-02a-comment-scroll.svg" download>Download editable SVG</a><a href="previews/scene-02a-comment-scroll.png" download>Download PNG preview</a></div></section>
              <section class="frame-card"><div class="frame-label"><b>02B · HASHTAG CLOUD</b><span>00:48–01:05</span></div><a href="scenes/scene-02b-hashtag-cloud.svg" target="_blank" aria-label="Open editable SVG for Scene 2B"><img src="scenes/scene-02b-hashtag-cloud.svg" alt="Centered hashtag word cloud fills the transition frame"></a><div class="file-links"><a href="scenes/scene-02b-hashtag-cloud.svg" download>Download editable SVG</a><a href="previews/scene-02b-hashtag-cloud.png" download>Download PNG preview</a></div></section>
            </div>'''
        else:
            svg_name = f"scene-{scene['no']:02d}-{scene['slug']}.svg"
            png_name = f"scene-{scene['no']:02d}-{scene['slug']}.png"
            frame_markup = f'<a href="scenes/{svg_name}" target="_blank" aria-label="Open editable SVG for scene {scene["no"]}"><img src="scenes/{svg_name}" alt="Editable scene illustration, scene {scene["no"]}"></a><div class="file-links"><a href="scenes/{svg_name}" download>Download editable SVG</a><a href="previews/{png_name}" download>Download PNG preview</a></div>'
        rows.append(f'''<article class="board-row">
          <header class="row-meta"><span class="scene-number">{scene['no']:02d}</span><span class="time">{scene['time']}</span><h2>{scene['title']}</h2><p>{scene['visual']}</p></header>
          <div class="row-frame">{frame_markup}</div>
          <dl class="row-notes"><div><dt>On-screen text</dt><dd>{scene['screen'].replace(chr(10), '<br>')}</dd></div><div><dt>Motion / camera</dt><dd>{scene['motion']}</dd></div><div><dt>Voice-over cue</dt><dd>{scene['voice']}</dd></div><div><dt>Assets & roles</dt><dd>{role_text}</dd><p class="asset-sources">{' · '.join(asset_links) if asset_links else 'Custom editable SVG vectors; no external illustration.'}</p></div></dl>
        </article>''')
    html = '''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="Editable production board and seven SVG story frames for Therapy Speak Gone Wrong.">
<title>Therapy Speak Gone Wrong · Production Board</title>
<style>
:root{--ink:#263238;--blue:#407bff;--blue-dark:#2859be;--paper:#f5f6f5;--line:#dce2e7;--muted:#687883;--coral:#ed8067;--white:#fff;--shadow:0 12px 30px rgba(30,47,63,.07)}
*{box-sizing:border-box}html{background:var(--paper);color:var(--ink);font-family:Arial,Helvetica,sans-serif}body{margin:0;padding:42px 28px 92px}.shell{max-width:1560px;margin:0 auto}.hero{display:grid;grid-template-columns:1fr auto;gap:32px;align-items:end;margin:8px 0 32px}.eyebrow{color:var(--blue-dark);font-size:12px;font-weight:700;letter-spacing:.14em;text-transform:uppercase}.hero h1{max-width:920px;margin:12px 0 12px;font-size:clamp(38px,5.5vw,76px);line-height:.98;letter-spacing:-.055em}.hero p{max-width:760px;margin:0;color:var(--muted);font-size:16px;line-height:1.55}.hero-actions{display:flex;flex-direction:column;gap:9px;align-items:stretch}.action{display:inline-block;padding:12px 16px;border:1px solid var(--line);border-radius:999px;background:var(--white);color:var(--ink);font-size:13px;font-weight:700;text-align:center;text-decoration:none}.action.primary{border-color:var(--blue);background:var(--blue);color:var(--white)}.style-strip{display:flex;flex-wrap:wrap;align-items:center;gap:10px;margin:28px 0 18px;padding:15px 18px;border:1px solid var(--line);border-radius:14px;background:var(--white);color:var(--muted);font-size:12px}.swatch{display:inline-flex;align-items:center;gap:6px}.swatch i{width:15px;height:15px;border:1px solid #d5dce2;border-radius:50%}.board{display:grid;gap:13px}.board-row{display:grid;grid-template-columns:185px minmax(340px,.98fr) minmax(420px,1.08fr);gap:18px;align-items:stretch;padding:15px;border:1px solid var(--line);border-radius:17px;background:var(--white);box-shadow:var(--shadow)}.row-meta{display:flex;flex-direction:column;align-items:flex-start;padding:5px 0}.scene-number{display:grid;place-items:center;width:34px;height:34px;border-radius:50%;background:var(--ink);color:white;font-size:12px;font-weight:700}.time{margin:12px 0 9px;color:var(--blue-dark);font-size:11px;font-weight:700;letter-spacing:.08em}.row-meta h2{margin:0;font-size:19px;line-height:1.12;letter-spacing:-.02em}.row-meta p{margin:12px 0 0;color:var(--muted);font-size:12px;line-height:1.45}.row-frame{min-width:0;display:flex;flex-direction:column;justify-content:center}.row-frame img{display:block;width:100%;border-radius:10px;background:#edf1f6}.frame-stack{display:grid;gap:12px}.frame-card{padding:9px;border:1px solid var(--line);border-radius:12px;background:#f9fbfd}.frame-label{display:flex;justify-content:space-between;gap:8px;margin-bottom:7px;color:var(--blue-dark);font-size:10px;letter-spacing:.04em}.frame-label span{color:var(--muted);font-weight:400}.file-links{display:flex;justify-content:space-between;gap:10px;padding:9px 2px 0}.file-links a,.asset-sources a{color:var(--blue-dark);font-size:11px;font-weight:700;text-decoration:none}.file-links a:hover,.asset-sources a:hover{text-decoration:underline}.row-notes{display:grid;grid-template-columns:1fr 1fr;gap:10px 14px;align-content:center;margin:0;padding:2px 4px}.row-notes div{min-width:0}.row-notes dt{margin-bottom:4px;color:var(--blue-dark);font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase}.row-notes dd{margin:0;color:#50616d;font-size:11px;line-height:1.4}.asset-sources{margin:7px 0 0}.asset-sources a{margin-right:6px}.audit{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:28px 0;padding:20px;border:1px solid var(--line);border-radius:16px;background:#edf2f8}.audit h2{margin:0 0 8px;font-size:16px}.audit p{margin:0;color:#566974;font-size:12px;line-height:1.5}.footer{margin-top:24px;padding-top:16px;border-top:1px solid var(--line);color:var(--muted);font-size:11px;line-height:1.5}.footer a{color:var(--blue-dark)}
@media(max-width:1100px){.board-row{grid-template-columns:145px minmax(300px,1fr)}.row-notes{grid-column:2;grid-template-columns:1fr 1fr}.hero{grid-template-columns:1fr}.hero-actions{flex-direction:row;justify-content:flex-start}.action{min-width:180px}}
@media(max-width:700px){body{padding:24px 12px 60px}.hero h1{font-size:42px}.board-row{grid-template-columns:1fr;gap:12px;padding:13px}.row-meta{display:grid;grid-template-columns:40px 1fr;column-gap:8px;align-items:center}.scene-number{grid-row:span 2}.time{margin:0 0 4px}.row-meta h2{grid-column:2}.row-meta p{grid-column:1/-1;margin-top:10px}.row-notes{grid-column:auto;grid-template-columns:1fr 1fr}.hero-actions{flex-wrap:wrap}.action{min-width:0}.audit{grid-template-columns:1fr}}
</style>
</head>
<body><main class="shell">
<header class="hero"><div><div class="eyebrow">Editable production board · Layout C</div><h1>Therapy Speak Gone Wrong</h1><p>Seven scenes from the supplied script, arranged as one production row per scene. Every main frame is a standalone editable SVG; the Word board carries the same copy, motion notes, visual roles, and asset credits.</p></div><div class="hero-actions"><a class="action primary" href="Storyboard_Therapy_Speak_Gone_Wrong.docx" download>Download Word storyboard</a><a class="action" href="Storyboard_Editable_Source.zip" download>Download editable source bundle</a></div></header>
<div class="style-strip"><strong>Shared visual system</strong><span class="swatch"><i style="background:#f8f7f2"></i>paper</span><span class="swatch"><i style="background:#e6f0fc"></i>pale blue</span><span class="swatch"><i style="background:#407bff"></i>blue</span><span class="swatch"><i style="background:#263238"></i>ink</span><span class="swatch"><i style="background:#ed8067"></i>coral</span><span class="swatch"><i style="background:#b9d8ca"></i>mint</span><span>· flat vectors, restrained shadows, repeated time and scene labels</span></div>
<section class="board" aria-label="Seven scene production board">
''' + "\n".join(rows) + '''
</section>
<section class="audit"><div><h2>Visual audit</h2><p>Storyset illustrations share the same flat Rafiki vector language, blue #407BFF, dark ink #263238, warm neutrals, and soft shadows. Custom backgrounds and labels use that palette; each scene changes its composition to match its action.</p></div><div><h2>Editing & attribution</h2><p>Open or download each SVG to edit its vector layers and text. The source bundle contains the HTML board and all scene SVGs. Storyset art is free with attribution; use the per-scene links and review <a href="https://storyset.com/terms" target="_blank" rel="noreferrer">Storyset terms</a> for final use.</p></div></section>
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
