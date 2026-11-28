#!/usr/bin/env python3
"""THROWAWAY prototype: three 9:16 compositions for a representative Scene 4B frame."""
from pathlib import Path
import importlib.util
from xml.etree import ElementTree as ET

HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("build_storyboard", HERE / "build_storyboard.py")
board = importlib.util.module_from_spec(spec)
spec.loader.exec_module(board)
ET.register_namespace("", board.SVG_NS)

W, H = 1080, 1920
CREAM, PLUM, ESPRESSO = board.CREAM, board.PLUM, board.ESPRESSO
PEACH, TERRACOTTA, SAGE = board.PEACH, board.TERRACOTTA, board.SAGE


def frame(variant, name, art_y, art_h, cards):
    root = ET.Element(board.tag("svg"), {
        "width": str(W), "height": str(H), "viewBox": f"0 0 {W} {H}",
        "role": "img", "data-prototype-variant": variant,
    })
    board.add(root, "title").text = f"Prototype {variant}: {name}"
    board.add(root, "desc").text = "Portrait 9:16 adaptation study using the existing Co-workers Storyset SVG."
    board.rect(root, 0, 0, W, H, CREAM)
    board.rect(root, 0, 0, W, 20, PLUM)
    art = board.add(root, "g", id=f"variant-{variant}-existing-storyset-art")
    board.embed_asset(art, "coworkers", 0, art_y, W, art_h,
                      omit_ids={"Chat"}, color_map={"#407BFF": PLUM}, raka_expression=True,
                      child_omit_indices={"character-2": (52,)})
    for item in cards:
        y, label, lines, fill, text_color = item
        board.rect(root, 52, y + 9, 976, 204, ESPRESSO, rx=28, opacity=0.12)
        board.rect(root, 52, y, 976, 204, fill, rx=28,
                   stroke=PLUM if fill == CREAM else fill, sw=3 if fill == CREAM else 0)
        board.txt(root, 88, y + 44, label, 18, TERRACOTTA if label == "RAKA" else PLUM, 700, letter=1.2)
        board.txt(root, 88, y + 104, lines[0], 29, text_color, 600)
        board.txt(root, 88, y + 150, lines[1], 29, text_color, 600)
    return root


raka = ["Kamu selalu bikin aku", "merasa salah!"]
dina = ["Aku cuma kasih masukan.", "Jangan baper terus, deh."]

# A puts the two dialogue beats first, then lets the meeting illustration resolve the moment.
a = frame("A", "Dialogue First", 550, 1050, [
    (110, "RAKA", raka, PLUM, CREAM),
    (340, "DINA", dina, CREAM, PLUM),
])

# B lets the characters establish the context before the two readable dialogue cards.
b = frame("B", "Scene First", 90, 1100, [
    (1220, "RAKA", raka, PLUM, CREAM),
    (1450, "DINA", dina, CREAM, PLUM),
])

# C builds a central conversation: a smaller full scene anchors the frame between two large replies.
c = frame("C", "Conversation Column", 415, 890, [
    (96, "RAKA", raka, PLUM, CREAM),
    (1456, "DINA", dina, CREAM, PLUM),
])

for key, root in (("A", a), ("B", b), ("C", c)):
    ET.indent(root, space="  ")
    (HERE / f"portrait-prototype-{key}.svg").write_bytes(
        ET.tostring(root, encoding="utf-8", xml_declaration=True)
    )

page = '''<!doctype html><html lang="id"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Portrait storyboard prototype</title><style>*{box-sizing:border-box}body{margin:0;background:#261a2a;color:#fff6ec;font:15px/1.45 system-ui,sans-serif;display:grid;place-items:center;min-height:100vh;padding:24px 24px 88px}.wrap{width:min(100%,430px)}.note{margin:0 0 12px;color:#ffc4a3;font-size:12px;letter-spacing:.12em;text-transform:uppercase}.title{margin:0 0 14px;font-size:22px}.frame{width:100%;aspect-ratio:9/16;background:#fff6ec;border-radius:22px;overflow:hidden;box-shadow:0 18px 55px #0008}.frame img{display:block;width:100%;height:100%}.switch{position:fixed;bottom:18px;left:50%;transform:translateX(-50%);display:flex;align-items:center;gap:16px;padding:10px 14px;background:#fff6ec;color:#2a1f2d;border-radius:999px;box-shadow:0 8px 28px #0007}.switch button{border:0;border-radius:50%;width:36px;height:36px;background:#3d2645;color:#fff6ec;font-size:20px;cursor:pointer}.switch strong{min-width:180px;text-align:center;font-size:13px}</style><main class="wrap"><p class="note">Prototype · 9:16 · Scene 4B</p><h1 class="title">Three portrait compositions</h1><div class="frame"><img id="art" alt="Scene 4B portrait composition"></div></main><nav class="switch" aria-label="Prototype variants"><button id="prev" aria-label="Previous variant">‹</button><strong id="label"></strong><button id="next" aria-label="Next variant">›</button></nav><script>const variants=[['A','Dialogue First'],['B','Scene First'],['C','Conversation Column']];let i=Math.max(0,variants.findIndex(x=>x[0]===new URLSearchParams(location.search).get('variant')));function show(){const [key,name]=variants[i];document.getElementById('art').src=`portrait-prototype-${key}.svg`;document.getElementById('label').textContent=`${key} · ${name}`;const u=new URL(location);u.searchParams.set('variant',key);history.replaceState({},'',u)}function step(d){i=(i+d+variants.length)%variants.length;show()}document.getElementById('prev').onclick=()=>step(-1);document.getElementById('next').onclick=()=>step(1);onkeydown=e=>{if(e.target.matches('input,textarea,[contenteditable=true]'))return;if(e.key==='ArrowLeft')step(-1);if(e.key==='ArrowRight')step(1)};show();</script></html>'''
(HERE / "portrait-prototype.html").write_text(page, encoding="utf-8")
