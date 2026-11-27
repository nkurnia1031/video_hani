#!/usr/bin/env python3
"""Build an editable, layered Raka SVG rig from the credited Co-workers source."""

from copy import deepcopy
import json
from pathlib import Path
from xml.etree import ElementTree as ET

import cairosvg


ROOT = Path(__file__).resolve().parent
SOURCE = ROOT.parent / "assets" / "Co-workers_Rafiki.svg"
OUT = ROOT / "Raka_Rig.svg"
PREVIEW = ROOT / "Raka_Rig_Preview.png"
MANIFEST = ROOT / "Raka_Rig.json"
SVG = "http://www.w3.org/2000/svg"
NS = {"s": SVG}
PLUM = "#3D2645"
PEACH = "#FFC3BD"
ESPRESSO = "#2A1F2D"

RIG_PACKAGE = {
    "character_id": "raka",
    "rig_type": "svg_rig",
    "version": 1,
    "coordinate_space": {"viewBox": "170 140 280 290", "units": "svg-user-units"},
    "side_convention": "left/right mean viewer-left/viewer-right in the neutral SVG; SVG x increases toward viewer-right.",
    "source": "https://storyset.com/illustration/co-workers/rafiki",
    "parts": [
        {"id": "pelvis", "kind": "skeleton-root", "parent": "raka-rig", "layer": 5},
        {"id": "rig-legs", "kind": "group", "parent": "pelvis", "layer": 10},
        {"id": "left-hip", "kind": "thigh", "parent": "rig-legs", "layer": 11, "joint": "leftHip"},
        {"id": "left-knee", "kind": "calf", "parent": "left-hip", "layer": 12, "joint": "leftKnee"},
        {"id": "left-ankle", "kind": "foot", "parent": "left-knee", "layer": 13, "joint": "leftAnkle"},
        {"id": "right-hip", "kind": "thigh", "parent": "rig-legs", "layer": 14, "joint": "rightHip"},
        {"id": "right-knee", "kind": "calf", "parent": "right-hip", "layer": 15, "joint": "rightKnee"},
        {"id": "right-ankle", "kind": "foot", "parent": "right-knee", "layer": 16, "joint": "rightAnkle"},
        {"id": "upper-body", "kind": "torso-root", "parent": "pelvis", "layer": 20, "joint": "torso"},
        {"id": "torso", "kind": "shirt", "parent": "upper-body", "layer": 21},
        {"id": "left-shoulder", "kind": "upper-arm", "parent": "upper-body", "layer": 30, "joint": "leftShoulder"},
        {"id": "left-elbow", "kind": "forearm", "parent": "left-shoulder", "layer": 31, "joint": "leftElbow"},
        {"id": "left-wrist", "kind": "hand", "parent": "left-elbow", "layer": 32, "joint": "leftWrist"},
        {"id": "right-shoulder", "kind": "upper-arm", "parent": "upper-body", "layer": 33, "joint": "rightShoulder"},
        {"id": "right-elbow", "kind": "forearm", "parent": "right-shoulder", "layer": 34, "joint": "rightElbow"},
        {"id": "right-wrist", "kind": "hand", "parent": "right-elbow", "layer": 35, "joint": "rightWrist"},
        {"id": "source-sleeves", "kind": "shirt-sleeves", "parent": "upper-body", "layer": 40},
        {"id": "head", "kind": "head", "parent": "upper-body", "layer": 60, "joint": "head"},
        {"id": "neck", "kind": "neck", "parent": "head", "layer": 50},
    ],
    "joints": {
        "torso": {"id": "upper-body", "parent": "pelvis", "pivot": [340, 286], "rotation": [-12, 12], "label": "Torso"},
        "head": {"id": "head", "parent": "upper-body", "pivot": [344, 205], "rotation": [-25, 25], "label": "Kepala"},
        "leftShoulder": {"id": "left-shoulder", "parent": "upper-body", "pivot": [316, 238], "rotation": [-90, 115], "label": "Bahu kiri"},
        "leftElbow": {"id": "left-elbow", "parent": "left-shoulder", "pivot": [302, 269], "rotation": [-115, 120], "label": "Siku kiri"},
        "leftWrist": {"id": "left-wrist", "parent": "left-elbow", "pivot": [300, 304], "rotation": [-45, 45], "label": "Pergelangan kiri"},
        "rightShoulder": {"id": "right-shoulder", "parent": "upper-body", "pivot": [365, 238], "rotation": [-115, 90], "label": "Bahu kanan"},
        "rightElbow": {"id": "right-elbow", "parent": "right-shoulder", "pivot": [374, 269], "rotation": [-115, 120], "label": "Siku kanan"},
        "rightWrist": {"id": "right-wrist", "parent": "right-elbow", "pivot": [382, 302], "rotation": [-45, 45], "label": "Pergelangan kanan"},
        "leftHip": {"id": "left-hip", "parent": "rig-legs", "pivot": [316, 286], "rotation": [-25, 25], "label": "Pinggul kiri"},
        "leftKnee": {"id": "left-knee", "parent": "left-hip", "pivot": [311, 350], "rotation": [0, 70], "label": "Lutut kiri"},
        "leftAnkle": {"id": "left-ankle", "parent": "left-knee", "pivot": [307, 391], "rotation": [-25, 25], "label": "Pergelangan kaki kiri"},
        "rightHip": {"id": "right-hip", "parent": "rig-legs", "pivot": [349, 286], "rotation": [-25, 25], "label": "Pinggul kanan"},
        "rightKnee": {"id": "right-knee", "parent": "right-hip", "pivot": [349, 350], "rotation": [-70, 0], "label": "Lutut kanan"},
        "rightAnkle": {"id": "right-ankle", "parent": "right-knee", "pivot": [359, 390], "rotation": [-25, 25], "label": "Pergelangan kaki kanan"},
    },
    "presets": {
        "neutral": {},
        "wave": {"rightShoulder": -55, "rightElbow": -75, "rightWrist": 0, "head": -4},
        "arms-up": {"leftShoulder": 115, "leftElbow": -10, "rightShoulder": -105, "rightElbow": 10},
        "open": {"leftShoulder": 85, "leftElbow": 32, "rightShoulder": -75, "rightElbow": -26},
    },
}

ET.register_namespace("", SVG)


def tag(name):
    return f"{{{SVG}}}{name}"


def add(parent, name, **attrs):
    node = ET.SubElement(parent, tag(name), {k.replace("_", "-"): str(v) for k, v in attrs.items()})
    return node


def add_path(parent, element_id, d, fill, opacity=None, soft_edge=False):
    attrs = {"id": element_id, "d": d, "fill": fill}
    if opacity is not None:
        attrs["opacity"] = str(opacity)
    if soft_edge:
        attrs.update({"stroke": fill, "stroke_width": 1.5, "stroke_linejoin": "round", "stroke_linecap": "round"})
    return add(parent, "path", **attrs)


def add_joint_pad(parent, element_id, pivot, radius, fill):
    x, y = pivot
    return add(parent, "circle", id=element_id, cx=x, cy=y, r=radius, fill=fill, data_joint_pad="true")


def add_joint_group(parent, element_id, pivot, role):
    x, y = pivot
    spec = next((item for item in RIG_PACKAGE["joints"].values() if item["id"] == element_id), None)
    attrs = {
        "id": element_id,
        "data_joint": role,
        "data_pivot_x": x,
        "data_pivot_y": y,
        "transform": f"rotate(0 {x} {y})",
    }
    if spec:
        if spec["pivot"] != [x, y]:
            raise ValueError(f"Pivot mismatch for {element_id}: SVG={pivot} manifest={spec['pivot']}")
        attrs["data_parent"] = spec["parent"]
        attrs["data_rotation_min"] = spec["rotation"][0]
        attrs["data_rotation_max"] = spec["rotation"][1]
    return add(parent, "g", **attrs)


def build_svg():
    source_root = ET.parse(SOURCE).getroot()
    source_character = source_root.find('.//s:g[@id="character-2"]', NS)
    source_parts = list(source_character)

    root = ET.Element(tag("svg"), {
        "width": "920", "height": "953", "viewBox": "170 140 280 290",
        "role": "img", "aria-labelledby": "raka-title raka-desc",
    })
    add(root, "title", id="raka-title").text = "Raka · Editable Character Rig"
    add(root, "desc", id="raka-desc").text = (
        "A standing, cup-free Raka assembled from the Co-workers · Rafiki source character. "
        "Head, shoulders, elbows, wrists, hips, knees, and ankles are independent pivot groups."
    )
    add(root, "metadata").text = (
        "Source: Storyset Co-workers · Rafiki, character-2. "
        "The chair and blue cup are excluded. Original head, torso, sleeves, socks, cuffs, and shoes are retained; "
        "the source seated legs and arm gestures are replaced by jointed standing segments. "
        "Of the 42 person drawables (excluding chair and cup props), 14 (33.3%) are adapted/recolored and 28 (66.7%) remain unchanged. "
        "New limb bridges are drawn in the same flat style. Free with attribution: https://storyset.com/illustration/co-workers/rafiki"
    )
    style = add(root, "style")
    style.text = """
      .joint-guide { fill: #F2B134; stroke: #2A1F2D; stroke-width: 1.5; }
      #rig-guides { opacity: 0; pointer-events: none; }
    """

    rig = add(root, "g", id="raka-rig", data_rig="cutout-character", data_source="co-workers-rafiki-character-2")

    def part(index, element_id, parent):
        source_part = deepcopy(source_parts[index])
        source_part.set("id", element_id)
        source_part.set("data-source-child", str(index))
        for key, value in list(source_part.attrib.items()):
            source_part.set(key, value.replace("#407BFF", PLUM))
        parent.append(source_part)
        return source_part

    pelvis = add(rig, "g", id="pelvis", data_rig_part="pelvis", data_layer=5)

    # Legs are behind the shirt. Each thigh, calf and foot has its own pivot.
    legs = add(pelvis, "g", id="rig-legs", data_layer=10)

    left_hip = add_joint_group(legs, "left-hip", (316, 286), "hip-left")
    left_hip.set("data-layer", "11")
    add_path(left_hip, "left-thigh", "M309 281 C303 300 300 328 300 350 L322 352 C326 330 331 304 336 282 Z", PLUM, soft_edge=True)
    add_path(left_hip, "left-thigh-shade", "M307 300 C304 316 303 335 304 348 L311 349 C310 329 315 308 319 291 Z", ESPRESSO, .18)
    left_knee = add_joint_group(left_hip, "left-knee", (311, 350), "knee-left")
    left_knee.set("data-layer", "12")
    add_path(left_knee, "left-calf", "M301 344 C299 360 297 378 295 390 L317 395 C322 378 323 360 323 346 Z", PLUM, soft_edge=True)
    add_path(left_knee, "left-calf-shade", "M307 352 C304 367 302 381 301 390 L307 392 C312 377 313 363 313 350 Z", ESPRESSO, .16)
    add_joint_pad(left_knee, "left-knee-pad", (311, 350), 8, PLUM)
    left_ankle = add_joint_group(left_knee, "left-ankle", (307, 391), "ankle-left")
    left_ankle.set("data-layer", "13")
    add_joint_pad(left_ankle, "left-ankle-pad", (307, 391), 4, PLUM)
    left_foot = add(left_ankle, "g", id="left-foot", transform="translate(60 0)", data_joint="foot-left")
    for index, name in [(11, "left-sock"), (16, "left-trouser-cuff"), (12, "left-shoe"),
                        (13, "left-shoe-shade"), (17, "left-lace-a"), (18, "left-lace-b")]:
        part(index, name, left_foot)

    right_hip = add_joint_group(legs, "right-hip", (349, 286), "hip-right")
    right_hip.set("data-layer", "14")
    add_path(right_hip, "right-thigh", "M333 282 C337 304 340 329 338 352 L360 350 C361 326 358 301 355 281 Z", PLUM, soft_edge=True)
    add_path(right_hip, "right-thigh-shade", "M343 291 C349 311 351 333 348 349 L356 348 C358 326 355 303 351 286 Z", ESPRESSO, .14)
    right_knee = add_joint_group(right_hip, "right-knee", (349, 350), "knee-right")
    right_knee.set("data-layer", "15")
    add_path(right_knee, "right-calf", "M338 346 C339 363 342 379 346 392 L368 386 C362 370 359 357 360 344 Z", PLUM, soft_edge=True)
    add_path(right_knee, "right-calf-shade", "M350 350 C352 367 356 379 360 388 L367 386 C362 369 359 357 359 347 Z", ESPRESSO, .14)
    add_joint_pad(right_knee, "right-knee-pad", (349, 350), 8, PLUM)
    right_ankle = add_joint_group(right_knee, "right-ankle", (359, 390), "ankle-right")
    right_ankle.set("data-layer", "16")
    add_joint_pad(right_ankle, "right-ankle-pad", (359, 390), 4, PLUM)
    right_foot = add(right_ankle, "g", id="right-foot", transform="translate(48 0)", data_joint="foot-right")
    for index, name in [(24, "right-sock"), (30, "right-trouser-cuff"), (25, "right-shoe"),
                        (28, "right-shoe-shade"), (26, "right-lace-a"), (27, "right-lace-b")]:
        part(index, name, right_foot)

    # Everything above the hips follows one torso pivot; arms and head keep
    # their own child pivots for independent posing.
    upper_body = add_joint_group(pelvis, "upper-body", (340, 286), "torso")
    upper_body.set("data-layer", "20")

    # Original shirt silhouette and its source collar remain intact.
    torso = add(upper_body, "g", id="torso", data_rig_part="torso")
    torso.set("data-layer", "21")
    part(31, "source-tshirt", torso)
    part(32, "source-tshirt-shade", torso)
    part(35, "source-collar", torso)

    # Skin segments are new joint bridges, while original short sleeves are retained.
    left_arm = add_joint_group(upper_body, "left-shoulder", (316, 238), "shoulder-left")
    left_arm.set("data-layer", "30")
    add_joint_pad(left_arm, "left-shoulder-pad", (316, 238), 4, PEACH)
    add_path(left_arm, "left-upper-arm", "M310 239 C304 247 301 257 298 268 L307 274 C314 264 319 252 322 244 Z", PEACH, soft_edge=True)
    left_elbow = add_joint_group(left_arm, "left-elbow", (302, 269), "elbow-left")
    left_elbow.set("data-layer", "31")
    add_joint_pad(left_elbow, "left-elbow-pad", (302, 269), 5, PEACH)
    add_path(left_elbow, "left-forearm", "M298 265 C292 277 292 291 295 303 L305 307 C310 293 309 278 306 269 Z", PEACH, soft_edge=True)
    left_wrist = add_joint_group(left_elbow, "left-wrist", (300, 304), "wrist-left")
    left_wrist.set("data-layer", "32")
    add_joint_pad(left_wrist, "left-wrist-pad", (300, 304), 3, PEACH)
    add_path(left_wrist, "left-hand", "M295 300 C291 299 288 303 288 308 C288 313 291 319 295 320 C298 320 300 317 300 314 C302 318 305 319 307 316 C309 313 305 307 302 303 L300 300 Z", PEACH)
    add_path(left_wrist, "left-thumb", "M299 307 C296 304 293 304 291 306 C293 310 296 312 300 312 Z", PEACH)

    right_arm = add_joint_group(upper_body, "right-shoulder", (365, 238), "shoulder-right")
    right_arm.set("data-layer", "33")
    add_joint_pad(right_arm, "right-shoulder-pad", (365, 238), 4, PEACH)
    add_path(right_arm, "right-upper-arm", "M362 239 C369 247 374 257 377 268 L368 274 C361 264 358 252 356 244 Z", PEACH, soft_edge=True)
    right_elbow = add_joint_group(right_arm, "right-elbow", (374, 269), "elbow-right")
    right_elbow.set("data-layer", "34")
    add_joint_pad(right_elbow, "right-elbow-pad", (374, 269), 5, PEACH)
    add_path(right_elbow, "right-forearm", "M369 267 C373 279 376 293 378 304 L388 303 C389 288 384 274 378 267 Z", PEACH, soft_edge=True)
    right_wrist = add_joint_group(right_elbow, "right-wrist", (382, 302), "wrist-right")
    right_wrist.set("data-layer", "35")
    add_joint_pad(right_wrist, "right-wrist-pad", (382, 302), 3, PEACH)
    add_path(right_wrist, "right-hand", "M378 298 C382 296 388 299 391 303 C394 307 393 313 390 316 C387 319 383 316 381 313 C379 317 375 318 373 315 C371 312 375 305 378 302 Z", PEACH)
    add_path(right_wrist, "right-thumb", "M379 306 C382 302 386 302 389 305 C386 309 383 311 379 311 Z", PEACH)

    # Keep the original shirt sleeves fixed on the torso; only skin segments
    # swing around the shoulder pivots, so the shirt does not detach while Raka waves.
    sleeves = add(upper_body, "g", id="source-sleeves", data_rig_part="sleeves")
    sleeves.set("data-layer", "40")
    part(22, "source-left-sleeve", sleeves)
    part(23, "source-left-sleeve-shade", sleeves)
    part(34, "source-right-sleeve", sleeves)

    # The neck is a child of the head pivot: both rotate together from the base
    # of the neck while the shoulder and collar remain on the torso.
    head = add_joint_group(upper_body, "head", (344, 205), "neck-head")
    head.set("data-layer", "60")
    neck = add(head, "g", id="neck", data_rig_part="neck")
    neck.set("data-layer", "50")
    part(36, "source-neck", neck)
    part(37, "source-neck-shade", neck)
    for index in range(38, 51):
        part(index, f"source-head-{index}", head)

    guides = add(rig, "g", id="rig-guides", data_role="joint-guides")
    bones = [
        ((316, 238), (302, 269)), ((302, 269), (300, 304)),
        ((365, 238), (374, 269)), ((374, 269), (382, 302)),
        ((316, 286), (311, 350)), ((311, 350), (307, 391)),
        ((349, 286), (349, 350)), ((349, 350), (359, 390)),
        ((344, 205), (344, 238)), ((340, 286), (340, 217)),
    ]
    for index, (start, end) in enumerate(bones):
        add(guides, "line", id=f"bone-guide-{index+1}", x1=start[0], y1=start[1], x2=end[0], y2=end[1],
            stroke=ESPRESSO, stroke_width=2, stroke_dasharray="4 4", opacity=.55)
    for x, y, label in [
        (316, 238, "L shoulder"), (302, 269, "L elbow"), (300, 304, "L wrist"),
        (365, 238, "R shoulder"), (374, 269, "R elbow"), (382, 302, "R wrist"),
        (316, 286, "L hip"), (311, 350, "L knee"), (307, 391, "L ankle"),
        (349, 286, "R hip"), (349, 350, "R knee"), (359, 390, "R ankle"),
        (344, 205, "neck"), (340, 286, "torso"),
    ]:
        add(guides, "circle", cx=x, cy=y, r=4, **{"class": "joint-guide"}, data_joint=label)

    return root


def write_viewer():
    html = r'''<!doctype html>
<html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Raka · Character Rig</title>
<style>
:root{--plum:#3d2645;--ink:#2a1f2d;--sage:#7fb89a;--peach:#ffc4a3;--paper:#fff6ec;--muted:#756a78;--line:#e7ddd3;--white:#fff}
*{box-sizing:border-box}body{margin:0;background:var(--paper);color:var(--ink);font:15px/1.45 Inter,Arial,sans-serif}
.shell{width:min(1180px,calc(100% - 36px));margin:36px auto 64px}.top{display:flex;justify-content:space-between;gap:24px;align-items:end;margin-bottom:24px}.eyebrow{font-size:11px;font-weight:800;letter-spacing:.14em;text-transform:uppercase;color:var(--plum)}h1{font-size:clamp(34px,5vw,58px);line-height:1;margin:9px 0;letter-spacing:-.045em}.lead{max-width:690px;color:var(--muted);margin:0}.download{border-radius:999px;padding:12px 16px;background:var(--plum);color:white;text-decoration:none;font-weight:700;white-space:nowrap}
.layout{display:grid;grid-template-columns:minmax(360px,1fr) minmax(410px,1fr);gap:18px;align-items:stretch}.stage,.panel{border:1px solid var(--line);border-radius:20px;background:var(--white);box-shadow:0 16px 40px #2a1f2d12}.stage{display:grid;place-items:center;min-height:680px;background:radial-gradient(circle at 50% 42%,#f4e7da,white 66%);overflow:hidden}.stage object{width:100%;height:680px;border:0}.panel{padding:22px}.panel-head{display:flex;justify-content:space-between;align-items:center;gap:10px}.panel h2{font-size:19px;margin:0}.status{font-size:11px;background:#e6f1e9;color:#365d47;border-radius:999px;padding:5px 9px;font-weight:700}.presets{display:flex;flex-wrap:wrap;gap:7px;margin:17px 0 20px}.presets button,.tools button{border:1px solid var(--line);border-radius:999px;background:white;color:var(--ink);padding:9px 12px;font:inherit;font-size:12px;font-weight:700;cursor:pointer}.presets button:hover,.tools button:hover{border-color:var(--plum);color:var(--plum)}.control-grid{display:grid;grid-template-columns:1fr 1fr;gap:13px 18px}.control{display:grid;grid-template-columns:1fr auto;gap:3px 8px;align-items:center}.control label{font-size:12px;font-weight:700}.control output{font-variant-numeric:tabular-nums;color:var(--muted);font-size:11px}.control input{grid-column:1/-1;width:100%;accent-color:var(--plum)}.tools{display:flex;gap:8px;flex-wrap:wrap;margin-top:20px;padding-top:16px;border-top:1px solid var(--line)}.help{margin:17px 0 0;color:var(--muted);font-size:12px}.credit{margin-top:21px;padding-top:14px;border-top:1px solid var(--line);font-size:11px;color:var(--muted)}.credit a{color:var(--plum);font-weight:700}.joint-toggle{display:flex;align-items:center;gap:7px;color:var(--muted);font-size:12px}.joint-toggle input{accent-color:var(--plum)}.downloads{display:grid;gap:8px;min-width:146px}.download.secondary{background:#fff;color:var(--plum);border:1px solid var(--line);text-align:center}
.side-note{margin:14px 0 0;padding:9px 11px;border-radius:10px;background:#f7f0e8;color:#64576a;font-size:11px}.notice{min-height:18px;margin:10px 0 0;color:var(--muted);font-size:12px}.notice[data-state="error"]{color:#a53e26;font-weight:700}.notice[data-state="success"]{color:#365d47;font-weight:700}
@media(max-width:860px){.top{align-items:start;flex-direction:column}.layout{grid-template-columns:1fr}.stage{min-height:570px}.stage object{height:570px}}
@media(max-width:520px){.shell{width:calc(100% - 22px);margin:18px auto 38px}.panel{padding:16px}.stage{min-height:500px}.stage object{height:500px}.control-grid{grid-template-columns:1fr}.download{white-space:normal}}
</style></head>
<body><main class="shell"><header class="top"><div><div class="eyebrow">Editable vector character · Co-workers / Rafiki</div><h1>Raka, siap di-rig.</h1><p class="lead">Pose dasar berdiri. Torso, kepala, bahu, siku, pergelangan, pinggul, lutut, dan mata kaki punya pivot serta batas gerak. Sambungan dilapisi bentuk pengisi agar tetap menyatu saat ditekuk.</p></div><div class="downloads"><a class="download" href="Raka_Rig.svg" download>Unduh SVG rig</a><a class="download secondary" href="Raka_Rig.json" download>Unduh manifest JSON</a></div></header>
<section class="layout"><div class="stage"><object id="rakaObject" type="image/svg+xml" data="Raka_Rig.svg" aria-label="Raka standing rig preview"></object></div><div class="panel"><div class="panel-head"><h2>Atur pose</h2><span class="status" id="status">Neutral</span></div>
<p class="side-note" id="sideConvention">Kiri layar dan kanan layar mengacu pada sudut pandang penonton dalam pose netral.</p>
<div class="presets"><button data-preset="neutral">Berdiri</button><button data-preset="wave">Melambai</button><button data-preset="arms-up">Angkat kedua tangan</button><button data-preset="open">Gesture terbuka</button></div>
<div class="control-grid" id="controls"></div>
<div class="tools"><button id="waveMotion">▶ Gerakan lambaian</button><button id="guides">Tampilkan sambungan</button><button id="reset">Reset pose</button><button id="downloadPose">Unduh pose SVG</button><button id="exportPoseJson">Unduh pose JSON</button><button id="loadPoseButton">Impor pose JSON</button><input id="poseFile" type="file" accept="application/json,.json" hidden></div>
<p class="notice" id="poseNotice" aria-live="polite"></p>
<p class="help">Viewer membaca pivot dan batas rotasi dari manifest JSON. SVG menyimpan nama group, parent, urutan layer, koordinat pivot, dan batas geraknya agar rig mudah dipindahkan ke Adobe Animate, After Effects, HyperFrames, Anime.js, atau editor SVG lain.</p>
<div class="credit">Karakter dasar dari <a href="https://storyset.com/illustration/co-workers/rafiki" target="_blank" rel="noreferrer">Co-workers · Rafiki</a>. Kursi dan gelas sumber tidak disertakan; sambungan lengan dan kaki yang tidak tersedia digambar dengan warna dan bentuk yang mengikuti ilustrasi sumber.</div>
</div></section></main>
<script>
const rigPackage=window.rigPackage=__RIG_PACKAGE__;
const joints=rigPackage.joints;
const initial=Object.fromEntries(Object.keys(joints).map(key=>[key,0]));
const presets=Object.fromEntries(Object.entries(rigPackage.presets).map(([key,pose])=>[key,{...initial,...pose}]));
const object=document.getElementById('rakaObject');const controls=document.getElementById('controls');const status=document.getElementById('status');
let svgDoc=null, angles={...initial}, waveFrame=0, waveOn=false, poseFrame=0;
function transformFor(key,value){const [x,y]=joints[key].pivot;const node=svgDoc?.getElementById(joints[key].id);if(node)node.setAttribute('transform',`rotate(${value} ${x} ${y})`)}
function setAngle(key,value,updateControl=true){const joint=joints[key];const bounded=Math.max(joint.rotation[0],Math.min(joint.rotation[1],Number(value)));angles[key]=bounded;transformFor(key,bounded);const output=document.querySelector(`[data-output="${key}"]`);if(output)output.value=`${Math.round(bounded)}°`;const input=document.querySelector(`[data-input="${key}"]`);if(updateControl&&input)input.value=bounded}
function stopPoseTween(){if(poseFrame)cancelAnimationFrame(poseFrame);poseFrame=0}
function applyPose(pose,label,onComplete=null){stopWave();stopPoseTween();const from={...angles},to={...initial,...pose},start=performance.now(),duration=420;status.textContent=label;const tick=now=>{const t=Math.min(1,(now-start)/duration),ease=1-Math.pow(1-t,3);Object.keys(joints).forEach(key=>setAngle(key,from[key]+(to[key]-from[key])*ease));if(t<1)poseFrame=requestAnimationFrame(tick);else{poseFrame=0;if(onComplete)onComplete()}};poseFrame=requestAnimationFrame(tick)}
Object.entries(joints).forEach(([key,joint])=>{const wrap=document.createElement('div');wrap.className='control';wrap.innerHTML=`<label for="ctrl-${key}">${joint.label}</label><output data-output="${key}">0°</output><input id="ctrl-${key}" data-input="${key}" type="range" min="${joint.rotation[0]}" max="${joint.rotation[1]}" value="0" step="1" aria-label="${joint.label}">`;controls.appendChild(wrap);wrap.querySelector('input').addEventListener('input',e=>{stopWave();stopPoseTween();setAngle(key,e.target.value,false);status.textContent='Custom'})});
object.addEventListener('load',()=>{svgDoc=object.contentDocument;Object.keys(joints).forEach(key=>transformFor(key,angles[key]));});
document.querySelectorAll('[data-preset]').forEach(button=>button.addEventListener('click',()=>applyPose(presets[button.dataset.preset],button.textContent)));
document.getElementById('reset').addEventListener('click',()=>applyPose(initial,'Neutral'));
document.getElementById('guides').addEventListener('click',()=>{const guide=svgDoc?.getElementById('rig-guides');if(!guide)return;const shown=guide.style.opacity==='1';guide.style.opacity=shown?'0':'1';document.getElementById('guides').textContent=shown?'Tampilkan sambungan':'Sembunyikan sambungan'});
function stopWave(){waveOn=false;if(waveFrame)cancelAnimationFrame(waveFrame);waveFrame=0;document.getElementById('waveMotion').textContent='▶ Gerakan lambaian'}
document.getElementById('waveMotion').addEventListener('click',()=>{if(waveOn){stopWave();return}applyPose(presets.wave,'Melambai',()=>{waveOn=true;document.getElementById('waveMotion').textContent='■ Hentikan lambaian';const start=performance.now();const tick=now=>{if(!waveOn)return;const swing=Math.sin((now-start)/240)*18;setAngle('rightWrist',swing);waveFrame=requestAnimationFrame(tick)};waveFrame=requestAnimationFrame(tick)})});
document.getElementById('downloadPose').addEventListener('click',()=>{if(!svgDoc)return;const xml=new XMLSerializer().serializeToString(svgDoc.documentElement);const blob=new Blob([xml],{type:'image/svg+xml;charset=utf-8'});const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download='Raka_Pose.svg';link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1000)});
function downloadBlob(blob,name){const link=document.createElement('a');link.href=URL.createObjectURL(blob);link.download=name;link.click();setTimeout(()=>URL.revokeObjectURL(link.href),1000)}
document.getElementById('exportPoseJson').addEventListener('click',()=>{const pose={character_id:rigPackage.character_id,rig_type:'raka-svg-pose',manifest_version:rigPackage.version,angles:{...angles}};downloadBlob(new Blob([JSON.stringify(pose,null,2)],{type:'application/json;charset=utf-8'}),'Raka_Pose.json')});
const poseFile=document.getElementById('poseFile'),poseNotice=document.getElementById('poseNotice');
document.getElementById('loadPoseButton').addEventListener('click',()=>poseFile.click());
poseFile.addEventListener('change',async event=>{const file=event.target.files?.[0];if(!file)return;try{const pose=JSON.parse(await file.text());if(pose.character_id!==rigPackage.character_id||pose.rig_type!=='raka-svg-pose'||pose.manifest_version!==rigPackage.version)throw new Error('Karakter atau versi manifest tidak cocok.');const keys=Object.keys(joints),provided=Object.keys(pose.angles||{});const missing=keys.filter(key=>!Object.prototype.hasOwnProperty.call(pose.angles||{},key)),extra=provided.filter(key=>!Object.prototype.hasOwnProperty.call(joints,key));if(missing.length||extra.length)throw new Error(`Daftar sendi tidak lengkap. Kurang: ${missing.join(', ')||'—'}; asing: ${extra.join(', ')||'—'}.`);for(const [key,value] of Object.entries(pose.angles)){const [min,max]=joints[key].rotation;if(!Number.isFinite(value)||value<min||value>max)throw new Error(`${joints[key].label} harus berada antara ${min}° dan ${max}°.`)}applyPose(pose.angles,'Pose dimuat',()=>{poseNotice.dataset.state='success';poseNotice.textContent=`Pose “${file.name}” berhasil dimuat.`})}catch(error){poseNotice.dataset.state='error';poseNotice.textContent=`Gagal memuat pose: ${error.message}`}finally{poseFile.value=''}});
</script></body></html>'''
    html = html.replace("__RIG_PACKAGE__", json.dumps(RIG_PACKAGE, ensure_ascii=False, separators=(",", ":")))
    (ROOT / "index.html").write_text(html, encoding="utf-8")


def main():
    svg = build_svg()
    tree = ET.ElementTree(svg)
    tree.write(OUT, encoding="utf-8", xml_declaration=True)
    MANIFEST.write_text(json.dumps(RIG_PACKAGE, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    cairosvg.svg2png(url=str(OUT), write_to=str(PREVIEW), output_width=920, output_height=953,
                     background_color="#FFF6EC")
    write_viewer()
    print("Generated", OUT.name, MANIFEST.name, PREVIEW.name, "and index.html")


if __name__ == "__main__":
    main()
