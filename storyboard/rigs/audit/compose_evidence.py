#!/usr/bin/env python3
"""Compose actual Playwright screenshots into five-angle joint review sheets."""

import json
import math
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parent
EVIDENCE = ROOT / "evidence"
FRAMES = EVIDENCE / "joint-frames"
CELL_W, CELL_H = 205, 190
GAP = 12
SAMPLES = ("1", "2", "3", "4", "5")


def font(size):
    try:
        return ImageFont.truetype("DejaVuSans.ttf", size)
    except OSError:
        return ImageFont.load_default()


def joint_angles(bounds):
    low, high = bounds
    js_round = lambda value: math.floor(value + 0.5)
    if low < 0 < high:
        return (low, js_round(low / 2), 0, js_round(high / 2), high)
    span = high - low
    return (low, js_round(low + span * 0.25), js_round(low + span * 0.5), js_round(low + span * 0.75), high)


def compose_group(title, joints, manifest, filename):
    columns = len(SAMPLES)
    rows = len(joints)
    board = Image.new(
        "RGB",
        (GAP + columns * (CELL_W + GAP), 68 + rows * (CELL_H + GAP)),
        "#fff6ec",
    )
    draw = ImageDraw.Draw(board)
    draw.text((GAP, 12), f"{title} · exact degrees shown per sample", fill="#2a1f2d", font=font(16))
    for column, label in enumerate(SAMPLES):
        x = GAP + column * (CELL_W + GAP)
        draw.text((x + 4, 42), label, fill="#3d2645", font=font(12))

    for row, key in enumerate(joints):
        bounds = manifest["joints"][key]["rotation"]
        for column, angle in enumerate(joint_angles(bounds)):
            x = GAP + column * (CELL_W + GAP)
            y = 68 + row * (CELL_H + GAP)
            draw.rounded_rectangle((x, y, x + CELL_W, y + CELL_H), radius=8, fill="#ffffff")
            if column == 0:
                draw.text((x + 5, y + 5), key, fill="#2a1f2d", font=font(11))
            draw.text((x + CELL_W - 44, y + 5), f"{angle}°", fill="#3d2645", font=font(10))
            screenshot = FRAMES / f"{key}-{angle}.png"
            if not screenshot.exists():
                raise FileNotFoundError(f"Missing actual browser frame: {screenshot}")
            with Image.open(screenshot) as source:
                image = source.convert("RGB")
                image.thumbnail((CELL_W - 12, CELL_H - 26), Image.Resampling.LANCZOS)
                paste_x = x + (CELL_W - image.width) // 2
                paste_y = y + 21 + (CELL_H - 22 - image.height) // 2
                board.paste(image, (paste_x, paste_y))
    board.save(EVIDENCE / filename, optimize=True)


if __name__ == "__main__":
    manifest = json.loads((ROOT.parent / "Raka_Rig.json").read_text(encoding="utf-8"))
    compose_group("Head and torso · limits and intermediate angles", ["head", "torso"], manifest, "joint-sweep-head-torso.png")
    compose_group(
        "Arms · limits and intermediate angles",
        ["leftShoulder", "leftElbow", "leftWrist", "rightShoulder", "rightElbow", "rightWrist"],
        manifest,
        "joint-sweep-arms.png",
    )
    compose_group(
        "Legs · limits and intermediate angles",
        ["leftHip", "leftKnee", "leftAnkle", "rightHip", "rightKnee", "rightAnkle"],
        manifest,
        "joint-sweep-legs.png",
    )
