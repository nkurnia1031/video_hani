# Therapy Speak Gone Wrong storyboard

The selected **C production board** contains seven 16:9 scene frames as standalone, editable SVGs. The website presents one production row per scene; the Word document keeps the same storyboard in a printable format.

## Files

- `index.html` — production board website with each scene's timing, visual description, on-screen copy, motion notes, voice-over cue, and asset links.
- `scenes/scene-*.svg` — editable 1920 × 1080 scene compositions. Storyset artwork is embedded as SVG groups; backgrounds, type, chips, cards, and icons are editable vector elements.
- `previews/scene-*.png` — raster previews used in the website and Word document.
- `previews/contact-sheet.png` — seven-scene visual audit sheet.
- `Storyboard_Therapy_Speak_Gone_Wrong.docx` — landscape production board with contact sheet, per-scene layouts, source links, and visual audit notes.
- `Storyboard_Editable_Source.zip` — self-contained website, SVGs, additional illustration, PNG previews, and Word file.
- `build_storyboard.py` — rebuilds the scene SVGs, previews, contact sheet, Word document, website, and source bundle.

## Editing

Open any SVG in Inkscape, Illustrator, Figma, or another SVG editor. The scene background, headings, labels, cards, connectors, and custom icons are vector objects. Embedded Storyset illustrations remain vector groups and can be moved, resized, or edited separately.

Edit `build_storyboard.py` to change the scene copy or visual composition, then run:

```sh
python storyboard/build_storyboard.py
```

The build script uses Python packages `CairoSVG`, `Pillow`, and `python-docx`.

## Visual system and audit

The nine illustrations used in the scenes share the flat Rafiki vector style, blue `#407BFF`, ink `#263238`, warm neutral backgrounds, and soft shadows. Custom scene backgrounds use that same palette, while a scene-specific accent signals the emotional shift. The contact sheet is the quick pacing and color check; full-size scene previews are in `previews/`.

No images were generated. The presenter is a footage placeholder, and the scene backgrounds, label chips, article cards, domino cards, expression icons, and supporting marks are handmade SVG vectors.

## Storyset attribution and download links

Storyset lists the illustrations below as free with attribution and allows modifications and derivative works under its license summary. Add Storyset/Freepik attribution when publishing the final video, and check the current [Storyset terms](https://storyset.com/terms).

| Scene | Illustration | Role | Download page |
| --- | --- | --- | --- |
| 2 | Feed | Social media context | [Feed · Rafiki](https://storyset.com/illustration/feed/rafiki) |
| 3 | Texting | Unanswered message | [Texting · Rafiki](https://storyset.com/illustration/texting/rafiki) |
| 3 | Feeling Blue | Sadness after a setback | [Feeling Blue · Rafiki](https://storyset.com/illustration/feeling-blue/rafiki) |
| 3 | Couple Stress | Disagreement context | [Couple Stress · Rafiki](https://storyset.com/illustration/couple-stress/rafiki) |
| 4 | Problem Solving | Searching for context | [Problem Solving · Rafiki](https://storyset.com/illustration/problem-solving/rafiki) |
| 5 | Online Document | Digital ethics page | [Online Document · Rafiki](https://storyset.com/illustration/online-document/rafiki) |
| 5 | Research Paper | Scientific basis | [Research Paper · Rafiki](https://storyset.com/illustration/research-paper/rafiki) |
| 6 | To do list | Checklist illustration | [To do list · Rafiki](https://storyset.com/illustration/to-do-list/rafiki) |
| 7 | Social Interaction | Group with varied expressions | [Social Interaction · Rafiki](https://storyset.com/illustration/social-interaction/rafiki) |
