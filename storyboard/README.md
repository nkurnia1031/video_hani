# Therapy Speak Gone Wrong storyboard

The selected **C production board** contains seven scenes and fifteen 16:9 keyframes as standalone, editable SVGs. Scene 02A uses the Opsi B cream room, espresso phone, plum UI details, and restrained warm accents. Scene 02B stacks hashtag cards in terracotta, mustard, and peach; 02C centers the two script statements in sage and terracotta. `scene-02-komponen.svg` contains six additional editable comment groups (`comment-05` through `comment-10`) for the scroll animation. Scene 03A–03C present each case separately with its assigned background tint, terracotta label, and sage correction. Scene 03D combines the three examples behind centered on-screen text. Scene 04A–04D link self-misunderstanding, conflict, stigma, and a final turn toward context and evidence; 04D ends with on-screen text. Animation frames contain only the on-screen composition, without board titles, timestamps, footers, or production notes. The website presents one production row per scene; the Word document keeps the same storyboard in a printable format.

## Files

- `index.html` — production board website with each scene's timing, visual description, on-screen copy, motion notes, voice-over cue, and asset links.
- `scenes/scene-*.svg` — editable 1920 × 1080 scene compositions. Scene 2 has three frames; Scene 3 has three individual cases and a combined text frame; Scene 4 has four linked consequence frames ending with on-screen text. Storyset artwork is embedded as SVG groups; backgrounds, type, chips, cards, and icons are editable vector elements.
- `scenes/scene-02-komponen.svg` — six separate editable comment groups for the Scene 2 phone scroll.
- `previews/scene-*.png` — raster previews used in the website and Word document.
- `previews/contact-sheet.png` — seven-scene/fifteen-frame visual audit sheet.
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

Opsi B is **Warm Peach & Mustard**: deep plum `#3D2645`, sage `#7FB89A`, terracotta `#E4572E`, mustard `#F2B134`, peach `#FFC4A3`, cream `#FFF6EC`, and espresso `#2A1F2D`. Scene 2 keeps cream, espresso, and plum dominant; its terminology and hashtags use terracotta, mustard, and peach. Scene 3 uses peach `#FFE3D3`, sage `#DCEFE4`, and mustard `#FCE9C2` case tints. Scene 4 continues those tints in order—peach, mustard, sage, then cream—and uses a plum closing panel. Terracotta marks careless labels; sage carries context and correction. Scene 4 SVG metadata records the source-art changes; each adaptation keeps at least 60% of drawable elements unchanged. The contact sheet and full-size previews are the visual style audit.

No images were generated. The presenter is a footage placeholder. Scene backgrounds, label chips, article cards, connectors, and simple non-character icons are editable SVG vectors; characters and illustrated situations come from credited Storyset assets.

## Storyset attribution and download links

Storyset lists the illustrations below as free with attribution and allows modifications and derivative works under its license summary. Add Storyset/Freepik attribution when publishing the final video, and check the current [Storyset terms](https://storyset.com/terms).

| Scene | Illustration | Role | Download page |
| --- | --- | --- | --- |
| 2 | Feed | Social media context | [Feed · Rafiki](https://storyset.com/illustration/feed/rafiki) |
| 3 | Texting | Unanswered message | [Texting · Rafiki](https://storyset.com/illustration/texting/rafiki) |
| 3 | Feeling Blue | Sadness after a setback | [Feeling Blue · Rafiki](https://storyset.com/illustration/feeling-blue/rafiki) |
| 3 | Couple Stress | Disagreement context | [Couple Stress · Rafiki](https://storyset.com/illustration/couple-stress/rafiki) |
| 4 | Feeling Blue | Ordinary emotions misread as a personal problem | [Feeling Blue · Rafiki](https://storyset.com/illustration/feeling-blue/rafiki) |
| 4 | Couple Stress | Disagreement labeled as a personality flaw | [Couple Stress · Rafiki](https://storyset.com/illustration/couple-stress/rafiki) |
| 4 | Bullying | Social exclusion as a visual metaphor for stigma | [Bullying · Rafiki](https://storyset.com/illustration/bullying/rafiki) |
| 4 | Research Paper | Evidence and context resolve the chain | [Research Paper · Rafiki](https://storyset.com/illustration/research-paper/rafiki) |
| 5 | Online Document | Digital ethics page | [Online Document · Rafiki](https://storyset.com/illustration/online-document/rafiki) |
| 5 | Research Paper | Scientific basis | [Research Paper · Rafiki](https://storyset.com/illustration/research-paper/rafiki) |
| 6 | To do list | Checklist illustration | [To do list · Rafiki](https://storyset.com/illustration/to-do-list/rafiki) |
| 7 | Social Interaction | Group with varied expressions | [Social Interaction · Rafiki](https://storyset.com/illustration/social-interaction/rafiki) |
