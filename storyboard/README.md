# Therapy Speak Gone Wrong storyboard

The selected **C production board** contains seven scenes and twelve standalone, editable 16:9 animation frames. Scenes 3–4 now form one five-beat story over **80 seconds**: Raka mislabels his sadness, clashes with Dina, labels her narcissistic from one defensive reply, spreads the label at work, then warns the group and removes her before he begins to doubt. The same Storyset Co-workers illustration is reused in every beat so Raka and Dina remain visually consistent.

## Scene 3–4 timing and VO

| Beat | Time | Duration | Story action |
| --- | --- | ---: | --- |
| 03A | 01:05–01:20 | 15 sec | Raka misses a target, feels sad for several days, and calls himself depressed. |
| 03B | 01:20–01:38 | 18 sec | Raka hears Dina’s meeting feedback as an attack; both respond defensively. |
| 03C | 01:38–01:52 | 14 sec | Raka labels Dina “narsistik” based on one reply. |
| 04A | 01:52–02:06 | 14 sec | Raka shares the label; coworkers begin avoiding Dina. |
| 04B | 02:06–02:25 | 19 sec | Raka posts a warning, removes Dina, then starts to doubt his conclusion. |

The Indonesian-English voice-over appears beat by beat with each duration in the Word and web boards. Object transitions link the self-label cloud to the meeting bubbles, the “NARSISTIK” stamp, group notifications, and the final warning post. On-screen text makes each beat understandable without VO. Final copy: **“SATU MOMEN BUKAN BUKTI · Pahami konteks sebelum memberi label.”**

## Files

- `index.html` — the production board with frame previews, timings, VO, motion notes, asset credits, and SVG/PNG downloads.
- `scenes/scene-*.svg` — editable scene frames. Scene 2 has three frames; Scene 3 has three linked beats; Scene 4 has two linked beats. Frames contain the background, illustration, props, overlays, and on-screen text only.
- `scenes/scene-02-komponen.svg` — six additional editable comment groups for the Scene 2 scroll animation.
- `assets/Co-workers_Rafiki.svg` — the downloaded Storyset source illustration used across all five Scene 3–4 beats.
- `previews/scene-*.png` — rendered frame previews used by the website and Word board.
- `previews/contact-sheet.png` — the 7-scene/12-frame visual audit sheet.
- `Storyboard_Therapy_Speak_Gone_Wrong.docx` — landscape production board with the contact sheet, per-scene frames, beat-level VO, and asset links.
- `Storyboard_Editable_Source.zip` — the web board, Word file, scene/source SVGs, and previews.
- `build_storyboard.py` — rebuilds the SVG frames, previews, Word document, website, and source bundle.

## Visual system and asset audit

Opsi B is **Warm Peach & Mustard**: deep plum `#3D2645`, sage `#7FB89A`, terracotta `#E4572E`, mustard `#F2B134`, peach `#FFC4A3`, cream `#FFF6EC`, and espresso `#2A1F2D`. Scenes 3–4 keep cream and plum dominant, with terracotta reserved for mistaken labels and harmful actions. Sage carries the final corrective message. Each frame uses no more than three dominant colors; muted colors already present in the licensed illustration remain as small supporting details.

The Storyset source contains 158 drawable elements. In beats 03A–04A, 36 elements (22.8%) change or are omitted: 28 blue elements are recolored plum, seven placeholder-chat elements are omitted, and Raka’s mouth is adjusted to read as concerned. The remaining **122 elements (77.2%) stay unchanged**. In final frame 04B, Dina stands alone in the office and a close-up from the same Raka SVG group appears in a separate doubt inset. The conservative audit counts Raka’s 54-element body as omitted from the office, eight recolored elements, and the changed mouth: 63/158 (39.9%), leaving at least **60.1% unchanged**. Dialog, cloud, target card, stamp, notifications, warning post, and closing text are separate editable layers.

The complete frames were rendered and reviewed at full size and in the contact sheet. The audit checked character continuity, readability of all five beats without narration, the transition-object chain, palette, and the <40% artwork-edit limit. No character art was generated or drawn from scratch.

## Editing

Open any SVG in Inkscape, Illustrator, Figma, or another SVG editor. Character artwork remains a grouped SVG; IDs identify the animation layers, including `emotion-cloud`, `raka-quote`, `dina-quote`, `narcissistic-stamp`, `rumor-notifications`, `warning-post`, `remove-dina-action`, `raka-doubt`, and `final-onscreen-text`.

Edit `build_storyboard.py` to change the scene copy or composition, then run:

```sh
python storyboard/build_storyboard.py
```

The builder uses Python packages `CairoSVG`, `Pillow`, and `python-docx`.

## Storyset credit and source

The [Co-workers · Rafiki illustration](https://storyset.com/illustration/co-workers/rafiki) is listed as free for personal and commercial use with attribution. Include Storyset credit when publishing and check the current [Storyset terms](https://storyset.com/terms). The editable original is included in `assets/Co-workers_Rafiki.svg`.
