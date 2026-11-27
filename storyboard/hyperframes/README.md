# Scene 2–3 · HyperFrames source

This folder contains the editable 35-second Scene 2 project for Therapy Speak Gone Wrong. The nested `scene-03/` folder is a separate 50-second HyperFrames project, loaded by the same web page.

- index.html is the 1920×1080 root timeline. It sequences three 16:9 compositions.
- compositions/scene-02a.html animates the existing room SVG, props, person, phone, and ten comments.
- compositions/scene-02b.html animates the existing layered hashtag cloud.
- compositions/scene-02c.html reveals the two existing closing statements.
- ../motion.html embeds the root composition with the HyperFrames web player and provides scene jump buttons.
- scene-03/index.html is the 50-second Scene 3 root timeline. It sequences three separate 15-second cases and the 5-second closing frame.
- scene-03/compositions/scene-03a.html through scene-03d.html animate the four existing Scene 3 SVG frames. Their source SVGs are copied unchanged into scene-03/assets/scenes/ for local composition loading.
- Scene 3 captions reproduce the PDF voice-over verbatim and reveal word by word with a mustard active-word tween. Timing is an editable simulation; no voice recording is embedded.
- Both players share ../motion.html and provide separate seek, play, and section controls. Scene 3 has its own root project so each sequence can play without overlapping timelines or audio.

Edit the SVG layer markup and paused GSAP timelines inside each composition's template. The character and room illustration comes from the existing, credited Storyset Feed · Rafiki asset; no character art was generated. Scene 3 uses the existing Texting, Student Stress, and Couple Stress Storyset SVGs without redrawing their characters. Added comments were brought over from the editable Scene 2 source SVG. Caveat is bundled in assets/ with its SIL Open Font License and loaded by the Scene 2 root composition.

For a local HyperFrames Studio preview, use the installed HyperFrames CLI. The project pins CLI 0.9.99 for reproducibility. The static web player loads GSAP and @hyperframes/player from jsDelivr.
