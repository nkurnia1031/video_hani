# Scene 2 · HyperFrames source

This is the editable 35-second Scene 2 animation project for Therapy Speak Gone Wrong.

- index.html is the 1920×1080 root timeline. It sequences three 16:9 compositions.
- compositions/scene-02a.html animates the existing room SVG, props, person, phone, and ten comments.
- compositions/scene-02b.html animates the existing layered hashtag cloud.
- compositions/scene-02c.html reveals the two existing closing statements.
- ../motion.html embeds the root composition with the HyperFrames web player and provides scene jump buttons.

Edit the SVG layer markup and paused GSAP timelines inside each composition's template. The character and room illustration comes from the existing, credited Storyset Feed · Rafiki asset; no character art was generated. Added comments were brought over from the editable Scene 2 source SVG. Caveat is bundled in assets/ with its SIL Open Font License and loaded by the root composition.

For a local HyperFrames Studio preview, use the installed HyperFrames CLI. The project pins CLI 0.9.99 for reproducibility. The static web player loads GSAP and @hyperframes/player from jsDelivr.
