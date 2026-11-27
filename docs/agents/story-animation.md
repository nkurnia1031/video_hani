# Storyboard Animation Guide

Apply this guide when planning, editing, or reviewing animated scenes for “Therapy Speak Gone Wrong.” It complements the HyperFrames instructions in `storyboard/hyperframes/AGENTS.md` and the illustration constraints in the root `AGENTS.md`.

## Build from the script

1. Read the source script and map each exact voice-over line to its scene, visual beat, on-screen text, and time range. Keep the spoken wording unchanged unless the user explicitly asks for a script revision.
2. Set scene and transition durations from the voice-over pacing. Keep animation cues attached to the phrase or action they explain; don't force the voice-over to fit arbitrary visual timing.
3. Keep each scene frame to the composition that will appear in the animation. Follow the root `AGENTS.md` for production-board material and final-frame on-screen text.

## Animate with intent

- Use editable SVG/HTML layers and seekable GSAP tweens. Give objects a clear entrance, action, and settle point. Stagger related elements and vary easing or direction to create rhythm; avoid making every object use the same slide-in.
- Motivate transitions with the scene's graphic language or a visible object. Let a transition connect the scenes, then leave enough time for the new composition and its text to read.
- Give small secondary motions a physical cause: a plant responds to a light breeze, a hand follows a gesture, or a prop reacts to an interaction. Use restrained amplitude, eased acceleration, and a short settle or follow-through.
- Rig only the part that should move. Separate a plant from its pot and pivot the stem/foliage at the pot rim; pivot a limb at its joint. Keep attachment points fixed. For a head or neck, use a believable joint and inspect the shoulder/neck connection through the whole motion. If the source illustration cannot support a natural rig, leave that part still.
- Review motion at its start, midpoint, peak, and return. A still frame can look correct while the motion between frames feels detached or abrupt.

## Choose music and sound effects

- Choose music from the scene's voice-over tone. Reflective or explanatory passages call for restrained, non-cheerful beds; keep music under the narration and use fades at scene boundaries.
- Give sound effects distinct jobs. Match each effect's attack and duration to the tween: a pop can land with a comment card, a whoosh can accompany a fast full-screen word build, and a soft hit can land with closing text. Avoid repeating the same swish on unrelated actions.
- Add a scroll effect only when its texture and timing fit the actual scroll tween. If the available sound feels late, long, or unlike the movement, omit it and let the music carry the action.
- Use a consistent cue and level for paired closing statements unless the script calls for a contrast. Place the cue at the visual landing, not merely somewhere during the text's hold.
- Match ambience to the scene's emotional shape. A rising wind or storm can support a dense, overwhelming word cloud when it fits the narration; it is an option, not a default for every cloud.
- Use audio with a clear reuse license and record its source and license in the project's credits. Check transitions for overlapping tails and check every cue in actual playback against the music and voice-over; file metadata or a timeline entry alone does not prove the sound is audible.

## Review before handoff

1. Preview the scene at the relevant voice-over beats. Capture snapshots around object entrances, transition boundaries, motion peaks, and final text; inspect the images rather than relying on a passing automated check.
2. Listen through the edited section. Confirm that each intended effect is audible, lands with the visual beat, and does not mask narration. Remove or retime any effect that fails this check.
3. Run the HyperFrames check required by `storyboard/hyperframes/AGENTS.md`; resolve errors and review warnings. Rebuild the editable source package when its contents changed.

## Scene 2 reference

- Scene 02A uses staggered object entrances and a pop for the highlighted comment. The feed scroll has no separate effect because the available shuffle did not match its tween.
- Scene 02B builds the handwritten hashtag cloud rapidly; recorded wind and thunder rise into the crowded hero-word moment.
- Scene 02C uses the same gentle landing sound and level for both closing statements.
- Shelf foliage moves independently from its pot with a pivot at the rim. The phone user's head remains still because the attempted head motion looked detached from the neck.

Treat these as approved examples of timing and sound roles, not a fixed template for every later scene.
