# Deep visual and functional rig audit

**Local browser result:** 328 checks passed, 0 failed.

**GitHub Actions:** 328 checks passed, 0 failed on `ubuntu-latest` in [run 46004678457](https://github.com/nkurnia1031/video_hani/actions/runs/46004678457). The workflow uploaded a 14.2 MB evidence artifact with a 14-day retention period.
**Tested:** 2026-11-27 with Playwright 1.65.1, Chromium 159.0.8154.6, Node v22.23.3.

The audit opened the editable viewer in Chromium and manipulated its actual controls. It tested all 14 sliders at five angles each, all four presets, one combined pose, pose interpolation, 48 samples across the wrist wave, stop/reset, SVG and JSON export/import, invalid import rejection, and 1440×1000, 820×1100, and 390×844 viewports. The tested angle values are printed in the contact sheets. Screenshot evidence comes from that browser run.

## A. Critical Bugs

No critical bugs remain in the tested build. The browser reported no JavaScript exceptions, console errors, or failed requests.

## B. Major Visual Issues

- **Head and neck — fixed.** The neck was outside the head joint group, so it stayed behind when the head rotated. It is now a child of the head group. The baseline head-negative screenshot shows the former pose; the current head and torso sweep shows both rotation limits and intermediate samples.
- **Knees — fixed.** The former −95° to +95° limits allowed the knee to bend backwards and caused the trouser and calf shapes to overlap heavily. The left knee now flexes from 0° to 70° and the right from −70° to 0°, matching their screen-side bend direction. The five-angle leg sheet shows the corrected range.
- **Sampled connections — no visible gap found.** I inspected the current joint sweeps and the combined stress pose for exposed gaps, clipping, disconnected limbs, and obvious joint intersections. None were visible within the tested limits. These are discrete angle samples plus selected combined poses, not a pixel-level proof for every possible combination.

Evidence: [baseline head pose](evidence/baseline-head-negative.png), [baseline knee extremes](evidence/baseline-joint-extremes-legs.png), [head and torso sweep](evidence/joint-sweep-head-torso.png), [arm sweep](evidence/joint-sweep-arms.png), [leg sweep](evidence/joint-sweep-legs.png), [combined stress pose](evidence/pose-combined-stress.png).

## C. Minor Visual Issues

No remaining issue was visible in the sampled poses. This is a cutout SVG rig: parts rotate as authored shapes; it does not deform the source paths like a mesh rig.

## D. Functional Issues

- **Pose JSON import/export — fixed.** The viewer now exports the active joint angles and imports matching poses. Browser tests round-tripped the pose and confirmed that wrong-version, incomplete, and out-of-range files are rejected without changing the current pose.
- **Sliders and presets — passed.** All 14 controls matched their manifest labels and bounds, changed only the intended joint, displayed the active angle, and kept the pivot anchored. All four presets reached their targets without inheriting a prior pose.
- **Tween, wave, and reset — passed.** Preset interpolation was observed at an intermediate frame; the repeated wrist motion completed multiple cycles within its limits and stopped cleanly; reset returned all controls to zero.
- **Responsive viewer — passed.** No horizontal overflow or hidden controls was found at the three tested viewport sizes.

Evidence: [neutral](evidence/pose-neutral.png), [wave preset](evidence/pose-wave.png), [arms-up preset](evidence/pose-arms-up.png), [open preset](evidence/pose-open.png), [mid-transition](evidence/transition-mid-wave.png), [JSON-restored pose](evidence/pose-json-restored.png), [desktop](evidence/viewport-1440.png), [tablet](evidence/viewport-820.png), [mobile](evidence/viewport-390.png).

## E. Structural Issues

- **Neck parent — fixed.** `neck` now belongs to `head`; `head` belongs to `upper-body`.
- **Pelvis root — fixed.** `upper-body` and `rig-legs` now share `pelvis` as their parent. The torso rotation pivots at the pelvis root.
- **Manifest and SVG hierarchy — passed.** Every declared part, parent, layer, pivot, and limit matches the rendered SVG. There are no duplicate SVG IDs. Left and right explicitly mean the viewer's left and right in the neutral pose.

The structure and pose manifest are maintained by [build_raka_rig.py](../build_raka_rig.py), [Raka_Rig.json](../Raka_Rig.json), and [Raka_Rig.svg](../Raka_Rig.svg). The viewer and test instructions are in [index.html](../index.html) and [README.md](README.md).

## Remaining scope limit

Five samples per joint and representative combined/preset poses do not enumerate every mathematical combination of 14 joint angles. The current pass applies to the configured limits and exercised browser states; changing a limit, pivot, parent, or source shape should rerun the GitHub Actions audit and receive another visual review.
