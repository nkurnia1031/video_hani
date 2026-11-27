# Raka rig audit

The audit opens the actual pose viewer in Chromium and interacts with every slider. It checks SVG/manifest structure, pivot stability, joint limits, isolated controls, presets, combined rotations, tween interpolation, the repeating wrist wave, reset behavior, SVG and JSON serialization, invalid-pose rejection, responsive viewports, and browser errors.

Run locally:

```sh
npm ci
npx playwright install --with-deps chromium
python -m pip install Pillow
npm run audit
```

The command writes a JSON result, pose screenshots, three contact sheets covering all joints at their tested angles, and viewport screenshots under `evidence/`. GitHub Actions runs the same command on pushes and pull requests that touch `storyboard/rigs/`, and stores the generated files as a 14-day workflow artifact.

Set `PLAYWRIGHT_BROWSERS_PATH` if Chromium is installed in a non-default local cache.
