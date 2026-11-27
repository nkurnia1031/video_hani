const fs = require('node:fs');
const http = require('node:http');
const path = require('node:path');
const { chromium } = require('playwright');

const AUDIT_DIR = __dirname;
const REPO_ROOT = path.resolve(AUDIT_DIR, '../../..');
const EVIDENCE_DIR = path.join(AUDIT_DIR, 'evidence');
const FRAME_DIR = path.join(EVIDENCE_DIR, 'joint-frames');
const MANIFEST_PATH = path.join(REPO_ROOT, 'storyboard/rigs/Raka_Rig.json');
const REPORT_PATH = path.join(EVIDENCE_DIR, 'browser-audit.json');
const manifest = JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf8'));
const results = [];
const errors = [];

function check(name, condition, details = '') {
  results.push({ name, status: condition ? 'PASS' : 'FAIL', details: condition ? '' : details });
  if (!condition) errors.push(`${name}${details ? `: ${details}` : ''}`);
}

function angleCases([min, max]) {
  if (min < 0 && max > 0) return [...new Set([min, Math.round(min / 2), 0, Math.round(max / 2), max])];
  const span = max - min;
  return [...new Set([min, Math.round(min + span * 0.25), Math.round(min + span * 0.5), Math.round(min + span * 0.75), max])];
}

function fileType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return ({
    '.html': 'text/html; charset=utf-8', '.svg': 'image/svg+xml', '.json': 'application/json',
    '.png': 'image/png', '.css': 'text/css', '.js': 'text/javascript', '.woff2': 'font/woff2',
  })[ext] || 'application/octet-stream';
}

function startServer() {
  const server = http.createServer((request, response) => {
    const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    const candidate = path.resolve(REPO_ROOT, `.${pathname}`);
    if (!candidate.startsWith(REPO_ROOT + path.sep) && candidate !== REPO_ROOT) {
      response.writeHead(403).end('Forbidden');
      return;
    }
    fs.readFile(candidate, (error, body) => {
      if (error) {
        response.writeHead(error.code === 'ENOENT' ? 404 : 500).end('Not found');
        return;
      }
      response.writeHead(200, { 'content-type': fileType(candidate), 'cache-control': 'no-store' }).end(body);
    });
  });
  return new Promise(resolve => server.listen(0, '127.0.0.1', () => resolve(server)));
}

async function readAngles(page) {
  return page.evaluate(() => Object.fromEntries(
    [...document.querySelectorAll('.control input[type="range"]')]
      .map(input => [input.dataset.input, Number(input.value)]),
  ));
}

async function setSlider(page, key, angle) {
  await page.locator(`[data-input="${key}"]`).evaluate((input, value) => {
    input.value = String(value);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }, angle);
}

async function readJointState(page, key) {
  return page.evaluate(jointKey => {
    const joint = window.rigPackage.joints[jointKey];
    const object = document.querySelector('#rakaObject');
    const doc = object.contentDocument;
    const node = doc.getElementById(joint.id);
    const parent = node.parentElement;
    const pivot = doc.documentElement.createSVGPoint();
    pivot.x = joint.pivot[0];
    pivot.y = joint.pivot[1];
    const atJoint = pivot.matrixTransform(node.getScreenCTM());
    const parentAtJoint = pivot.matrixTransform(parent.getScreenCTM());
    const localRect = doc.documentElement.getBoundingClientRect();
    const hit = doc.elementFromPoint(atJoint.x - localRect.left, atJoint.y - localRect.top);
    const matrix = node.getScreenCTM();
    const determinant = matrix.a * matrix.d - matrix.b * matrix.c;
    return {
      transform: node.getAttribute('transform'),
      parentId: parent.id,
      pivotError: Math.hypot(atJoint.x - parentAtJoint.x, atJoint.y - parentAtJoint.y),
      pivotHit: hit?.id || hit?.tagName || null,
      determinant,
    };
  }, key);
}

async function downloadFile(page, buttonId, fileName) {
  const downloadPromise = page.waitForEvent('download');
  await page.locator(`#${buttonId}`).click();
  const download = await downloadPromise;
  const destination = path.join(EVIDENCE_DIR, fileName);
  await download.saveAs(destination);
  return destination;
}

async function main() {
  fs.rmSync(FRAME_DIR, { recursive: true, force: true });
  fs.mkdirSync(FRAME_DIR, { recursive: true });
  fs.mkdirSync(EVIDENCE_DIR, { recursive: true });
  const browser = await chromium.launch({ headless: true });
  const server = await startServer();
  const address = server.address();
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 1 });
  const browserErrors = [];
  page.on('pageerror', error => browserErrors.push(`pageerror: ${error.message}`));
  page.on('console', message => { if (message.type() === 'error') browserErrors.push(`console: ${message.text()}`); });
  page.on('requestfailed', request => browserErrors.push(`requestfailed: ${request.url()} ${request.failure()?.errorText || ''}`));

  try {
    const pageUrl = `http://127.0.0.1:${address.port}/storyboard/rigs/index.html`;
    await page.goto(pageUrl, { waitUntil: 'networkidle' });
    await page.waitForFunction(() => document.querySelector('#rakaObject')?.contentDocument?.getElementById('raka-rig'));
    await page.screenshot({ path: path.join(EVIDENCE_DIR, 'browser-desktop-neutral.png'), fullPage: true });

    const structure = await page.evaluate(() => {
      const object = document.querySelector('#rakaObject');
      const doc = object.contentDocument;
      const root = doc.documentElement;
      const group = id => doc.getElementById(id);
      const labels = [...document.querySelectorAll('.control label')].map(label => label.textContent.trim());
      const duplicateIds = [...doc.querySelectorAll('[id]')].map(node => node.id).filter((id, i, ids) => ids.indexOf(id) !== i);
      return {
        pageRigPackage: window.rigPackage,
        controlCount: document.querySelectorAll('.control input[type="range"]').length,
        labels,
        duplicateIds,
        viewBox: root.getAttribute('viewBox'),
        neckParent: group('neck')?.parentElement?.id,
        headParent: group('head')?.parentElement?.id,
        upperBodyParent: group('upper-body')?.parentElement?.id,
        legsParent: group('rig-legs')?.parentElement?.id,
        pelvisParent: group('pelvis')?.parentElement?.id,
        importControl: Boolean(document.querySelector('#poseFile')),
        jsonExportControl: Boolean(document.querySelector('#exportPoseJson')),
        sideConvention: document.querySelector('#sideConvention')?.textContent?.trim() || '',
        sourceCupIncluded: Boolean(doc.querySelector('[data-source-child="52"]')),
      };
    });

    check('browser loads viewer and SVG rig', Boolean(structure.pageRigPackage && structure.controlCount));
    check('manifest matches viewer data', JSON.stringify(structure.pageRigPackage) === JSON.stringify(manifest), 'embedded rig package and Raka_Rig.json differ');
    check('all 14 sliders are present', structure.controlCount === 14, `found ${structure.controlCount}`);
    check('SVG viewBox matches manifest', structure.viewBox === manifest.coordinate_space.viewBox, `${structure.viewBox} != ${manifest.coordinate_space.viewBox}`);
    check('no duplicate SVG IDs', structure.duplicateIds.length === 0, structure.duplicateIds.join(', '));
    check('head group follows neck group', structure.neckParent === 'head' && structure.headParent === 'upper-body', `neck parent=${structure.neckParent}, head parent=${structure.headParent}`);
    check('torso and legs share a pelvis parent', structure.upperBodyParent === 'pelvis' && structure.legsParent === 'pelvis' && structure.pelvisParent === 'raka-rig', `upper-body=${structure.upperBodyParent}, legs=${structure.legsParent}, pelvis=${structure.pelvisParent}`);
    check('pose JSON import/export controls exist', structure.importControl && structure.jsonExportControl, `import=${structure.importControl}, JSON export=${structure.jsonExportControl}`);
    check('left/right meaning is explicit in viewer', /kiri layar/i.test(structure.sideConvention) && /kanan layar/i.test(structure.sideConvention), structure.sideConvention || 'side note missing');
    check('chair and cup are excluded', !structure.sourceCupIncluded, 'source cup element found');

    const embeddedManifestCheck = await page.evaluate(packageData => {
      const doc = document.querySelector('#rakaObject').contentDocument;
      const ids = new Set([...doc.querySelectorAll('[id]')].map(node => node.id));
      const missingParts = packageData.parts.filter(part => !ids.has(part.id)).map(part => part.id);
      const badParents = packageData.parts.filter(part => doc.getElementById(part.id)?.parentElement?.id !== part.parent).map(part => `${part.id}→${doc.getElementById(part.id)?.parentElement?.id}`);
      const badLayers = packageData.parts.filter(part => doc.getElementById(part.id)?.getAttribute('data-layer') !== String(part.layer)).map(part => part.id);
      const badJoints = Object.entries(packageData.joints).flatMap(([key, joint]) => {
        const node = doc.getElementById(joint.id);
        const rotation = joint.rotation;
        return !node || node.getAttribute('data-parent') !== joint.parent ||
          Number(node.getAttribute('data-pivot-x')) !== joint.pivot[0] ||
          Number(node.getAttribute('data-pivot-y')) !== joint.pivot[1] ||
          Number(node.getAttribute('data-rotation-min')) !== rotation[0] ||
          Number(node.getAttribute('data-rotation-max')) !== rotation[1] ? [key] : [];
      });
      return { missingParts, badParents, badLayers, badJoints };
    }, manifest);
    check('manifest part IDs exist in SVG', embeddedManifestCheck.missingParts.length === 0, embeddedManifestCheck.missingParts.join(', '));
    check('manifest parent-child hierarchy matches SVG', embeddedManifestCheck.badParents.length === 0, embeddedManifestCheck.badParents.join(', '));
    check('manifest layer order matches SVG groups', embeddedManifestCheck.badLayers.length === 0, embeddedManifestCheck.badLayers.join(', '));
    check('manifest pivots and constraints match SVG attributes', embeddedManifestCheck.badJoints.length === 0, embeddedManifestCheck.badJoints.join(', '));

    const baseDeterminants = {};
    for (const [key, joint] of Object.entries(manifest.joints)) {
      const input = page.locator(`[data-input="${key}"]`);
      check(`${key}: slider limits and label match manifest`,
        Number(await input.getAttribute('min')) === joint.rotation[0] &&
        Number(await input.getAttribute('max')) === joint.rotation[1] &&
        await page.locator(`label[for="ctrl-${key}"]`).textContent() === joint.label,
        `range/label differs for ${key}`);

      const values = angleCases(joint.rotation);
      for (const angle of values) {
        const before = await page.evaluate(() => Object.fromEntries(
          [...document.querySelector('#rakaObject').contentDocument.querySelectorAll('[data-joint][data-pivot-x]')]
            .map(node => [node.id, node.getAttribute('transform')]),
        ));
        await setSlider(page, key, angle);
        const state = await readJointState(page, key);
        const expectedTransform = `rotate(${angle} ${joint.pivot[0]} ${joint.pivot[1]})`;
        const output = await page.locator(`[data-output="${key}"]`).textContent();
        const after = await page.evaluate(() => Object.fromEntries(
          [...document.querySelector('#rakaObject').contentDocument.querySelectorAll('[data-joint][data-pivot-x]')]
            .map(node => [node.id, node.getAttribute('transform')]),
        ));
        const unrelatedChanged = Object.entries(before).filter(([id, value]) => id !== joint.id && after[id] !== value).map(([id]) => id);
        check(`${key}: ${angle}° reaches its SVG group and display`, state.transform === expectedTransform && output === `${angle}°`, `${state.transform}; output=${output}`);
        check(`${key}: pivot stays fixed and covered at ${angle}°`, state.pivotError < 0.5 && Boolean(state.pivotHit) && !['svg', 'raka-rig', 'pelvis'].includes(state.pivotHit), `pivot error=${state.pivotError.toFixed(3)}, hit=${state.pivotHit}`);
        check(`${key}: rotation does not scale or shear at ${angle}°`, Math.abs(state.determinant - (baseDeterminants[key] ?? state.determinant)) < 0.01, `determinant=${state.determinant}`);
        if (baseDeterminants[key] === undefined) baseDeterminants[key] = state.determinant;
        check(`${key}: isolated slider does not alter other joint transforms at ${angle}°`, unrelatedChanged.length === 0, unrelatedChanged.join(', '));
        await page.locator('#rakaObject').screenshot({ path: path.join(FRAME_DIR, `${key}-${angle}.png`) });
      }
      await setSlider(page, key, 0);
    }

    const presetButtons = await page.locator('[data-preset]').evaluateAll(buttons => buttons.map(button => ({ id: button.dataset.preset, label: button.textContent.trim() })));
    check('all four documented presets are present', ['neutral', 'wave', 'arms-up', 'open'].every(id => presetButtons.some(button => button.id === id)), JSON.stringify(presetButtons));
    const presetChecks = [];
    for (const presetId of ['neutral', 'wave', 'arms-up', 'open']) {
      const contamination = Object.fromEntries(Object.entries(manifest.joints).map(([key, joint], index) => [key, index % 2 ? joint.rotation[0] : joint.rotation[1]]));
      for (const [key, value] of Object.entries(contamination)) await setSlider(page, key, value);
      const startAngles = await readAngles(page);
      await page.locator(`[data-preset="${presetId}"]`).click();
      if (presetId === 'wave') {
        await page.waitForTimeout(110);
        const intermediate = await page.locator('#rakaObject').evaluate(object => Number(object.contentDocument.getElementById('right-shoulder').getAttribute('transform').match(/rotate\(([-\d.]+)/)[1]));
        const from = startAngles.rightShoulder, to = manifest.presets.wave.rightShoulder;
        check('preset transition interpolates through an intermediate pose', Math.abs(intermediate - from) > 0.5 && intermediate > Math.min(from, to) && intermediate < Math.max(from, to), `${from}° → ${intermediate}° → ${to}°`);
        await page.locator('#rakaObject').screenshot({ path: path.join(EVIDENCE_DIR, 'transition-mid-wave.png') });
      }
      await page.waitForTimeout(470);
      const current = await readAngles(page);
      const target = { ...Object.fromEntries(Object.keys(manifest.joints).map(key => [key, 0])), ...manifest.presets[presetId] };
      const mismatch = Object.entries(target).filter(([key, value]) => Math.abs(current[key] - value) > 0.5).map(([key]) => `${key}:${current[key]}!=${target[key]}`);
      check(`preset ${presetId} sets every joint without inherited angles`, mismatch.length === 0, mismatch.join(', '));
      presetChecks.push({ id: presetId, current, target });
      await page.locator('#rakaObject').screenshot({ path: path.join(EVIDENCE_DIR, `pose-${presetId}.png`) });
    }

    await page.locator('#waveMotion').click();
    await page.waitForTimeout(500);
    const waveSamples = [];
    for (let index = 0; index < 48; index += 1) {
      waveSamples.push(await page.locator('#rakaObject').evaluate(object => Number(object.contentDocument.getElementById('right-wrist').getAttribute('transform').match(/rotate\(([-\d.]+)/)[1])));
      await page.waitForTimeout(100);
    }
    const signChanges = waveSamples.slice(1).reduce((count, value, index) => count + (Math.sign(value) !== Math.sign(waveSamples[index]) ? 1 : 0), 0);
    const waveBounds = manifest.joints.rightWrist.rotation;
    check('wave animation completes three cycles within wrist constraints', signChanges >= 5 && Math.min(...waveSamples) >= waveBounds[0] && Math.max(...waveSamples) <= waveBounds[1], `sign changes=${signChanges}; range=${Math.min(...waveSamples)}..${Math.max(...waveSamples)}`);
    await page.locator('#waveMotion').click();
    const stoppedValue = await page.locator('#rakaObject').evaluate(object => object.contentDocument.getElementById('right-wrist').getAttribute('transform'));
    await page.waitForTimeout(350);
    const stableValue = await page.locator('#rakaObject').evaluate(object => object.contentDocument.getElementById('right-wrist').getAttribute('transform'));
    check('wave stop cancels repeated updates', stoppedValue === stableValue, `${stoppedValue} → ${stableValue}`);
    await page.locator('#reset').click();
    await page.waitForTimeout(470);
    check('reset returns all controls and transforms to zero', Object.values(await readAngles(page)).every(value => value === 0), JSON.stringify(await readAngles(page)));

    const combination = { torso: 8, head: -13, leftShoulder: 60, leftElbow: -80, leftWrist: 24, rightShoulder: -48, rightElbow: 75, rightWrist: -21, leftHip: 18, leftKnee: 32, leftAnkle: -12, rightHip: -16, rightKnee: -28, rightAnkle: 10 };
    for (const [key, angle] of Object.entries(combination)) await setSlider(page, key, angle);
    const combinedPivotErrors = [];
    for (const [key] of Object.entries(manifest.joints)) {
      const state = await readJointState(page, key);
      if (state.pivotError >= 0.5 || !state.pivotHit) combinedPivotErrors.push(`${key}:${state.pivotError.toFixed(2)}/${state.pivotHit}`);
    }
    check('combined parent-child rotations keep every joint anchored and covered', combinedPivotErrors.length === 0, combinedPivotErrors.join(', '));
    await page.locator('#rakaObject').screenshot({ path: path.join(EVIDENCE_DIR, 'pose-combined-stress.png') });

    if (structure.importControl && structure.jsonExportControl) {
      const expectedAngles = await readAngles(page);
      const jsonPath = await downloadFile(page, 'exportPoseJson', 'Raka_Pose_exported.json');
      const poseData = JSON.parse(fs.readFileSync(jsonPath, 'utf8'));
      check('pose JSON export stores all finite joint angles', poseData.character_id === manifest.character_id && Object.keys(poseData.angles || {}).length === 14 && Object.values(poseData.angles).every(Number.isFinite), JSON.stringify(poseData));
      const svgPath = await downloadFile(page, 'downloadPose', 'Raka_Pose_exported.svg');
      const svgText = fs.readFileSync(svgPath, 'utf8');
      const svgMismatches = Object.entries(expectedAngles).filter(([key, angle]) => {
        const id = manifest.joints[key].id;
        const tag = svgText.match(new RegExp(`<g\\b(?=[^>]*\\bid="${id}")[^>]*>`));
        const serializedAngle = tag?.[0].match(/transform="rotate\(([-\d.]+)/)?.[1];
        return !serializedAngle || Math.abs(Number(serializedAngle) - angle) > 0.01;
      }).map(([key]) => key);
      check('SVG pose export matches visible joint angles', svgMismatches.length === 0, svgMismatches.join(', '));
      for (const key of Object.keys(manifest.joints)) await setSlider(page, key, 0);
      await page.locator('#poseFile').setInputFiles(jsonPath);
      await page.waitForTimeout(470);
      const restored = await readAngles(page);
      const restoreMismatches = Object.entries(expectedAngles).filter(([key, value]) => Math.abs(restored[key] - value) > 0.5).map(([key]) => key);
      check('imported pose JSON restores the exact exported angles', restoreMismatches.length === 0, restoreMismatches.join(', '));
      await page.locator('#rakaObject').screenshot({ path: path.join(EVIDENCE_DIR, 'pose-json-restored.png') });

      const invalidPoses = [
        { ...poseData, manifest_version: manifest.version + 1 },
        { ...poseData, angles: Object.fromEntries(Object.entries(poseData.angles).filter(([key]) => key !== 'head')) },
        { ...poseData, angles: { ...poseData.angles, head: manifest.joints.head.rotation[1] + 1 } },
      ];
      for (const [index, invalidPose] of invalidPoses.entries()) {
        await page.locator('#poseNotice').evaluate(node => { node.removeAttribute('data-state'); node.textContent = ''; });
        await page.locator('#poseFile').setInputFiles({
          name: `invalid-pose-${index + 1}.json`, mimeType: 'application/json',
          buffer: Buffer.from(JSON.stringify(invalidPose)),
        });
        await page.waitForFunction(() => {
          const notice = document.querySelector('#poseNotice');
          return notice?.dataset.state === 'error' && notice.textContent.startsWith('Gagal memuat pose:');
        });
        const rejected = await page.locator('#poseNotice').getAttribute('data-state') === 'error';
        const afterRejectedImport = await readAngles(page);
        const unchanged = Object.entries(expectedAngles).every(([key, value]) => Math.abs(afterRejectedImport[key] - value) < 0.5);
        check(`invalid pose ${index + 1} is rejected without changing the rig`, rejected && unchanged, `rejected=${rejected}; unchanged=${unchanged}`);
      }
    }

    const viewportChecks = [];
    for (const viewport of [{ width: 1440, height: 1000 }, { width: 820, height: 1100 }, { width: 390, height: 844 }]) {
      await page.setViewportSize(viewport);
      await page.waitForTimeout(100);
      const metrics = await page.evaluate(() => ({
        width: window.innerWidth, scrollWidth: document.documentElement.scrollWidth,
        stageWidth: document.querySelector('.stage').getBoundingClientRect().width,
        stageHeight: document.querySelector('.stage').getBoundingClientRect().height,
        controlsVisible: [...document.querySelectorAll('.control input')].every(input => input.getBoundingClientRect().width > 0),
      }));
      const okay = metrics.scrollWidth <= metrics.width + 1 && metrics.stageWidth > 0 && metrics.stageHeight > 0 && metrics.controlsVisible;
      check(`responsive viewer at ${viewport.width}×${viewport.height}`, okay, JSON.stringify(metrics));
      viewportChecks.push({ viewport, metrics, okay });
      await page.screenshot({ path: path.join(EVIDENCE_DIR, `viewport-${viewport.width}.png`), fullPage: true });
    }

    check('browser has no runtime, console, or failed-request errors', browserErrors.length === 0, browserErrors.join('\n'));
    const report = {
      browser: await browser.version(),
      playwright: require('playwright/package.json').version,
      node: process.version,
      pageUrl,
      generatedAt: new Date().toISOString(),
      result: errors.length ? 'FAIL' : 'PASS',
      summary: { passed: results.filter(result => result.status === 'PASS').length, failed: errors.length },
      results,
      failures: errors,
      structure,
      viewportChecks,
      presets: presetChecks,
      waveSamples,
    };
    fs.writeFileSync(REPORT_PATH, `${JSON.stringify(report, null, 2)}\n`);
    console.log(JSON.stringify(report.summary));
    for (const failure of errors) console.error(`FAIL ${failure}`);
    if (errors.length) process.exitCode = 1;
  } finally {
    await browser.close();
    await new Promise(resolve => server.close(resolve));
  }
}

main().catch(error => {
  fs.writeFileSync(REPORT_PATH, `${JSON.stringify({ result: 'ERROR', message: error.stack }, null, 2)}\n`);
  console.error(error);
  process.exitCode = 1;
});
