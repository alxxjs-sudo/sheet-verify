/**
 * Checks every internal link across the README, the changelog and `docs/`.
 *
 * The documentation is split across seven files that link to each other by
 * relative path and by heading anchor, and both break silently. Renaming a
 * heading leaves every link to it pointing at nothing, and a reader finds out
 * by landing at the top of the wrong page -- if they say anything at all.
 * Moving prose between files, which is how these documents get maintained,
 * breaks several at once.
 *
 * Only internal links are followed. An external URL failing is somebody else's
 * outage and not a reason to fail a build here.
 *
 * Links inside fenced code blocks are ignored, and so are headings inside them:
 * a fence showing what `report.md` looks like is a sample, not a document
 * structure, and counting its headings as anchors would let a real broken link
 * pass because a code sample happened to contain the same words.
 */
import { readdir, readFile } from 'node:fs/promises';
import { existsSync } from 'node:fs';
import { dirname, join, normalize, relative, resolve, sep } from 'node:path';

const ROOT = resolve(process.argv.find((a) => !a.startsWith('-') && a.endsWith('/')) ?? '.');

/** Lines outside fenced code blocks. */
function prose(text) {
  const out = [];
  let fenced = false;
  for (const line of text.split('\n')) {
    if (line.trimStart().startsWith('```')) { fenced = !fenced; continue; }
    if (!fenced) out.push(line);
  }
  return out;
}

/** GitHub's heading -> anchor rule, near enough for our own headings. */
const slug = (h) =>
  h.replace(/`/g, '').toLowerCase().replace(/[^a-z0-9 _-]/g, '').trim().replace(/ /g, '-');

const files = ['README.md', 'CHANGELOG.md'];
for (const name of (await readdir(join(ROOT, 'docs')).catch(() => [])).sort()) {
  if (name.endsWith('.md')) files.push(join('docs', name));
}

const anchors = new Map();
const bodies = new Map();
for (const f of files) {
  const text = await readFile(join(ROOT, f), 'utf8');
  const lines = prose(text);
  // Inline code spans go too: prose explaining the link syntax writes that
  // syntax down, and the first run of this reported that sentence as broken.
  bodies.set(f, lines.join('\n').replace(/`[^`\n]*`/g, ''));
  anchors.set(
    normalize(f).split(sep).join('/'),
    new Set(lines.map((l) => /^#{1,6}\s+(.*?)\s*$/.exec(l)?.[1]).filter(Boolean).map(slug)),
  );
}

// Kept by file rather than as one flat list. A move between documents breaks
// several links at once, and grouped they read as the one edit they are
// instead of as five unrelated faults.
const problems = new Map();
const note = (f, line) => problems.set(f, [...(problems.get(f) ?? []), line]);
let checked = 0;
let broken = 0;

for (const f of files) {
  for (const m of bodies.get(f).matchAll(/\]\(([^)]+)\)/g)) {
    const target = m[1].trim();
    if (/^(https?:|mailto:)/.test(target)) continue;
    checked++;

    const [path, anchor] = target.split('#');
    let key = normalize(f).split(sep).join('/');

    if (path) {
      const resolved = normalize(join(dirname(join(ROOT, f)), path));
      if (!existsSync(resolved)) {
        note(f, `${target}\n       no such file`);
        broken++;
        continue;
      }
      key = relative(ROOT, resolved).split(sep).join('/');
    }

    // A link into a file that is not one of ours -- source, a config -- is
    // checked for existence only. There are no headings to look for.
    if (anchor && anchors.has(key) && !anchors.get(key).has(anchor)) {
      note(f, `${target}\n       no heading "${anchor}" in ${key}`);
      broken++;
    }
  }
}

const headings = [...anchors.values()].reduce((total, set) => total + set.size, 0);

if (broken) {
  console.error('');
  for (const [f, lines] of problems) {
    console.error(`${f}  (${lines.length})`);
    for (const line of lines) console.error(`  ->  ${line}`);
  }
  console.error(
    `\n${broken} broken link(s) in ${problems.size} file(s)` +
    ` — of ${checked} internal link(s) across ${files.length} file(s)`,
  );
  process.exit(1);
}
console.log(
  `${checked} internal link(s) and ${headings} heading anchor(s)` +
  ` across ${files.length} file(s), all resolve`,
);
