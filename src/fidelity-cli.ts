#!/usr/bin/env node
/**
 * Usage: sheet-fidelity <template.xlsx> [more.xlsx ...]
 *
 * Exits non-zero if any template loses content on a no-op round-trip, so it
 * can gate a build.
 */
import { checkFidelity, formatFidelity, type FidelityReport } from './fidelity.js';

const files = process.argv.slice(2).filter((a) => !a.startsWith('-'));

if (!files.length) {
  console.error('Usage: sheet-fidelity <template.xlsx> [more.xlsx ...]');
  process.exit(2);
}

const reports: FidelityReport[] = [];
for (const f of files) {
  try {
    reports.push(await checkFidelity(f));
  } catch (e) {
    reports.push({
      file: f, ok: false,
      findings: [{ severity: 'critical', detail: `could not be processed: ${(e as Error).message}` }],
    });
  }
}

console.log(formatFidelity(reports));
process.exit(reports.every((r) => r.ok) ? 0 : 1);
