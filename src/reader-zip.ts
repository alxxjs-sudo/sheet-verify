import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, extname, join } from 'node:path';
import JSZip from 'jszip';
import type {
  ResolvedSpec, SheetModel, SheetReader, TableRequest, WorkbookReader,
} from './types.js';

/**
 * A zip archive, read as a workbook whose sheets are its members.
 *
 * This is the same move the CSV reader already makes -- "a CSV is one table,
 * so it presents itself as a one-sheet workbook" -- one level up. A workbook
 * in this codebase is nothing more than a list of named tables, and an archive
 * of CSVs is exactly that. Presenting it as one means every layer above gets
 * the archive for free: members pair by name, a member the golden had and the
 * report does not is a *removed sheet* and so a defect, a new one is reported
 * for review, and keys, tolerances and the layer 2 sweep all address a member
 * the way they address any other sheet.
 *
 * Nothing about the archive itself is compared, deliberately. Two zips of
 * identical content differ byte-for-byte because each carries the modified
 * time of the moment it was written -- every archive in one real tree did,
 * while every member inside was identical. Reporting that would be reporting
 * the clock.
 *
 * Members are extracted to a temporary folder rather than read from memory,
 * because the readers that handle them take a path: delegating by path is what
 * makes a zip of .xlsx work the day someone produces one, without this file
 * knowing anything about the formats inside it.
 */
export class ZipReader implements WorkbookReader {
  readonly extensions = ['.zip'];

  /**
   * How a member finds its own reader. Passed in rather than imported so this
   * file never reaches back into the registry that owns it -- the registry
   * constructs this, and a cycle between the two would be a needless trap.
   */
  constructor(private readonly readerFor: (path: string) => SheetReader) {}

  /**
   * Members a registered reader can handle, in archive order.
   *
   * Directory entries and anything unreadable are skipped rather than
   * reported: an archive is free to carry a manifest or a licence file, and
   * neither is a table. What is *not* skipped is a member on one side with no
   * partner on the other -- that is a sheet difference, and the comparison
   * above decides what it means.
   */
  private async members(path: string): Promise<{ name: string; body: Buffer }[]> {
    const zip = await JSZip.loadAsync(await readBinary(path));
    const out: { name: string; body: Buffer }[] = [];
    for (const name of Object.keys(zip.files)) {
      const entry = zip.files[name]!;
      if (entry.dir) continue;
      if (!this.handled(name)) continue;
      out.push({ name, body: await entry.async('nodebuffer') });
    }
    return out;
  }

  private handled(name: string): boolean {
    try {
      this.readerFor(name);
      return true;
    } catch {
      return false;
    }
  }

  async readWorkbook(
    path: string,
    tablesFor: (sheet: string) => TableRequest[],
  ): Promise<{ sheets: string[]; models: Map<string, SheetModel> }> {
    const members = await this.members(path);
    const sheets = members.map((m) => m.name);
    const models = new Map<string, SheetModel>();

    // Only extract what something asked for. An archive of forty members with
    // one configured costs one file on disk, which is the same bargain the
    // Excel reader makes when it lists every tab and builds one.
    const wanted = members.filter((m) => tablesFor(m.name).length > 0);
    if (!wanted.length) return { sheets, models };

    const dir = await mkdtemp(join(tmpdir(), 'sheet-verify-zip-'));
    try {
      for (const m of wanted) {
        // Flattened, and named for its position: a member can sit in a folder
        // inside the archive, and the name is only ever used to give the file
        // an extension its reader will recognise.
        const file = join(dir, `${models.size}-${basename(m.name)}`);
        await writeFile(file, m.body);
        const reader = this.readerFor(m.name);
        for (const req of tablesFor(m.name)) {
          models.set(req.key, { ...(await reader.read(file, req.spec)), table: req.table });
        }
      }
    } finally {
      await rm(dir, { recursive: true, force: true });
    }

    return { sheets, models };
  }

  /**
   * The single-table form, for an archive holding one.
   *
   * `unused.zip` in one real tree holds exactly one `unused_policy.csv`, and
   * pointing `verifySheet` at it should mean what it obviously means. More
   * than one member has no obvious answer, so it says so and names them rather
   * than picking.
   */
  async read(path: string, spec: ResolvedSpec): Promise<SheetModel> {
    const members = await this.members(path);
    if (members.length === 1) {
      const only = members[0]!;
      const dir = await mkdtemp(join(tmpdir(), 'sheet-verify-zip-'));
      try {
        const file = join(dir, basename(only.name));
        await writeFile(file, only.body);
        return await this.readerFor(only.name).read(file, spec);
      } finally {
        await rm(dir, { recursive: true, force: true });
      }
    }
    throw new Error(
      members.length === 0
        ? `sheet-verify: "${basename(path)}" holds nothing any reader handles`
        : `sheet-verify: "${basename(path)}" holds ${members.length} tables `
          + `(${members.map((m) => m.name).join(', ')}), so which one to read is a guess. `
          + 'Compare it as a workbook, where every member is a sheet.',
    );
  }
}

/** Read as bytes. Separate so the import stays out of the class body. */
async function readBinary(path: string): Promise<Buffer> {
  const { readFile } = await import('node:fs/promises');
  return readFile(path);
}

/** Whether a path looks like an archive this reader handles. */
export const isArchive = (path: string): boolean => extname(path).toLowerCase() === '.zip';
