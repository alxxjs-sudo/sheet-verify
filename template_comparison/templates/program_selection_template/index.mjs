/**
 * The Program / Treaty Selection template: everything specific to it.
 *
 * Adding another template means adding another folder like this one. Nothing in
 * ../../ changes, and neither does sheet-verify.
 */
import { projectRows, COLUMNS as PAYLOAD_COLUMNS } from './from-payload.mjs';
import { projectTableRows, COLUMNS as UI_COLUMNS } from './from-ui.mjs';
import { ROLEPLAY_BLOCK } from './roleplay-block.mjs';
import { ARITHMETIC, COVERED_BY_ARITHMETIC } from './arithmetic.mjs';

export default {
  sheet: 'Treaties',
  headerRow: 3,

  // Every real row names its treaty; the trailing rows of the sheet do not.
  rowMarker: 'Treaty Name',

  /**
   * Two independent readings of the same template.
   *
   * The payload is what the client sent; the table is what a person could see.
   * Checking against both means a bug that corrupts the request is caught by one
   * and a bug that misrenders the screen by the other. A single source would
   * call either one correct, having no second opinion.
   */
  sources: [
    {
      name: 'payload',
      file: 'payload_data.json',
      label: 'the download payload',
      columns: PAYLOAD_COLUMNS,
      project: projectRows,
      // Both sides carry the treaty id, so identity is exact.
      key: (cell) => String(cell('Treaty Id')),
    },
    {
      name: 'table',
      file: 'table_data.json',
      label: 'the table on screen',
      columns: UI_COLUMNS,
      project: projectTableRows,
      // The UI never shows the treaty id. NOT company + name, which collides:
      // one company can carry the same treaty name in several programs.
      key: (cell) => `${cell('Edison Program Name')}|${cell('Treaty Name')}`,
    },
  ],

  /**
   * The fills are the contract: A1 of this template reads "*Only fields
   * highlighted yellowish are editable in the UI". Keyed by meaning rather than
   * by colour, so a repaint changes one hex here instead of a column list.
   */
  fills: {
    editable: {
      argb: 'FFFFFF99',
      columns: [
        'Treaty Aggregate Limit',
        'Treaty Premium',
        'Treaty Participation',
        'Treaty Cap Limit',
        'Treaty Discount',
        'Treaty Tags',
      ],
    },
    divider: {
      argb: 'FFFF0000',
      row: 'header',
      columns: ['ROLePlay Data ->'],
      // Present only when the block behind it is. See `blocks` below: the
      // download asks whether to include ROLePlay data, so the divider is
      // absent by request rather than by fault. Absent is fine; absent when the
      // block IS there, or present without its red, is not.
      optional: true,
    },
  },

  /**
   * Columns that arrive together or not at all.
   *
   * Downloading this template opens a prompt -- "Would you like to include
   * ROLePlay data in the template for advanced filtration?" -- and the answer
   * decides whether the divider and the 144 columns behind it are written. So
   * their absence is a choice, not a defect, and their PARTIAL presence is a
   * defect that would otherwise pass unnoticed: every column that did arrive
   * would be correct.
   */
  blocks: {
    'ROLePlay data': {
      lead: 'ROLePlay Data ->',
      columns: ROLEPLAY_BLOCK,
      // The prompt's answer is recorded in the request, so the request is the
      // authority on whether the block belongs. Letting the divider decide only
      // proves the sheet agrees with itself: a download asked for ROLePlay data
      // that came back with none of it is perfectly self-consistent and exactly
      // the bug worth catching.
      requested: {
        file: 'payload_data.json',
        value: (data) => data.includeRoleplayData === true,
      },
    },
  },

  /**
   * Columns no source can speak for, and why.
   *
   * Everything here is reported as excused rather than passed. A column that is
   * neither checked nor listed here fails the case -- the whole point being
   * that "nobody looked" must never print the same as "it agreed".
   */
  /**
   * ROLePlay columns that are arithmetic on their neighbours. No capture
   * carries these figures, but the sheet computes some of them from columns
   * that ARE verified, and ties the rest to each other.
   */
  derived: ARITHMETIC,

  unverifiable: {
    'modelling output the server joins in; in neither capture':
      ROLEPLAY_BLOCK.filter((c) => !COVERED_BY_ARITHMETIC.includes(c)),
    'a divider label rather than data -- its fill is checked instead': ['ROLePlay Data ->'],
  },

  /**
   * Header names the template writes twice, on purpose.
   *
   * The ROLePlay block repeats the program identity columns it is keyed on, so
   * "Edison Program Name" appears at E3 and again at AI3. Lookup is by name and
   * the first wins, which leaves the second unreachable. Declared here so it
   * reads as known rather than as an oversight -- and so a name that starts
   * repeating for some other reason is reported instead of absorbed.
   */
  duplicateHeaders: ['Edison Program Name'],
};
