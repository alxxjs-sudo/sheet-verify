/**
 * The Program / Treaty Selection template: everything specific to it.
 *
 * Adding another template means adding another folder like this one. Nothing in
 * ../../ changes, and neither does sheet-verify.
 */
import { projectRows, COLUMNS as PAYLOAD_COLUMNS } from './from-payload.mjs';
import { projectTableRows, COLUMNS as UI_COLUMNS } from './from-ui.mjs';

export default {
  sheet: 'Treaties',
  headerRow: 3,

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
    },
  },
};
