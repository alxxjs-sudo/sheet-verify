/**
 * The Overrides template: everything specific to it.
 *
 * One download kind, two shapes. Overriding a marketplace layer produces a
 * "MarketPlace Layer" sheet; overriding a MetaRisk treaty produces a "MetaRisk
 * Treaty" sheet, with different columns, a different key and different derived
 * values. They arrive in the same folder because they are the same feature, so
 * this descriptor names both variants and the workbook decides which it is.
 *
 * The contract here is stated differently from the program selection template.
 * That sheet paints its editable columns yellow; this one paints nothing and
 * marks them with a trailing " *" in the header instead. A1 says only "*Do not
 * alter the existing "Type" row and sheet name, as it will make the template
 * unrecognizable by the application" -- which makes the Type row and the sheet
 * name part of the contract too, and both are checked.
 */
import { COLUMNS as ID_COLUMNS, projectRows as projectIds, declares as payloadDeclares }
  from './from-payload.mjs';
import { columnsFor, projectRowsFor, declares as tableDeclares, isComplete }
  from './from-table.mjs';

/** Two strings differ, said the way the sheet says it. */
const differs = (a, b) => (String(a ?? '').trim() === String(b ?? '').trim() ? 'No' : 'Yes');

/** Which column carries the id the request asked for. */
const KEY_COLUMN = {
  layerExternalId: 'GCMP Layer ID',
  treatyId: 'MetaRisk Treaty ID',
};

/**
 * The sources, which are the same two questions for either variant: were these
 * the rows asked for, and do the values match what the screen showed.
 */
const sources = (keyField) => [
  {
    name: 'payload',
    file: 'payload_data.json',
    label: 'the download request',
    columns: ID_COLUMNS,
    project: projectIds,
    declares: payloadDeclares,
    key: (cell) => String(cell(KEY_COLUMN[keyField])),
  },
  {
    name: 'table',
    file: 'table_data.json',
    label: 'the table on screen',
    columns: columnsFor,
    project: projectRowsFor(keyField),
    declares: tableDeclares,
    complete: isComplete,
    // Read off a rendered page, which collapses runs of whitespace before it
    // paints: "Layer 2 -  85M XS 35M" reaches the screen with one space. The
    // request body carries the string exactly and is compared exactly.
    collapseWhitespace: true,
    key: (cell) => String(cell(KEY_COLUMN[keyField])),
  },
];

export default {
  headerRow: 4,

  variants: [
    {
      sheet: 'MarketPlace Layer',
      // The value in C2, which the sheet says must not be altered.
      declared: 'marketplaceLayer',
      rowMarker: 'GCMP Layer ID',
      sources: sources('layerExternalId'),

      markers: {
        suffix: ' *',
        columns: [
          'Edison Layer Number of Reinstatements',
          'Edison Layer Reinstatement Description',
          'Edison Layer Occurrence Limit',
          'Edison Layer Aggregate Limit',
          'Edison Client Level GeoScope',
          'Edison Layer Level GeoScope',
          'Edison Segment',
        ],
      },

      /**
       * Two Yes/No columns the sheet works out from its own rows. Nothing
       * carries them -- not the request, not the screen -- so this is the only
       * check they will ever get, and getting one wrong tells a user a field
       * was left alone when it was changed.
       */
      derived: [
        {
          column: 'Edison Client Level GeoScope differs from Property COE GeoScope',
          from: ['Edison Client Level GeoScope *', 'Property COE GeoScope'],
          value: differs,
        },
        {
          column: 'Edison Segment differs from Property COE Segment',
          from: ['Edison Segment *', 'Property COE Segment'],
          value: differs,
        },
      ],

      unverifiable: {
        'the currency of an override; the screen shows the amount without one':
          ['Edison Layer Occurrence Limit Currency', 'Edison Layer Aggregate Limit Currency'],
      },
    },

    {
      sheet: 'MetaRisk Treaty',
      declared: 'metaRiskTreaty',
      rowMarker: 'MetaRisk Treaty ID',
      sources: sources('treatyId'),

      markers: {
        suffix: ' *',
        columns: [
          'Edison Treaty Number of Reinstatements',
          'Edison Treaty Reinstatement Description',
          'Edison Treaty Occurrence Limit',
          'Edison Treaty Aggregate Limit',
          'Edison Treaty Premium',
        ],
      },

      unverifiable: {
        'a currency; the screen states it inside the amount, so it is checked there': [
            'MetaRisk Treaty Occurrence Limit Currency',
            'MetaRisk Treaty Aggregate Limit Currency',
            'MetaRisk Treaty Premium Currency',
            'Edison Treaty Occurrence Limit Currency',
            'Edison Treaty Aggregate Limit Currency',
            'Edison Treaty Premium Currency',
          ],
      },
    },
  ],

  /**
   * The one fill this template uses: the header band. No column list, so every
   * column must carry it -- a header that lost its colour is a header the app
   * may no longer recognise, and A1 is explicit that the sheet has to stay
   * recognisable.
   */
  fills: {
    header: { argb: 'FF006D9E', row: 'header' },
  },
};
