'use strict';

/* ============================================================
   _template.js  starter for a new StitchConfig
   ============================================================
   Copy this file to configs/<your-id>.js, rename the id/label/clientCodes,
   then fill in the fields described below. Not loaded by index.html  add
   an explicit <script> tag to enable a new config.
*/

/**
 * @typedef {Object} StitchConfig
 * @property {string} id  unique slug (drives cache keys + dropdown)
 * @property {string} label  short human label
 * @property {string} [description]  one-sentence summary
 * @property {string[]} clientCodes  which client codes can see this config
 * @property {{primary: InputDef, secondary: InputDef}} inputs
 * @property {MatchRule[]} matchStrategy
 * @property {(candidates: Object[], secondaryRow: Object) => {row: Object, score: number}} [tiebreaker]
 * @property {(row: Object) => boolean} [testRowFilter]
 * @property {Object<string, (primary: Object, secondary: Object, matchInfo: Object) => any>} [derivedFields]
 * @property {ColumnDef[]} defaultColumns
 * @property {OutputSheet[]} outputSheets
 * @property {DashboardSpec} dashboard
 */

/**
 * @typedef {Object} InputDef
 * @property {string} label  drop-zone label
 * @property {string[]} requiredCols  CSV columns that must be present
 */

/**
 * @typedef {Object} MatchRule
 * @property {string} method  display name (e.g. "ContactID")
 * @property {string} primaryCol  CSV column on the primary input
 * @property {string} secondaryCol  CSV column on the secondary input
 * @property {'trim'|'email'|'phone'|((s:string)=>string)} normalize
 */

/**
 * @typedef {Object} ColumnDef
 * @property {string} key  stable key (drives column-picker storage)
 * @property {string} label  human-readable header
 * @property {'primary'|'secondary'|'derived'} source
 * @property {string} sourceField  CSV column name (or derivedFields key)
 * @property {string} [fallbackField]  fall back to the other input if blank
 */

/**
 * @typedef {Object} OutputSheet
 * @property {string} name  sheet tab name
 * @property {(wb: any, ctx: Object) => Object} builder  receives workbook + shared ctx
 */

/**
 * @typedef {Object} DashboardSpec
 * @property {Array<Object>} kpis
 * @property {Array<Object>} charts
 * @property {Array<Object>} filters
 */

/* Skeleton  uncomment and adapt:

(function() {
  window.registerStitchConfig({
    id: 'your-config-id',
    label: 'Primary  Secondary',
    description: 'Stitch ...',
    clientCodes: ['your-client'],

    inputs: {
      primary:   { label: 'Primary CSV',   requiredCols: [] },
      secondary: { label: 'Secondary CSV', requiredCols: [] },
    },

    matchStrategy: [
      // { method: 'ContactID', primaryCol: 'Contact ID', secondaryCol: 'Contact ID', normalize: 'trim' },
    ],

    derivedFields: {
      // your_field: (primary, secondary, matchInfo) => ...,
    },

    defaultColumns: [
      // { key: 'first_name', label: 'First Name', source: 'primary', sourceField: 'First Name' },
    ],

    outputSheets: [
      // { name: 'Stitched Data', builder: (wb, ctx) => { ... } },
    ],

    dashboard: { kpis: [], charts: [], filters: [] },
  });
})();

*/
