// Default square cell size in px for paint fallback and reset. Matches
// the "Resize cells to 18×18" menu item so painting/resetting lands on
// the same baseline that menu action sets up.
const CELL_PX = 18;

/* Maps the sidebar's border-style strings to Sheets' native BorderStyle
   enum. Hoisted to module scope so both paintArt (writer) and pasteShape
   (re-applier) can use it without duplicating the table. */
const BORDER_STYLE_MAP = {
  'thin':   SpreadsheetApp.BorderStyle.SOLID,
  'medium': SpreadsheetApp.BorderStyle.SOLID_MEDIUM,
  'thick':  SpreadsheetApp.BorderStyle.SOLID_THICK,
  'dashed': SpreadsheetApp.BorderStyle.DASHED,
  'dotted': SpreadsheetApp.BorderStyle.DOTTED,
  'double': SpreadsheetApp.BorderStyle.DOUBLE,
};

/* ----------------------------------------------------------------------
   Border reading via the Sheets Advanced Service

   Apps Script's built-in Range API can write borders but not read them
   — there's no getBorder counterpart to setBorder. To make copy/paste
   carry borders, we read them directly off the spreadsheet via the
   Sheets REST API (exposed in Apps Script as the "Sheets" Advanced
   Service). This is the source of truth — no separate state to keep
   in sync — and it also picks up borders the user added manually via
   the Sheets toolbar, not just ones this tool wrote.

   Setup requirement: the Sheets API service must be enabled once per
   Apps Script project (Services panel → "Google Sheets API"). The
   manifest remembers it across code pastes, so it only needs to be
   clicked the first time a new project is set up. If the service
   isn't enabled, readBordersFromAPI_ returns null and copy degrades
   to bg/text/font only — no crash. ----------------------------------- */

/* Convert a 1-indexed column number to its Sheets A1 letter (1 → A,
   26 → Z, 27 → AA, 702 → ZZ, 703 → AAA, …). Pure string math — no
   API round-trip — so it's safe to call in tight per-cell loops.
   Used by paintArt's border-batching path to build A1 lists for
   RangeList.setBorder. */
function colNumToLetter_(n) {
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/* Convert a Sheets-API rgb color ({red, green, blue} as 0–1 floats,
   plus optional alpha) to the #rrggbb string the setBorder API wants.
   Missing channels default to 0, which matches the API's convention
   that omitted = 0. */
function apiColorToHex_(rgb) {
  if (!rgb) return '#000000';
  const to255 = v => Math.max(0, Math.min(255, Math.round((v || 0) * 255)));
  const hex = n => n.toString(16).padStart(2, '0');
  return '#' + hex(to255(rgb.red)) + hex(to255(rgb.green)) + hex(to255(rgb.blue));
}

/* Map the Sheets-API border style enum to the short string names this
   tool uses internally (and that BORDER_STYLE_MAP keys off of). Unknown
   styles fall back to 'thin' on the apply side. */
const API_BORDER_STYLE_TO_NAME = {
  'SOLID':        'thin',
  'SOLID_MEDIUM': 'medium',
  'SOLID_THICK':  'thick',
  'DASHED':       'dashed',
  'DOTTED':       'dotted',
  'DOUBLE':       'double',
};

/* Read the borders for every cell in a range and return a map keyed by
   absolute "row,col" → { m, c, s } in the same shape paintArt writes.
   Uses the Sheets Advanced Service with a tight `fields` mask so the
   response stays small. Returns null when the Sheets service isn't
   loaded into the project — caller treats that as "no border info" and
   carries on without crashing.

   Per-cell collapse: paintArt-produced borders are uniform across the
   sides of a cell (single style + color), so we capture the mask of
   which sides are drawn and a single representative style/color taken
   from the first present side. Cells with mixed per-side styles
   degrade to the first-side's style — acceptable since this tool
   never produces those itself. */
function readBordersFromAPI_(sheet, startRow, startCol, rows, cols) {
  if (typeof Sheets === 'undefined') return null;

  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const a1 = "'" + sheet.getName().replace(/'/g, "''") + "'!"
    + sheet.getRange(startRow, startCol, rows, cols).getA1Notation();

  // userEnteredFormat.borders is what setBorder writes to. The fields
  // mask trims the response to just borders so nothing else (formulas,
  // text, formatting) gets shipped down.
  const resp = Sheets.Spreadsheets.get(ssId, {
    ranges: [a1],
    fields: 'sheets.data.rowData.values.userEnteredFormat.borders',
  });

  const result = {};
  const data = resp && resp.sheets && resp.sheets[0] && resp.sheets[0].data && resp.sheets[0].data[0];
  if (!data || !data.rowData) return result;

  for (let r = 0; r < data.rowData.length; r++) {
    const rowVals = (data.rowData[r] && data.rowData[r].values) || [];
    for (let c = 0; c < rowVals.length; c++) {
      const borders = rowVals[c] && rowVals[c].userEnteredFormat && rowVals[c].userEnteredFormat.borders;
      if (!borders) continue;

      // Side order matches the mask bit layout used throughout this
      // file: top=1, right=2, bottom=4, left=8. Newer responses use
      // colorStyle.rgbColor; older ones use color directly. We accept
      // either by checking both. NONE / missing / unset all read as
      // no border on that side.
      const sides = [
        { name: 'top',    bit: 1, data: borders.top },
        { name: 'right',  bit: 2, data: borders.right },
        { name: 'bottom', bit: 4, data: borders.bottom },
        { name: 'left',   bit: 8, data: borders.left },
      ];
      let mask = 0, color = null, styleName = null;
      for (const side of sides) {
        const b = side.data;
        if (!b || !b.style || b.style === 'NONE') continue;
        mask |= side.bit;
        if (color === null) {
          const rgb = (b.colorStyle && b.colorStyle.rgbColor) || b.color;
          color = apiColorToHex_(rgb);
          styleName = API_BORDER_STYLE_TO_NAME[b.style] || 'thin';
        }
      }
      if (mask) {
        result[(startRow + r) + ',' + (startCol + c)] = { m: mask, c: color, s: styleName };
      }
    }
  }
  return result;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sheet Art Maker')
    .addItem('Open Sheet Art Maker', 'showSidebar')
    .addSeparator()
    .addItem('Resize cells to 18×18', 'resizeCellsTo18')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Sheet Art Maker');
  SpreadsheetApp.getUi().showSidebar(html);
}

/* Returns {startRow, startCol, rows, cols} for the user's intended target.
   A multi-cell selection scopes the op to that range; a single cell or no
   selection scopes it to the whole sheet. Shared by reset / resize so they
   behave consistently — "if you've selected something, act on that;
   otherwise act on the canvas." */
function getTargetRange(sheet) {
  const range = sheet.getActiveRange();
  if (!range || (range.getNumRows() === 1 && range.getNumColumns() === 1)) {
    return {
      startRow: 1,
      startCol: 1,
      rows: sheet.getMaxRows(),
      cols: sheet.getMaxColumns(),
    };
  }
  return {
    startRow: range.getRow(),
    startCol: range.getColumn(),
    rows: range.getNumRows(),
    cols: range.getNumColumns(),
  };
}

/* Resizes columns and rows to a uniform pixel size. Parameterized so
   additional menu items (32×32 etc.) can hook in with one-line wrappers. */
function resizeCellsTo(sizePx) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const { startRow, startCol, rows, cols } = getTargetRange(sheet);
  sheet.setColumnWidths(startCol, cols, sizePx);
  sheet.setRowHeightsForced(startRow, rows, sizePx);
}

function resizeCellsTo18() {
  // Also expand the sheet to at least 100 columns. The 18×18 menu item
  // is shorthand for "prep this sheet for art" — tiny cells plus enough
  // horizontal room to land a reasonably wide image. Rows aren't touched
  // since they grow naturally as needed.
  const sheet = SpreadsheetApp.getActiveSheet();
  const current = sheet.getMaxColumns();
  if (current < 100) {
    sheet.insertColumnsAfter(current, 100 - current);
  }
  resizeCellsTo(18);
}

function fetchImageAsDataUrl(url) {
  const blob = UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getBlob();
  return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

function paintArt(colors, sparkles, fontColors, fontSizes, colWidths, rowHeights, startRow, startCol, borderLayers) {
  const sheet = SpreadsheetApp.getActiveSheet();

  // Resolve origin. The sidebar sends explicit startRow/startCol when
  // the user is in Coordinates mode, and null/undefined when in
  // Selection mode — in which case we anchor the painting to the
  // top-left of the user's current selection (any selection, even a
  // single cell). Falls back to A1 if there's no active range at all.
  if (startRow == null || startCol == null) {
    const range = sheet.getActiveRange();
    startRow = range ? range.getRow()    : 1;
    startCol = range ? range.getColumn() : 1;
  }

  const rows = colors.length;
  const cols = colors[0].length;
  const lastCol = startCol + cols - 1;
  const lastRow = startRow + rows - 1;

  if (sheet.getMaxColumns() < lastCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());
  if (sheet.getMaxRows() < lastRow)    sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());

  // Apply per-column widths and per-row heights if provided, otherwise fall
  // back to a uniform default. Sheets only supports uniform sizing along an
  // axis (every cell in a column shares its width), so we set one width per
  // column and one height per row.
  if (colWidths && colWidths.length === cols) {
    for (let i = 0; i < cols; i++) {
      sheet.setColumnWidth(startCol + i, colWidths[i]);
    }
  } else {
    sheet.setColumnWidths(startCol, cols, CELL_PX);
  }

  if (rowHeights && rowHeights.length === rows) {
    for (let i = 0; i < rows; i++) {
      sheet.setRowHeightsForced(startRow + i, 1, rowHeights[i]);
    }
  } else {
    sheet.setRowHeightsForced(startRow, rows, CELL_PX);
  }

  const range = sheet.getRange(startRow, startCol, rows, cols);

  // Generic null-merge helper: when any cell in `arr` is null, read
  // the corresponding current values from the sheet (via the
  // `read` thunk) and substitute them in. Returns the merged 2D
  // array, or the original when no null is found — so the
  // getter only fires on paints that actually use hide-on-blank.
  // Used for backgrounds, cell values, font colors, and font sizes.
  const mergeNulls = function(arr, read) {
    let any = false;
    for (let r = 0; r < rows && !any; r++) {
      for (let c = 0; c < cols; c++) {
        if (arr[r][c] === null) { any = true; break; }
      }
    }
    if (!any) return arr;
    const existing = read();
    return arr.map(function(row, r) {
      return row.map(function(cell, c) {
        return cell === null ? existing[r][c] : cell;
      });
    });
  };

  // Cells with null in `colors` mean "hide on blank for bg" — preserve
  // whatever color is currently there. We also reuse this same set of
  // null-cells later to figure out which cells should keep their
  // EXISTING borders (the bg-side hide toggle covers borders too).
  // Track here so we don't re-scan.
  const preservedBg = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (colors[r][c] === null) preservedBg.push([r, c]);
    }
  }
  range.setBackgrounds(mergeNulls(colors, function() { return range.getBackgrounds(); }));

  if (sparkles) {
    // Cell value, font color, and font size all share the same
    // text-side hide toggle on the sidebar, so all three arrays
    // carry null in the same cells. Merge each against its current
    // value before writing.
    range.setValues(mergeNulls(sparkles, function() { return range.getValues(); }));
    range.setHorizontalAlignment('center');
    range.setVerticalAlignment('middle');
  }
  if (fontColors) range.setFontColors(mergeNulls(fontColors, function() { return range.getFontColors(); }));
  if (fontSizes)  range.setFontSizes(mergeNulls(fontSizes,  function() { return range.getFontSizes();  }));

  // Borders: a list of layers. Each layer is { borders, colors,
  // style } where `borders` is a 2D mask array (bit 0=top, 1=right,
  // 2=bottom, 3=left), `colors` is a 2D per-cell hex string (or
  // null for cells with no border in this layer), and `style` is
  // the sidebar's style string mapped via BORDER_STYLE_MAP. Layers
  // paint in order; shared edges go last-write-wins so the
  // sidebar's layer order (Silhouette first, Contour second) is
  // how stacking is controlled.
  //
  // Preserve path: if the sidebar marked any cells with null in
  // `colors` (hide-on-blank for bg), read those cells' existing
  // borders BEFORE clearing so we can re-apply them after — the
  // bg-side hide toggle implicitly covers borders since they read
  // as part of the cell's background presentation. This runs even
  // when no border layers are active, so toggling all border
  // sections off still preserves borders on hidden-on-blank cells.
  const hasBorderLayers = borderLayers && borderLayers.length > 0;

  let savedBorders = null;
  if (preservedBg.length) {
    savedBorders = readBordersFromAPI_(sheet, startRow, startCol, rows, cols);
  }
  range.setBorder(false, false, false, false, false, false);
  if (savedBorders) {
    for (let i = 0; i < preservedBg.length; i++) {
      const r = preservedBg[i][0], c = preservedBg[i][1];
      const saved = savedBorders[(startRow + r) + ',' + (startCol + c)];
      if (!saved || !saved.m) continue;
      const cell = sheet.getRange(startRow + r, startCol + c);
      cell.setBorder(
        (saved.m & 1) ? true : null,
        (saved.m & 8) ? true : null,
        (saved.m & 4) ? true : null,
        (saved.m & 2) ? true : null,
        null, null,
        saved.c,
        BORDER_STYLE_MAP[saved.s] || SpreadsheetApp.BorderStyle.SOLID
      );
    }
  }

  // Apply each layer in order. Within a layer, bucket cells by
  // (mask, color) and write each bucket via one RangeList.setBorder
  // call — per-cell setBorder is unusable once the mask gets dense
  // (the Contour layer easily produces 1000+ border cells on a
  // typical image, and each setBorder is its own ~tens-of-ms
  // round-trip). Bucket count is bounded by 15 masks × distinct
  // colors per layer, typically a few dozen.
  if (hasBorderLayers) {
    for (let li = 0; li < borderLayers.length; li++) {
      const layer = borderLayers[li];
      if (!layer || !layer.borders || !layer.colors) continue;
      const style = BORDER_STYLE_MAP[layer.style] || SpreadsheetApp.BorderStyle.SOLID;
      const borderBuckets = {};
      for (let r = 0; r < rows; r++) {
        const rowMasks = layer.borders[r];
        const rowColors = layer.colors[r];
        if (!rowMasks || !rowColors) continue;
        for (let c = 0; c < cols; c++) {
          const mask = rowMasks[c];
          const color = rowColors[c];
          if (!mask || !color) continue;
          const key = mask + '|' + color;
          let bucket = borderBuckets[key];
          if (!bucket) {
            bucket = { mask: mask, color: color, a1s: [] };
            borderBuckets[key] = bucket;
          }
          bucket.a1s.push(colNumToLetter_(startCol + c) + (startRow + r));
        }
      }
      for (const key in borderBuckets) {
        const b = borderBuckets[key];
        sheet.getRangeList(b.a1s).setBorder(
          (b.mask & 1) ? true : null,  // top
          (b.mask & 8) ? true : null,  // left
          (b.mask & 4) ? true : null,  // bottom
          (b.mask & 2) ? true : null,  // right
          null, null,
          b.color,
          style
        );
      }
    }
  }

  // Select the painted region so the user can immediately see its extent
  // and act on it (copy, move, resize, etc.) without re-selecting by hand.
  // We activate the whole bounding rect rather than a RangeList of just
  // the painted cells — Sheets only supports copy on a non-contiguous
  // selection when all sub-ranges sit in a single row or column, so a
  // gappy multi-row RangeList silently collapses to one cell on copy.
  // The bounding rect copies cleanly, with hide-on-blank gaps coming
  // through as default white cells (which paste over as white).
  //
  // Note: the spreadsheet UI only redraws the new selection once focus
  // leaves the sidebar — the client-side success handler calls
  // google.script.host.editor.focus() to trigger that.
  range.activate();

  return { rows, cols };
}

/* Clears a range back to: white background, no content, no borders,
   unmerged, default 18x18 cell size. */
function clearRangeToDefault(sheet, startRow, startCol, rows, cols) {
  const range = sheet.getRange(startRow, startCol, rows, cols);

  // Unmerge before other ops — merged cells block per-cell writes and
  // would error on setBackgrounds with a 2D color array of the wrong shape.
  range.breakApart();

  range.setBackground('#FFFFFF');
  range.clearContent();
  range.clearFormat();
  range.setBorder(false, false, false, false, false, false);

  sheet.setColumnWidths(startCol, cols, CELL_PX);
  sheet.setRowHeightsForced(startRow, rows, CELL_PX);
}

/* Resets cells back to a clean default state. Selection behavior matches
   the other utilities via getTargetRange. Returns { rows, cols } so the
   sidebar can confirm what was cleared. */
function resetPaintedAreas() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const { startRow, startCol, rows, cols } = getTargetRange(sheet);
  clearRangeToDefault(sheet, startRow, startCol, rows, cols);
  return { rows, cols };
}

/* ----------------------------------------------------------------------
   Shape-preserving copy / paste

   Sheets' built-in copy only works on a single contiguous rectangle (or
   a RangeList that's entirely within one row or one column) — so once
   you've painted art with hide-on-blank gaps, you can't just Ctrl+C the
   shape and paste it cleanly elsewhere. The clipboard collapses to one
   cell, or the gaps come along as default white that wipes out whatever
   was at the destination.

   These two functions sidestep Sheets' clipboard entirely. copySelection
   reads the active range and snapshots every "painted" cell (bg, value,
   font color/size, alignments, borders) into a JS object the sidebar
   holds in memory. pasteShape takes that snapshot, reads the
   destination's current cell state, overlays only the painted source
   cells on top, and writes the result back — non-painted source
   positions are passed through from the destination, so existing pixels
   there survive.

   Borders ride along via the Sheets Advanced Service (see
   readBordersFromAPI_ at the top of this file) — the Apps Script
   Range API can't read borders, but the underlying REST API can.
   This means manual borders set via the Sheets toolbar also carry,
   not just ones this tool wrote.

   Limitation: Sheets reports both never-painted and explicitly-painted
   white cells as #ffffff, so cells painted with a white palette swatch
   register as blank and won't transfer (a cell with only a border but
   no bg or text does still come along — the border lookup catches
   it). ---------------------------------------------------------------- */

/* Snapshot the painted cells inside the currently active range. Returns
   { rows, cols, cells } where rows/cols describe the source bounding box
   and cells is the flat list of painted positions with their formatting,
   each relative to the bbox top-left. Returns null if there's no active
   range or nothing in it looks painted. */
function copySelection() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  if (!range) return null;

  const rows = range.getNumRows();
  const cols = range.getNumColumns();
  const startRow = range.getRow();
  const startCol = range.getColumn();

  // Pull every formatting layer we care about in one batched read per
  // attribute. Cheaper than touching each cell individually, and the
  // arrays stay aligned by [row][col] indexing.
  const bgs        = range.getBackgrounds();
  const vals       = range.getValues();
  const fontColors = range.getFontColors();
  const fontSizes  = range.getFontSizes();
  const halign     = range.getHorizontalAlignments();
  const valign     = range.getVerticalAlignments();

  // Borders come from the Sheets Advanced Service — null when it's not
  // enabled in the project, in which case copy carries everything except
  // borders rather than crashing.
  const borderMap = readBordersFromAPI_(sheet, startRow, startCol, rows, cols) || {};

  const cells = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const bg = String(bgs[r][c] || '').toLowerCase();
      const val = vals[r][c];
      // A cell counts as painted if its bg isn't the default white, it
      // has any text content, or it has a tracked border. Case-insensitive
      // bg compare since Sheets' returned hex casing isn't guaranteed.
      const hasBg = bg !== '#ffffff';
      const hasVal = val !== '' && val != null;
      const border = borderMap[(startRow + r) + ',' + (startCol + c)];
      if (!hasBg && !hasVal && !border) continue;
      const cell = {
        r, c,
        bg: bgs[r][c],
        val: hasVal ? val : null,
        fontColor: fontColors[r][c],
        fontSize: fontSizes[r][c],
        halign: halign[r][c],
        valign: valign[r][c],
      };
      if (border) cell.border = border;
      cells.push(cell);
    }
  }

  if (!cells.length) return null;
  return { rows, cols, cells };
}

/* Write a clipboard snapshot at the top-left of the current selection.
   Reads the destination bounding box first so non-painted source
   positions keep whatever was already there — that's what makes the
   paste shape-aware instead of a rectangular stamp. Returns { startRow,
   startCol, rows, cols, painted } so the sidebar can confirm and
   re-activate the destination's bounding rect (rectangle, not RangeList,
   so subsequent native copy still works). */
function pasteShape(clipboard) {
  if (!clipboard || !clipboard.cells || !clipboard.cells.length) return null;
  const sheet = SpreadsheetApp.getActiveSheet();
  const sel = sheet.getActiveRange();
  if (!sel) return null;

  const startRow = sel.getRow();
  const startCol = sel.getColumn();
  const { rows, cols, cells } = clipboard;

  // Expand the sheet if the paste would land past its current bounds.
  // Matches paintArt's pad-on-write behavior so a paste into the bottom
  // edge doesn't error out.
  const lastRow = startRow + rows - 1;
  const lastCol = startCol + cols - 1;
  if (sheet.getMaxRows() < lastRow)    sheet.insertRowsAfter(sheet.getMaxRows(), lastRow - sheet.getMaxRows());
  if (sheet.getMaxColumns() < lastCol) sheet.insertColumnsAfter(sheet.getMaxColumns(), lastCol - sheet.getMaxColumns());

  const dest = sheet.getRange(startRow, startCol, rows, cols);

  // Read destination's current state. Each grid is the "before" picture
  // we'll mutate where painted source cells land, then write back whole.
  // Cells the source skips stay as the destination originally had them.
  const bgs        = dest.getBackgrounds();
  const vals       = dest.getValues();
  const fontColors = dest.getFontColors();
  const fontSizes  = dest.getFontSizes();
  const halign     = dest.getHorizontalAlignments();
  const valign     = dest.getVerticalAlignments();

  for (const cell of cells) {
    bgs[cell.r][cell.c]        = cell.bg;
    vals[cell.r][cell.c]       = cell.val == null ? '' : cell.val;
    fontColors[cell.r][cell.c] = cell.fontColor;
    fontSizes[cell.r][cell.c]  = cell.fontSize;
    halign[cell.r][cell.c]     = cell.halign;
    valign[cell.r][cell.c]     = cell.valign;
  }

  dest.setBackgrounds(bgs);
  dest.setValues(vals);
  dest.setFontColors(fontColors);
  dest.setFontSizes(fontSizes);
  dest.setHorizontalAlignments(halign);
  dest.setVerticalAlignments(valign);

  // Borders. Clear borders on every painted destination cell first, so a
  // painted cell with no border doesn't keep a stale border the
  // destination had — borders overwrite per painted cell, matching how
  // bg/value/font above replace their cells. Non-painted gap cells are left
  // untouched, so the destination's existing art (and its borders) survives
  // wherever the pasted shape has no cell. The sheet stays the source of
  // truth — a later copySelection reads these back via the Advanced Service.
  // One batched RangeList clear (the pattern paintArt uses), then apply.
  const paintedA1s = cells.map(cell =>
    colNumToLetter_(startCol + cell.c) + (startRow + cell.r)
  );
  if (paintedA1s.length) {
    sheet.getRangeList(paintedA1s).setBorder(false, false, false, false, false, false);
  }
  for (const cell of cells) {
    if (!cell.border) continue;
    const { m, c: color, s } = cell.border;
    const style = BORDER_STYLE_MAP[s] || SpreadsheetApp.BorderStyle.SOLID;
    sheet.getRange(startRow + cell.r, startCol + cell.c).setBorder(
      (m & 1) ? true : null,  // top
      (m & 8) ? true : null,  // left
      (m & 4) ? true : null,  // bottom
      (m & 2) ? true : null,  // right
      null, null,
      color,
      style
    );
  }

  // Activate the bounding rect so native Ctrl+C on the paste result
  // works normally (and so the user can see what just landed). We don't
  // use a RangeList of just the painted cells — Sheets won't copy a
  // gappy multi-row selection. The user can re-run copySelection on
  // this rect to pick up the shape again for another paste.
  dest.activate();

  return { startRow, startCol, rows, cols, painted: cells.length };
}

