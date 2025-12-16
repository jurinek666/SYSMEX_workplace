// file: office-scripts/transform-sap-simple.ts
/**
 * Minimalistický, efektivní Office Script pro SAP export.
 * Pozn.: Office Scripts neumí "Save As"; změny se ukládají do otevřeného sešitu.
 */

function main(workbook: ExcelScript.Workbook) {
  const log = ensureLogSheet(workbook);
  logInfo(log, "Start");

  try {
    const ws = autodetectDataWorksheet(workbook, log);
    const table = ensureWholeUsedRangeAsTable(ws, log);

    applyTransformations(ws, table, log);             // filtrování, sloupce, přejmenování, sort, delete row 2
    shrinkToFirstFourColumnsAndEnsureTblSAP(ws, log); // ořez na A:D + tbl_SAP
    ensureWorksheetNamedSAP(workbook, ws, log);       // přejmenování listu na "SAP"

    autofitAllColumns(ws);
    logInfo(log, "Done");
  } catch (e) {
    logError(log, "Chyba", (e as Error)?.message ?? String(e));
    throw e;
  }
}

/* ===================== Transformace ===================== */
function applyTransformations(
  ws: ExcelScript.Worksheet,
  table: ExcelScript.Table,
  log: ExcelScript.Worksheet
) {
  assertHeadersExist(table, ["Storage location"], "Chybí sloupec pro filtraci", log);
  assertHeadersExist(table, ["Material", "Material description", "Batch", "Total Quantity"], "Chybí povinné výstupní sloupce", log);

  // ponechat jen F010, F070 (ostatní řádky smazat)
  const storageHeader = resolveHeaderName(table, "Storage location");
  const allowed = new Set(["F010", "F070"]);
  deleteRowsWhere(table, row => !allowed.has(String(row[storageHeader] ?? "").trim()), log);

  // smazat sloupce Plant & Storage location
  deleteColumn(table, "Plant", log);
  deleteColumn(table, "Storage location", log);

  // seřadit sloupce dopředu (pořadí výstupu)
  reorderColumnsToFront(table, ["Material", "Material description", "Batch", "Total Quantity"], log);

  // přejmenovat sloupce
  renameColumn(table, "Material description", "Název", log);
  renameColumn(table, "Total Quantity", "Mnozstvi_SAP", log);
  assertHeadersExist(table, ["Material", "Název", "Batch", "Mnozstvi_SAP"], "Chybí přejmenované sloupce", log);

  // setřídit podle Mnozstvi_SAP sestupně
  const qtyIdx = requireHeader(table, "Mnozstvi_SAP");
  table.getSort().apply([{ key: qtyIdx, ascending: false }]);
  logInfo(log, "sort", "Mnozstvi_SAP DESC");

  // smazat řádek 2 (první datový pod hlavičkou)
  ws.getRange("2:2").delete(ExcelScript.DeleteShiftDirection.up);
  logInfo(log, "deleteRow", "Smazán řádek 2");
}

/* ===== Ořez na A:D a tabulka tbl_SAP ===== */
function shrinkToFirstFourColumnsAndEnsureTblSAP(ws: ExcelScript.Worksheet, log: ExcelScript.Worksheet) {
  const used = ws.getUsedRange();
  if (used) {
    const totalCols = used.getColumnCount();
    if (totalCols > 4) {
      ws.getRangeByIndexes(0, 4, used.getRowCount(), totalCols - 4)
        .delete(ExcelScript.DeleteShiftDirection.left);
      logInfo(log, "shrinkColumns", `Odstraněny sloupce od E`);
    }
  }

  const firstFour = ws.getRange("A:D").getUsedRange();
  if (!firstFour) throw new Error("V A:D nejsou data pro tabulku 'tbl_SAP'.");

  const lastRow = firstFour.getRowCount(); // začíná v řádku 1
  const target = ws.getRange(`A1:D${lastRow}`);

  // vytvořit/resize tabulku na A1:D{lastRow} a pojmenovat tbl_SAP
  const tables = ws.getTables();
  let tbl = tables.find(t => equalsIgnoreCase(t.getName(), "tbl_SAP"));
  if (!tbl && tables.length > 0) {
    tbl = tables[0];
    try { tbl.setName("tbl_SAP"); } catch {/* ignore collision */}
  }

  if (tbl) {
    tbl.resize(target);
    try { tbl.setName("tbl_SAP"); } catch {/* ignore collision */}
    logInfo(log, "ensureTblSAP", `Resize 'tbl_SAP' na ${target.getAddress()}`);
  } else {
    const newTbl = ws.addTable(target.getAddress(), true);
    try { newTbl.setName("tbl_SAP"); } catch {/* ignore collision */}
    logInfo(log, "ensureTblSAP", `Vytvořena 'tbl_SAP' na ${target.getAddress()}`);
  }
}

/* ===== Přejmenování listu ===== */
function ensureWorksheetNamedSAP(workbook: ExcelScript.Workbook, target: ExcelScript.Worksheet, log: ExcelScript.Worksheet) {
  const existing = workbook.getWorksheet("SAP");
  if (existing && existing.getId() !== target.getId()) {
    existing.delete(); // uvolnit název
    logWarn(log, "renameSheet", "Smazán existující list 'SAP'");
  }
  target.setName("SAP");
  logInfo(log, "renameSheet", "List přejmenován na 'SAP'");
}

/* ===================== Autodetekce + tabulka ===================== */
function autodetectDataWorksheet(workbook: ExcelScript.Workbook, log: ExcelScript.Worksheet): ExcelScript.Worksheet {
  const sheets = workbook.getWorksheets().filter(ws => ws.getName() !== "_log_transform");
  let best: ExcelScript.Worksheet | null = null, bestScore = -1;

  for (const ws of sheets) {
    const used = ws.getUsedRange();
    if (!used) continue;
    const usedValues = used.getValues();
    const headerRow = usedValues[0] ?? [];
    const nonEmptyHeaders = headerRow.filter(v => String(v ?? "").trim() !== "").length;
    if (nonEmptyHeaders === 0) continue;
    const score = used.getRowCount() * used.getColumnCount();
    if (score > bestScore) { bestScore = score; best = ws; }
  }

  if (!best) throw new Error("Nenašel jsem list s hlavičkou v řádku 1");
  logInfo(log, "autodetectDataWorksheet", `Vybrán list: ${best.getName()}`);
  return best;
}

function ensureWholeUsedRangeAsTable(ws: ExcelScript.Worksheet, log: ExcelScript.Worksheet): ExcelScript.Table {
  const existing = ws.getTables();
  if (existing.length > 0) {
    logInfo(log, "ensureTable", `Použita tabulka: ${existing[0].getName()}`);
    return existing[0];
  }
  const used = ws.getUsedRange();
  if (!used) throw new Error(`List '${ws.getName()}' neobsahuje data`);
  const tbl = ws.addTable(used.getAddress(), true);
  logInfo(log, "ensureTable", `Vytvořena tabulka z ${used.getAddress()}`);
  return tbl;
}

/* ===================== Log (jednoduchý) ===================== */
function ensureLogSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const name = "_log_transform";
  const ws = workbook.getWorksheet(name) ?? workbook.addWorksheet(name);
  if ((ws.getUsedRange()?.getRowCount() ?? 0) === 0) {
    ws.getRange("A1:C1").setValues([["Timestamp", "Action", "Detail"]]);
    ws.getRange("A1:C1").getFormat().getFont().setBold(true);
  }
  return ws;
}
function appendLog(log: ExcelScript.Worksheet, level: "INFO"|"WARN"|"ERROR", action: string, detail: string) {
  const used = log.getUsedRange();
  const next = (used ? used.getRowCount() : 0) + 1;
  log.getRange(`A${next}:C${next}`).setValues([[new Date().toISOString(), `${level}:${action}`, detail]]);
}
function logInfo(log: ExcelScript.Worksheet, action: string, detail = "") { appendLog(log, "INFO", action, detail); }
function logWarn(log: ExcelScript.Worksheet, action: string, detail = "") { appendLog(log, "WARN", action, detail); }
function logError(log: ExcelScript.Worksheet, action: string, detail = "") { appendLog(log, "ERROR", action, detail); }

/* ===================== Helpery ===================== */
function getHeaders(table: ExcelScript.Table): string[] {
  return table.getHeaderRowRange().getValues()[0].map(h => String(h ?? "").trim());
}
function equalsIgnoreCase(a: string, b: string) {
  return a.localeCompare(b, undefined, { sensitivity: "accent" }) === 0;
}
function requireHeader(table: ExcelScript.Table, header: string) {
  const headers = getHeaders(table);
  const idx = headers.findIndex(h => equalsIgnoreCase(h, header));
  if (idx < 0) throw new Error(`Sloupec '${header}' nenalezen`);
  return idx;
}
function resolveHeaderName(table: ExcelScript.Table, header: string): string {
  const headers = getHeaders(table);
  const i = headers.findIndex(h => equalsIgnoreCase(h, header));
  if (i < 0) throw new Error(`Sloupec '${header}' nenalezen`);
  return headers[i];
}
function assertHeadersExist(table: ExcelScript.Table, required: string[], errorPrefix: string, log: ExcelScript.Worksheet) {
  const headers = getHeaders(table);
  const missing = required.filter(req => !headers.some(h => equalsIgnoreCase(h, req)));
  if (missing.length > 0) { const msg = `${errorPrefix}: ${missing.join(", ")}`; logError(log, "assertHeadersExist", msg); throw new Error(msg); }
}
function deleteColumn(table: ExcelScript.Table, header: string, log: ExcelScript.Worksheet) {
  const headers = getHeaders(table);
  const idx = headers.findIndex(h => equalsIgnoreCase(h, header));
  if (idx < 0) { logWarn(log, "deleteColumn", `Nenalezen '${header}'`); return; }
  const name = table.getColumns()[idx].getName();
  table.getColumns()[idx].delete();
  logInfo(log, "deleteColumn", `Smazán '${name}'`);
}
function deleteRowsWhere(table: ExcelScript.Table, predicate: (row: Record<string, unknown>) => boolean, log: ExcelScript.Worksheet) {
  const headers = getHeaders(table);
  const body = table.getRangeBetweenHeaderAndTotal();
  const rows = body.getValues();
  let deleted = 0;
  for (let i = rows.length - 1; i >= 0; i--) {
    const obj: Record<string, unknown> = {};
    headers.forEach((h, c) => (obj[h] = rows[i][c]));
    if (predicate(obj)) {
      // use the previously-read body range to delete the row (avoid calling table.getDataBodyRange())
      body.getRows()[i].delete(ExcelScript.DeleteShiftDirection.up);
      deleted++;
    }
  }
  logInfo(log, "deleteRowsWhere", `Smazáno: ${deleted}`);
}
function renameColumn(table: ExcelScript.Table, oldName: string, newName: string, log: ExcelScript.Worksheet) {
  const headers = getHeaders(table);
  const idx = headers.findIndex(h => equalsIgnoreCase(h, oldName));
  if (idx < 0) { logWarn(log, "renameColumn", `Nenalezen '${oldName}'`); return; }
  table.getColumns()[idx].setName(newName);
  logInfo(log, "renameColumn", `'${oldName}' -> '${newName}'`);
}
function reorderColumnsToFront(table: ExcelScript.Table, desiredOrder: string[], log: ExcelScript.Worksheet) {
  let targetIndex = 0;
  for (const desired of desiredOrder) {
    const headers = getHeaders(table);
    const actual = headers.find(h => equalsIgnoreCase(h, desired));
    if (!actual) { logWarn(log, "reorderColumnsToFront", `Chybí '${desired}'`); continue; }
    moveColumnToIndex(table, actual, targetIndex, log);
    targetIndex++;
  }
  logInfo(log, "reorderColumnsToFront", desiredOrder.join(" | "));
}
function moveColumnToIndex(table: ExcelScript.Table, header: string, targetIndex: number, log: ExcelScript.Worksheet) {
  const headers = getHeaders(table);
  const currentIdx = headers.findIndex(h => equalsIgnoreCase(h, header));
  if (currentIdx < 0 || currentIdx === targetIndex) return;

  const srcCol = table.getColumns()[currentIdx];
  const srcName = srcCol.getName();
  const values = srcCol.getRangeBetweenHeaderAndTotal().getValues();

  const added = table.addColumn(targetIndex);
  added.setName(srcName);
  added.getRangeBetweenHeaderAndTotal().setValues(values);

  const cols = table.getColumns();
  for (let i = 0; i < cols.length; i++) {
    if (equalsIgnoreCase(cols[i].getName(), srcName) && i !== targetIndex) {
      table.getColumns()[i].delete();
      logInfo(log, "moveColumnToIndex", `'${srcName}' -> index ${targetIndex}`);
      break;
    }
  }
}
function autofitAllColumns(ws: ExcelScript.Worksheet) {
  const used = ws.getUsedRange();
  if (used) used.getFormat().autofitColumns();
}
