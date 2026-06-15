const DECLARATION_FIELDS = [
  { key: "boxNo", label: "标记号码", numeric: false, width: 24 },
  { key: "goodsName", label: "货物名称及规格", numeric: false, width: 12 },
  { key: "totalBoxes", label: "总箱数", numeric: true, width: 10 },
  { key: "quantity", label: "总数量", numeric: true, width: 8 },
  { key: "grossWeight", label: "总毛重", numeric: true, width: 8 },
  { key: "netWeight", label: "净重", numeric: true, width: 10 },
];
const DEFAULT_DECLARATION_LAYOUT = {
  boxNo: 1,
  goodsName: 3,
  totalBoxes: 1,
  quantity: 4,
  grossWeight: 2,
  netWeight: 1,
};
const LAYOUT_INPUT_IDS = {
  boxNo: "layoutBoxNo",
  goodsName: "layoutGoodsName",
  totalBoxes: "layoutTotalBoxes",
  quantity: "layoutQuantity",
  grossWeight: "layoutGrossWeight",
  netWeight: "layoutNetWeight",
};
const ARCHIVE_KEY = "packing-list-archive-v1";
const CONFIG_KEY = "packing-list-config-v1";

let archive = loadJson(ARCHIVE_KEY, { items: {}, updatedAt: "" });
let config = loadJson(CONFIG_KEY, {
  netMode: "fixed",
  netValue: 1,
  mergeBoxCells: false,
  mergeMixedNames: false,
  hideVisuals: false,
  declarationLayout: { ...DEFAULT_DECLARATION_LAYOUT },
  mixedNames: {},
  customsGroups: [],
});
let archiveHandle = null;
let shipments = [];

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  fileList: document.querySelector("#fileList"),
  results: document.querySelector("#results"),
  alerts: document.querySelector("#alerts"),
  summaryText: document.querySelector("#summaryText"),
  editNamesBtn: document.querySelector("#editNamesBtn"),
  downloadBtn: document.querySelector("#downloadBtn"),
  editDialog: document.querySelector("#editDialog"),
  editRows: document.querySelector("#editRows"),
  applyEditBtn: document.querySelector("#applyEditBtn"),
  netMode: document.querySelector("#netMode"),
  netValue: document.querySelector("#netValue"),
  pickArchiveBtn: document.querySelector("#pickArchiveBtn"),
  saveArchiveBtn: document.querySelector("#saveArchiveBtn"),
  importArchiveBtn: document.querySelector("#importArchiveBtn"),
  exportArchiveBtn: document.querySelector("#exportArchiveBtn"),
  archiveImport: document.querySelector("#archiveImport"),
  archiveExcelImport: document.querySelector("#archiveExcelImport"),
  mergeBoxCells: document.querySelector("#mergeBoxCells"),
  mergeMixedNames: document.querySelector("#mergeMixedNames"),
  hideVisuals: document.querySelector("#hideVisuals"),
  layoutInputs: {},
  editMixedNamesBtn: document.querySelector("#editMixedNamesBtn"),
  mixedNameDialog: document.querySelector("#mixedNameDialog"),
  mixedNameRows: document.querySelector("#mixedNameRows"),
  applyMixedNamesBtn: document.querySelector("#applyMixedNamesBtn"),
  editCustomsGroupsBtn: document.querySelector("#editCustomsGroupsBtn"),
  customsGroupDialog: document.querySelector("#customsGroupDialog"),
  customsShipmentChoices: document.querySelector("#customsShipmentChoices"),
  customsDialogGroups: document.querySelector("#customsDialogGroups"),
  customsGroupList: document.querySelector("#customsGroupList"),
  addCustomsGroupBtn: document.querySelector("#addCustomsGroupBtn"),
  importArchiveExcelBtn: document.querySelector("#importArchiveExcelBtn"),
  exportArchiveExcelBtn: document.querySelector("#exportArchiveExcelBtn"),
  archiveStatus: document.querySelector("#archiveStatus"),
};

init();

function init() {
  DECLARATION_FIELDS.forEach((field) => {
    els.layoutInputs[field.key] = document.querySelector(`#${LAYOUT_INPUT_IDS[field.key]}`);
  });
  config.declarationLayout = normalizeDeclarationLayout(config.declarationLayout);
  els.netMode.value = config.netMode || "fixed";
  els.netValue.value = Number(config.netValue ?? 1);
  els.mergeBoxCells.checked = Boolean(config.mergeBoxCells);
  els.mergeMixedNames.checked = Boolean(config.mergeMixedNames);
  els.hideVisuals.checked = Boolean(config.hideVisuals);
  DECLARATION_FIELDS.forEach((field) => {
    els.layoutInputs[field.key].value = getDeclarationLayout()[field.key];
  });
  updateArchiveStatus();
  renderCustomsGroupList();

  els.fileInput.addEventListener("change", () => handleFiles([...els.fileInput.files]));
  ["dragenter", "dragover"].forEach((name) => {
    els.dropzone.addEventListener(name, (event) => {
      event.preventDefault();
      els.dropzone.classList.add("dragging");
    });
  });
  ["dragleave", "drop"].forEach((name) => {
    els.dropzone.addEventListener(name, (event) => {
      event.preventDefault();
      els.dropzone.classList.remove("dragging");
    });
  });
  els.dropzone.addEventListener("drop", (event) => handleFiles([...event.dataTransfer.files]));

  els.netMode.addEventListener("change", saveConfigAndRebuild);
  els.netValue.addEventListener("input", saveConfigAndRebuild);
  els.mergeBoxCells.addEventListener("change", saveConfigAndRebuild);
  els.mergeMixedNames.addEventListener("change", saveConfigAndRebuild);
  els.hideVisuals.addEventListener("change", saveConfigAndRebuild);
  DECLARATION_FIELDS.forEach((field) => {
    els.layoutInputs[field.key].addEventListener("input", saveConfigAndRebuild);
  });
  els.editMixedNamesBtn.addEventListener("click", openMixedNameDialog);
  els.applyMixedNamesBtn.addEventListener("click", applyMixedNameDialog);
  els.editCustomsGroupsBtn.addEventListener("click", openCustomsGroupDialog);
  els.addCustomsGroupBtn.addEventListener("click", addCustomsGroupFromDialog);
  els.editNamesBtn.addEventListener("click", openEditDialog);
  els.applyEditBtn.addEventListener("click", applyEditDialog);
  els.downloadBtn.addEventListener("click", downloadWorkbook);
  els.pickArchiveBtn.addEventListener("click", pickArchiveFile);
  els.saveArchiveBtn.addEventListener("click", saveArchive);
  els.importArchiveBtn.addEventListener("click", () => els.archiveImport.click());
  els.archiveImport.addEventListener("change", importArchive);
  els.exportArchiveBtn.addEventListener("click", exportArchive);
  els.importArchiveExcelBtn.addEventListener("click", () => els.archiveExcelImport.click());
  els.archiveExcelImport.addEventListener("change", importArchiveExcel);
  els.exportArchiveExcelBtn.addEventListener("click", exportArchiveExcel);

  if (!window.XLSX) {
    setAlerts(["Excel 解析库没有加载成功。请确认浏览器能访问 cdn.jsdelivr.net，或者用本地依赖版本运行。"]);
  }
  refreshIcons();
}

async function handleFiles(files) {
  if (!window.XLSX) {
    setAlerts(["Excel 解析库没有加载成功，暂时无法读取上传文件。"]);
    return;
  }

  const xlsxFiles = files.filter((file) => /\.(xlsx|xls)$/i.test(file.name));
  if (!xlsxFiles.length) return;

  renderFileList(xlsxFiles);
  setAlerts([]);
  shipments = [];
  const alerts = [];

  for (const file of xlsxFiles) {
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
      const parsed = parseShipment(rows, file.name);
      shipments.push(parsed);
    } catch (error) {
      alerts.push(`${file.name}: ${error.message}`);
    }
  }

  rebuildOutputs();
  if (!shipments.length && !alerts.length) {
    alerts.push("没有解析到可用的 FBA 货件，请确认上传的是 FBA 装箱单 Excel。");
  }
  setAlerts(alerts);
}

function parseShipment(rows, fileName) {
  const shipmentId = findValueAfter(rows, "货件单号") || fileName.replace(/\.[^.]+$/, "");
  const mode = findValueAfter(rows, "装箱方式");
  const headerRowIndex = rows.findIndex((row) => row.some((cell) => normalize(cell) === "MSKU"));
  if (headerRowIndex < 0) throw new Error("未找到 MSKU 表头");

  const headers = rows[headerRowIndex].map(normalize);
  const idx = mapHeaders(headers);
  if (idx.msku == null) throw new Error("未找到 MSKU 列");

  const isMixed = String(mode).includes("多款") || headers.some((h) => /^第\d+箱$/.test(h));
  return isMixed
    ? parseMixedShipment(rows, headerRowIndex, idx, shipmentId, mode, fileName)
    : parseSingleSkuShipment(rows, headerRowIndex, idx, shipmentId, mode, fileName);
}

function parseSingleSkuShipment(rows, headerRowIndex, idx, shipmentId, mode, fileName) {
  const sourceRows = [];
  for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const msku = cell(row, idx.msku);
    if (!msku) continue;
    if (normalize(msku).includes("合计")) break;
    const boxCodes = expandBoxRange(cell(row, idx.boxNo));
    const perBoxQty = toNumber(cell(row, idx.singleBoxQty)) || 1;
    const boxWeight = toNumber(cell(row, idx.boxWeight)) || 0;
    sourceRows.push({
      msku,
      productName: cell(row, idx.productName),
      sku: cell(row, idx.sku),
      quantity: perBoxQty,
      boxWeight,
      boxCodes,
    });

    if (boxWeight && perBoxQty) {
      upsertArchive(mskuKey(msku), { unitGross: round(boxWeight / perBoxQty, 4) }, false);
    }
  }

  const rawItems = [];
  for (const row of sourceRows) {
    for (const boxNo of row.boxCodes) {
      rawItems.push({
        boxNo,
        msku: row.msku,
        productName: row.productName,
        sku: row.sku,
        quantity: row.quantity,
        grossWeight: row.boxWeight,
        grossSource: "input",
      });
    }
  }

  return finalizeShipment({ shipmentId, mode, fileName, rawItems });
}

function parseMixedShipment(rows, headerRowIndex, idx, shipmentId, mode, fileName) {
  const headers = rows[headerRowIndex].map(normalize);
  const boxColumns = headers
    .map((header, index) => ({ header, index }))
    .filter((item) => /^第\d+箱$/.test(item.header));

  if (!boxColumns.length) throw new Error("未找到第 N 箱列");

  const metaRows = findMixedMetaRows(rows);
  const sourceRows = [];
  for (let r = headerRowIndex + 1; r < rows.length; r += 1) {
    const row = rows[r];
    const msku = cell(row, idx.msku);
    if (!msku) continue;
    if (normalize(row).includes("合计")) break;
    sourceRows.push({
      row,
      msku,
      productName: cell(row, idx.productName),
      sku: cell(row, idx.sku),
    });
  }

  const unitWeights = inferUnitWeights(sourceRows, boxColumns, metaRows);
  const rawItems = [];

  for (const box of boxColumns) {
    const boxNo = cell(rows[metaRows.boxNo], box.index) || box.header;
    const boxWeight = toNumber(cell(rows[metaRows.weight], box.index));
    const entries = sourceRows
      .map((item) => ({ ...item, quantity: toNumber(cell(item.row, box.index)) }))
      .filter((item) => item.quantity > 0);
    const totalQty = entries.reduce((sum, item) => sum + item.quantity, 0);

    for (const item of entries) {
      const key = mskuKey(item.msku);
      const archivedWeight = Number(archive.items[key]?.unitGross || 0);
      const inferredWeight = unitWeights.get(key) || 0;
      let grossWeight = 0;
      let grossSource = "calculated";

      if (entries.length === 1) {
        grossWeight = boxWeight;
        grossSource = "input";
      } else if (archivedWeight) {
        grossWeight = archivedWeight * item.quantity;
        grossSource = "archive";
      } else if (inferredWeight) {
        grossWeight = inferredWeight * item.quantity;
        grossSource = "inferred";
      } else if (boxWeight && totalQty) {
        grossWeight = boxWeight * item.quantity / totalQty;
      }

      rawItems.push({
        boxNo,
        msku: item.msku,
        productName: item.productName,
        sku: item.sku,
        quantity: item.quantity,
        grossWeight,
        grossSource,
      });
    }
  }

  return finalizeShipment({ shipmentId, mode, fileName, rawItems });
}

function inferUnitWeights(sourceRows, boxColumns, metaRows) {
  const inferred = new Map();
  for (const box of boxColumns) {
    const weightRow = metaRows.rows[metaRows.weight] || [];
    const boxWeight = toNumber(cell(weightRow, box.index));
    if (!boxWeight) continue;
    const entries = sourceRows
      .map((item) => ({ key: mskuKey(item.msku), qty: toNumber(cell(item.row, box.index)) }))
      .filter((item) => item.qty > 0);
    if (entries.length === 1 && entries[0].qty) {
      inferred.set(entries[0].key, round(boxWeight / entries[0].qty, 4));
      upsertArchive(entries[0].key, { unitGross: round(boxWeight / entries[0].qty, 4) }, false);
    }
  }
  return inferred;
}

function finalizeShipment({ shipmentId, mode, fileName, rawItems }, mergeOutput = true) {
  const grouped = new Map();
  for (const item of rawItems) {
    const archived = archive.items[mskuKey(item.msku)] || {};
    const goodsName = archived.goodsName || item.productName || item.sku || item.msku;
    const key = `${item.boxNo}\u0000${goodsName}`;
    const existing = grouped.get(key) || {
      boxNo: item.boxNo,
      goodsName,
      totalBoxes: 1,
      quantity: 0,
      grossWeight: 0,
      grossSource: item.grossSource,
      mskus: new Set(),
      productNames: new Set(),
    };
    existing.quantity += item.quantity;
    existing.grossWeight += item.grossWeight;
    if (item.grossSource !== "input") existing.grossSource = item.grossSource;
    existing.mskus.add(item.msku);
    if (item.productName) existing.productNames.add(item.productName);
    grouped.set(key, existing);
  }

  const baseRows = [...grouped.values()]
    .sort((a, b) => compareBoxNo(a.boxNo, b.boxNo) || a.goodsName.localeCompare(b.goodsName, "zh-CN"))
    .map((row) => ({
      ...row,
      grossWeight: round(row.grossWeight, 2),
      netWeight: round(calcNetWeight(row.grossWeight), 2),
      mskus: [...row.mskus],
      productNames: [...row.productNames],
    }));

  const outputRows = mergeOutput && config.mergeMixedNames ? mergeMixedOutputRows(shipmentId, baseRows, rawItems) : baseRows;
  return { shipmentId, mode, fileName, rawItems, outputRows };
}

function findMixedMetaRows(rows) {
  const meta = { rows };
  rows.forEach((row, index) => {
    const text = row.map(normalize).join("|");
    if (text.includes("Weight of box")) meta.weight = index;
    if (text.includes("箱号")) meta.boxNo = index;
  });
  if (meta.weight == null || meta.boxNo == null) throw new Error("未找到箱重或箱号行");
  return meta;
}

function mapHeaders(headers) {
  return {
    msku: headers.indexOf("MSKU"),
    productName: headers.indexOf("品名"),
    sku: headers.indexOf("SKU"),
    singleBoxQty: headers.indexOf("单箱数量"),
    boxWeight: headers.findIndex((h) => h.includes("箱子毛重")),
    boxNo: headers.indexOf("箱号"),
  };
}

function expandBoxRange(value) {
  const text = normalize(value).replace(/[；;]/g, "");
  if (!text) return [];
  const rangeMatch = text.match(/^(.*?)(\d+)[~～](\d+)$/);
  if (!rangeMatch) return [text];

  const prefix = rangeMatch[1];
  const startText = rangeMatch[2];
  const end = Number(rangeMatch[3]);
  const start = Number(startText);
  const width = startText.length;
  const result = [];
  for (let n = start; n <= end; n += 1) {
    result.push(`${prefix}${String(n).padStart(width, "0")}`);
  }
  return result;
}

function rebuildOutputs() {
  const existingRaw = shipments.map((shipment) => ({
    shipmentId: shipment.shipmentId,
    mode: shipment.mode,
    fileName: shipment.fileName,
    rawItems: shipment.rawItems,
  }));
  shipments = existingRaw.map(finalizeShipment);
  pruneCustomsGroups();
  renderResults();
  renderCustomsGroupList();
  updateControls();
}

function renderResults() {
  els.results.innerHTML = "";
  els.results.classList.toggle("empty-state", shipments.length === 0);

  if (!shipments.length) {
    els.results.innerHTML = `<i data-lucide="boxes"></i><p>上传 FBA 装箱单后生成结果</p>`;
    els.summaryText.textContent = "等待上传文件";
    refreshIcons();
    return;
  }

  const displayShipments = getDisplayShipments();
  const totalRows = displayShipments.reduce((sum, item) => sum + item.outputRows.length, 0);
  els.summaryText.textContent = `已解析 ${shipments.length} 个货件，${displayShipments.length} 组输出，${totalRows} 行`;

  const template = document.querySelector("#shipmentTemplate");
  for (const shipment of displayShipments) {
    const node = template.content.cloneNode(true);
    const details = node.querySelector(".shipment");
    node.querySelector(".shipment-title").textContent = shipment.displayName || shipment.shipmentId;
    node.querySelector(".shipment-meta").textContent = `${shipment.mode || "未知装箱方式"} · ${shipment.outputRows.length} 行 · ${getMixedModeText()}`;
    const copyButton = node.querySelector(".copy-btn");
    copyButton.addEventListener("click", (event) => {
      event.preventDefault();
      copyShipment(shipment);
    });
    const visualWrap = node.querySelector(".visual-wrap");
    if (config.hideVisuals) {
      visualWrap.remove();
    } else {
      visualWrap.appendChild(buildVisual(shipment));
    }
    node.querySelector(".table-wrap").appendChild(buildTable(shipment.outputRows));
    details.dataset.shipmentId = shipment.shipmentId;
    els.results.appendChild(node);
  }
  refreshIcons();
}

function getMixedModeText() {
  const modes = [];
  if (config.mergeBoxCells) modes.push("合并单元格");
  if (config.mergeMixedNames) modes.push("合并名称");
  return modes.length ? modes.join(" / ") : "未启用混装处理";
}

function buildTable(rows) {
  const table = document.createElement("table");
  const thead = table.createTHead();
  const headRow = thead.insertRow();
  const layout = getDeclarationLayout();
  DECLARATION_FIELDS.forEach((field) => {
    appendHeaderCell(headRow, field.label, layout[field.key], field.numeric);
  });

  const tbody = table.createTBody();
  rows.forEach((row, rowIndex) => {
    const tr = tbody.insertRow();
    DECLARATION_FIELDS.forEach((field) => {
      if (isVerticalMergeContinuation(rows, rowIndex, field)) return;
      appendCell(tr, row[field.key], field, {
        colSpan: layout[field.key],
        rowSpan: getVerticalMergeDown(rows, rowIndex, field) + 1,
        warn: field.key === "grossWeight" && row.grossSource !== "input",
      });
    });
  });
  return table;
}

function appendHeaderCell(tr, value, colSpan = 1, numeric = false) {
  const th = document.createElement("th");
  th.textContent = value;
  if (colSpan > 1) th.colSpan = colSpan;
  if (numeric) th.className = "num";
  th.style.cssText = getDeclarationCellStyle("header");
  tr.appendChild(th);
}

function appendCell(tr, value, field, options = {}) {
  const td = tr.insertCell();
  td.textContent = value;
  if (field.numeric) td.className = "num";
  if (options.colSpan > 1) td.colSpan = options.colSpan;
  if (options.rowSpan > 1) td.rowSpan = options.rowSpan;
  td.style.cssText = getDeclarationCellStyle(field.key);
  if (options.warn) {
    td.classList.add("warn");
    td.title = "毛重为推算值或档案值";
  }
}

function buildVisual(shipment) {
  const wrap = document.createElement("div");
  wrap.className = "box-visual";
  const sections = shipment.customsMembers?.length
    ? shipment.customsMembers.map((member, index) => ({
      title: member.shipmentId,
      colorIndex: index,
      boxes: groupRawItemsByBox(member.rawItems),
    }))
    : [{ title: "", colorIndex: 0, boxes: groupRawItemsByBox(shipment.rawItems) }];

  sections.forEach((section) => {
    const sectionEl = document.createElement("div");
    sectionEl.className = `visual-section visual-tone-${section.colorIndex % 6}`;
    if (section.title) {
      const sectionTitle = document.createElement("div");
      sectionTitle.className = "visual-section-title";
      sectionTitle.textContent = section.title;
      sectionEl.appendChild(sectionTitle);
    }

    const grid = document.createElement("div");
    grid.className = "visual-grid";

    for (const box of section.boxes) {
      const card = document.createElement("div");
      card.className = "visual-box";
      const title = document.createElement("div");
      title.className = "visual-box-title";
      title.textContent = box.boxNo;
      card.appendChild(title);

      const stack = document.createElement("div");
      stack.className = "visual-stack";
      box.items.slice(0, 6).forEach((item) => {
        const chip = document.createElement("div");
        chip.className = "visual-item";
        chip.textContent = `${item.name} × ${item.quantity}`;
        stack.appendChild(chip);
      });
      if (box.items.length > 6) {
        const more = document.createElement("div");
        more.className = "visual-more";
        more.textContent = `+${box.items.length - 6} 项`;
        stack.appendChild(more);
      }
      card.appendChild(stack);
      grid.appendChild(card);
    }

    sectionEl.appendChild(grid);
    wrap.appendChild(sectionEl);
  });

  return wrap;
}

function mergeMixedOutputRows(shipmentId, rows, rawItems = []) {
  const byBox = new Map();
  rows.forEach((row) => {
    if (!byBox.has(row.boxNo)) byBox.set(row.boxNo, []);
    byBox.get(row.boxNo).push(row);
  });
  const rawByBox = new Map();
  rawItems.forEach((item) => {
    if (!rawByBox.has(item.boxNo)) rawByBox.set(item.boxNo, []);
    rawByBox.get(item.boxNo).push(item);
  });

  const merged = [];
  for (const [boxNo, group] of byBox.entries()) {
    const rawGroup = rawByBox.get(boxNo) || [];
    const rawNames = rawGroup
      .map((item) => archive.items[mskuKey(item.msku)]?.goodsName || item.productName || item.sku || item.msku)
      .filter(Boolean);
    const uniqueNames = [...new Set(rawNames)];
    const shouldMerge = group.length > 1 || uniqueNames.length > 1;

    if (!shouldMerge) {
      merged.push(group[0]);
      continue;
    }

    const names = uniqueNames.length ? uniqueNames : [...new Set(group.map((row) => row.goodsName).filter(Boolean))];
    const key = mixedNameKey(shipmentId, boxNo);
    const defaultName = defaultMixedName(names);
    merged.push({
      boxNo,
      goodsName: config.mixedNames?.[key] || defaultName,
      totalBoxes: 1,
      quantity: round(group.reduce((sum, row) => sum + Number(row.quantity || 0), 0), 2),
      grossWeight: round(group.reduce((sum, row) => sum + Number(row.grossWeight || 0), 0), 2),
      netWeight: round(group.reduce((sum, row) => sum + Number(row.netWeight || 0), 0), 2),
      grossSource: group.some((row) => row.grossSource !== "input") ? "calculated" : "input",
      mskus: [...new Set(group.flatMap((row) => row.mskus || []))],
      productNames: [...new Set(group.flatMap((row) => row.productNames || []))],
    });
  }
  return merged.sort((a, b) => compareBoxNo(a.boxNo, b.boxNo) || a.goodsName.localeCompare(b.goodsName, "zh-CN"));
}

function getMixedNameGroups() {
  const groups = [];
  shipments.forEach((shipment) => {
    const byBox = new Map();
    shipment.rawItems.forEach((item) => {
      if (!byBox.has(item.boxNo)) byBox.set(item.boxNo, []);
      byBox.get(item.boxNo).push(item);
    });
    byBox.forEach((items, boxNo) => {
      const names = [...new Set(items
        .map((item) => archive.items[mskuKey(item.msku)]?.goodsName || item.productName || item.sku || item.msku)
        .filter(Boolean))];
      if (names.length < 2) return;
      const key = mixedNameKey(shipment.shipmentId, boxNo);
      groups.push({
        key,
        shipmentId: shipment.shipmentId,
        boxNo,
        names,
        defaultName: defaultMixedName(names),
        savedName: config.mixedNames?.[key] || "",
      });
    });
  });
  return groups.sort((a, b) => a.shipmentId.localeCompare(b.shipmentId) || compareBoxNo(a.boxNo, b.boxNo));
}

function groupRawItemsByBox(rawItems) {
  const byBox = new Map();
  rawItems.forEach((item) => {
    if (!byBox.has(item.boxNo)) byBox.set(item.boxNo, new Map());
    const goodsName = archive.items[mskuKey(item.msku)]?.goodsName || item.productName || item.sku || item.msku;
    const productMap = byBox.get(item.boxNo);
    const current = productMap.get(goodsName) || { name: goodsName, quantity: 0 };
    current.quantity += Number(item.quantity || 0);
    productMap.set(goodsName, current);
  });
  return [...byBox.entries()]
    .sort(([a], [b]) => compareBoxNo(a, b))
    .map(([boxNo, productMap]) => ({
      boxNo,
      items: [...productMap.values()].sort((a, b) => a.name.localeCompare(b.name, "zh-CN")),
    }));
}

function getExcelMerges(rows) {
  const layout = getLayoutRanges();
  const merges = [];
  layout.forEach((field) => {
    if (field.span > 1) {
      merges.push({ s: { r: 0, c: field.start }, e: { r: 0, c: field.end } });
    }
  });
  rows.forEach((row, index) => {
    layout.forEach((field) => {
      if (field.span > 1) {
        merges.push({ s: { r: index + 1, c: field.start }, e: { r: index + 1, c: field.end } });
      }
    });
  });
  if (!config.mergeBoxCells || config.mergeMixedNames) return merges;
  let index = 0;
  while (index < rows.length) {
    let end = index + 1;
    while (end < rows.length && rows[end].boxNo === rows[index].boxNo) end += 1;
    if (end - index > 1) {
      const boxField = layout.find((field) => field.key === "boxNo");
      const totalBoxesField = layout.find((field) => field.key === "totalBoxes");
      if (boxField.span === 1) {
        merges.push({ s: { r: index + 1, c: boxField.start }, e: { r: end, c: boxField.start } });
      }
      if (totalBoxesField.span === 1) {
        merges.push({ s: { r: index + 1, c: totalBoxesField.start }, e: { r: end, c: totalBoxesField.start } });
      }
    }
    index = end;
  }
  return merges;
}

function getDisplayShipments() {
  const byId = new Map(shipments.map((shipment) => [shipment.shipmentId, shipment]));
  const used = new Set();
  const display = [];

  normalizeCustomsGroups(config.customsGroups).forEach((group) => {
    const members = group.map((id) => byId.get(id)).filter(Boolean);
    if (members.length < 2) return;
    members.forEach((member) => used.add(member.shipmentId));
    display.push(buildCustomsShipment(members));
  });

  shipments.forEach((shipment) => {
    if (!used.has(shipment.shipmentId)) display.push(shipment);
  });

  return display;
}

function buildCustomsShipment(members) {
  const ids = members.map((shipment) => shipment.shipmentId);
  return {
    shipmentId: ids.join("&"),
    displayName: getCustomsGroupName(ids),
    mode: "合并报关",
    fileName: ids.join("&"),
    rawItems: members.flatMap((shipment) => shipment.rawItems),
    outputRows: members.flatMap((shipment) => shipment.outputRows),
    customsMembers: members,
  };
}

function getCustomsGroupName(group) {
  return `${group.join("&")}的合并报关`;
}

function normalizeCustomsGroups(groups) {
  if (!Array.isArray(groups)) return [];
  const seen = new Set();
  return groups
    .map((group) => Array.isArray(group) ? [...new Set(group.map(normalize).filter(Boolean))] : [])
    .filter((group) => group.length >= 2)
    .filter((group) => {
      const key = [...group].sort().join("|");
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function pruneCustomsGroups() {
  const ids = new Set(shipments.map((shipment) => shipment.shipmentId));
  config.customsGroups = normalizeCustomsGroups(config.customsGroups)
    .map((group) => group.filter((id) => ids.has(id)))
    .filter((group) => group.length >= 2);
  persistConfig();
}

function mixedNameKey(shipmentId, boxNo) {
  return `${shipmentId}::${boxNo}`;
}

function defaultMixedName(names) {
  return `${names.join(" / ")} 混装`;
}

function openEditDialog() {
  const seen = new Map();
  for (const shipment of shipments) {
    for (const item of shipment.rawItems) {
      const key = mskuKey(item.msku);
      if (!seen.has(key)) {
        const saved = archive.items[key] || {};
        seen.set(key, {
          msku: item.msku,
          savedName: saved.goodsName || "",
          placeholder: item.productName || item.sku || item.msku,
          unitGross: saved.unitGross || "",
        });
      }
    }
  }

  els.editRows.innerHTML = "";
  for (const row of seen.values()) {
    const div = document.createElement("div");
    div.className = "edit-row";
    div.dataset.msku = row.msku;
    div.dataset.defaultName = row.placeholder;
    div.innerHTML = `
      <div>
        <strong>${escapeHtml(row.msku)}</strong>
        <div class="hint">${escapeHtml(row.placeholder)}</div>
      </div>
      <input class="goods-input" placeholder="${escapeHtml(row.placeholder)}" value="${escapeHtml(row.savedName)}">
      <input class="weight-input" type="number" min="0" step="0.0001" placeholder="单件毛重" value="${escapeHtml(row.unitGross)}">
      <div class="hint">留空名称时使用原表品名，毛重用于混装箱推算</div>
    `;
    els.editRows.appendChild(div);
  }
  els.editDialog.showModal();
}

function applyEditDialog() {
  syncArchiveFromEditRows();
  rebuildOutputs();
}

function syncArchiveFromEditRows() {
  for (const row of els.editRows.querySelectorAll(".edit-row")) {
    const key = mskuKey(row.dataset.msku);
    if (!key) continue;
    const goodsName = row.querySelector(".goods-input").value.trim();
    const unitGross = toNumber(row.querySelector(".weight-input").value);
    upsertArchive(key, {
      goodsName,
      unitGross: unitGross || "",
    }, true);
  }
}

function openMixedNameDialog() {
  if (!shipments.length) {
    setAlerts(["请先上传并成功解析至少一个 FBA 装箱单。"]);
    return;
  }
  if (!config.mergeMixedNames) {
    config.mergeMixedNames = true;
    els.mergeMixedNames.checked = true;
    persistConfig();
    rebuildOutputs();
  }

  const groups = getMixedNameGroups();
  els.mixedNameRows.innerHTML = "";

  if (!groups.length) {
    const empty = document.createElement("div");
    empty.className = "edit-row single-col";
    empty.textContent = "当前没有需要合并名称的混装箱。";
    els.mixedNameRows.appendChild(empty);
  }

  groups.forEach((group) => {
    const div = document.createElement("div");
    div.className = "edit-row mixed-name-row";
    div.dataset.key = group.key;
    div.innerHTML = `
      <div>
        <strong>${escapeHtml(group.shipmentId)}</strong>
        <div class="hint">${escapeHtml(group.boxNo)}</div>
      </div>
      <input class="mixed-name-input" placeholder="${escapeHtml(group.defaultName)}" value="${escapeHtml(group.savedName)}">
      <div class="hint">${escapeHtml(group.names.join(" / "))}</div>
    `;
    els.mixedNameRows.appendChild(div);
  });

  els.mixedNameDialog.showModal();
}

function applyMixedNameDialog() {
  const mixedNames = { ...(config.mixedNames || {}) };
  for (const row of els.mixedNameRows.querySelectorAll(".mixed-name-row")) {
    mixedNames[row.dataset.key] = row.querySelector(".mixed-name-input").value.trim();
  }
  config.mixedNames = mixedNames;
  persistConfig();
  rebuildOutputs();
}

function openCustomsGroupDialog() {
  if (shipments.length < 2) {
    setAlerts([`当前只成功解析到 ${shipments.length} 个货件，合并报关至少需要 2 个。`]);
    return;
  }
  renderCustomsGroupDialog();
  els.customsGroupDialog.showModal();
}

function renderCustomsGroupDialog() {
  els.customsShipmentChoices.innerHTML = "";
  shipments.forEach((shipment) => {
    const label = document.createElement("label");
    label.className = "choice-item";
    label.innerHTML = `
      <input type="checkbox" value="${escapeHtml(shipment.shipmentId)}">
      <span>${escapeHtml(shipment.shipmentId)}</span>
    `;
    els.customsShipmentChoices.appendChild(label);
  });
  renderCustomsDialogGroups();
}

function addCustomsGroupFromDialog() {
  const selected = [...els.customsShipmentChoices.querySelectorAll("input:checked")]
    .map((input) => input.value);
  const unique = [...new Set(selected)];
  if (unique.length < 2) {
    setAlerts(["合并报关至少需要选择两个 FBA 货件。"]);
    return;
  }

  const existing = normalizeCustomsGroups(config.customsGroups)
    .filter((group) => !group.some((id) => unique.includes(id)));
  existing.push(unique);
  config.customsGroups = existing;
  persistConfig();
  renderCustomsGroupDialog();
  rebuildOutputs();
}

function removeCustomsGroup(index) {
  const groups = normalizeCustomsGroups(config.customsGroups);
  groups.splice(index, 1);
  config.customsGroups = groups;
  persistConfig();
  renderCustomsGroupDialog();
  rebuildOutputs();
}

function renderCustomsDialogGroups() {
  renderCustomsGroupsInto(els.customsDialogGroups, true);
}

function renderCustomsGroupList() {
  renderCustomsGroupsInto(els.customsGroupList, false);
}

function renderCustomsGroupsInto(container, removable) {
  container.innerHTML = "";
  const groups = normalizeCustomsGroups(config.customsGroups);
  if (!groups.length) {
    const empty = document.createElement("p");
    empty.className = "muted small-note no-indent";
    empty.textContent = shipments.length ? "暂无合并报关组" : "上传后可设置合并报关";
    container.appendChild(empty);
    return;
  }

  groups.forEach((group, index) => {
    const row = document.createElement("div");
    row.className = "group-pill";
    const name = getCustomsGroupName(group);
    row.innerHTML = `<span>${escapeHtml(name)}</span>`;
    if (removable) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "icon-btn";
      button.title = "删除";
      button.innerHTML = `<i data-lucide="trash-2"></i>`;
      button.addEventListener("click", () => removeCustomsGroup(index));
      row.appendChild(button);
    }
    container.appendChild(row);
  });
  refreshIcons();
}

async function copyShipment(shipment) {
  const text = toDeclarationAoa(shipment.outputRows, { includeHeader: false, blankMergedContinuations: false }).map((row) => row.join("\t")).join("\n");
  const html = buildClipboardHtml(shipment.outputRows);
  if (window.ClipboardItem) {
    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([text], { type: "text/plain" }),
      }),
    ]);
    return;
  }
  await navigator.clipboard.writeText(text);
}

function downloadWorkbook() {
  const workbookXml = buildSpreadsheetXml(getDisplayShipments());
  const blob = new Blob([workbookXml], { type: "application/vnd.ms-excel;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `装箱单-${formatDate(new Date())}.xls`;
  a.click();
  URL.revokeObjectURL(url);
  persistArchive();
}

function toDeclarationAoa(rows, options = {}) {
  const data = [];
  if (options.includeHeader) {
    data.push(buildDeclarationHeaderRow());
  }
  rows.forEach((row, index) => {
    const continuation = options.blankMergedContinuations
      && config.mergeBoxCells
      && !config.mergeMixedNames
      && index > 0
      && rows[index - 1].boxNo === row.boxNo;
    data.push(buildDeclarationDataRow(row, continuation));
  });
  return data;
}

function buildClipboardHtml(rows) {
  const htmlRows = buildStyledRows(rows, { includeHeader: false, blankMergedContinuations: false });
  return `<table style="border-collapse:collapse;">${htmlRows}</table>`;
}

function buildSpreadsheetXml(displayShipments) {
  const worksheets = displayShipments.map((shipment) => {
    const sheetName = xmlEscape(safeSheetName(shipment.displayName || shipment.shipmentId));
    return `
      <Worksheet ss:Name="${sheetName}">
        <Table>
          ${buildSpreadsheetColumns()}
          ${buildSpreadsheetHeaderRow()}
          ${buildSpreadsheetDataRows(shipment.outputRows)}
        </Table>
        <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
          <DisplayGridlines/>
        </WorksheetOptions>
      </Worksheet>
    `;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8"?>
    <?mso-application progid="Excel.Sheet"?>
    <Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
      xmlns:html="http://www.w3.org/TR/REC-html40">
      <Styles>
        <Style ss:ID="Header">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders>${spreadsheetBorders()}</Borders>
          <Font ss:FontName="宋体" ss:Size="9" ss:Color="#000000"/>
        </Style>
        <Style ss:ID="BoxNo">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders>${spreadsheetBorders()}</Borders>
          <Font ss:FontName="Arial" ss:Size="10.5" ss:Color="#000000"/>
        </Style>
        <Style ss:ID="RedText">
          <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
          <Borders>${spreadsheetBorders()}</Borders>
          <Font ss:FontName="宋体" ss:Size="9" ss:Color="#FF0000"/>
        </Style>
        <Style ss:ID="RedTextLeft">
          <Alignment ss:Horizontal="Left" ss:Vertical="Center"/>
          <Borders>${spreadsheetBorders()}</Borders>
          <Font ss:FontName="宋体" ss:Size="9" ss:Color="#FF0000"/>
        </Style>
      </Styles>
      ${worksheets}
    </Workbook>`;
}

function spreadsheetBorders() {
  return `
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>
  `;
}

function buildSpreadsheetColumns() {
  return getLayoutColumns().map((column) => `<Column ss:Width="${column.wch * 7}"/>`).join("");
}

function buildSpreadsheetHeaderRow() {
  const cells = getLayoutRanges().map((field) => buildSpreadsheetCell(field.label, field, {
    styleId: "Header",
    type: "String",
  })).join("");
  return `<Row>${cells}</Row>`;
}

function buildSpreadsheetDataRows(rows) {
  return rows.map((row, rowIndex) => {
    const cells = getLayoutRanges().map((field) => {
      if (isVerticalMergeContinuation(rows, rowIndex, field)) return "";
      let value = row[field.key];
      const styleId = field.key === "boxNo" ? "BoxNo" : (field.key === "goodsName" ? "RedTextLeft" : "RedText");
      return buildSpreadsheetCell(value, field, {
        styleId,
        type: field.numeric ? "Number" : "String",
        mergeDown: getVerticalMergeDown(rows, rowIndex, field),
      });
    }).join("");
    return `<Row>${cells}</Row>`;
  }).join("");
}

function buildSpreadsheetCell(value, field, options = {}) {
  const mergeAcross = field.span > 1 ? ` ss:MergeAcross="${field.span - 1}"` : "";
  const mergeDown = options.mergeDown ? ` ss:MergeDown="${options.mergeDown}"` : "";
  const type = options.type || "String";
  const text = value == null || value === "" ? "" : String(value);
  return `<Cell ss:Index="${field.start + 1}" ss:StyleID="${options.styleId}"${mergeAcross}${mergeDown}><Data ss:Type="${type}">${xmlEscape(text)}</Data></Cell>`;
}

function getVerticalMergeDown(rows, rowIndex, field) {
  if (!config.mergeBoxCells || config.mergeMixedNames) return 0;
  if (field.key !== "boxNo" && field.key !== "totalBoxes") return 0;
  if (rowIndex > 0 && rows[rowIndex - 1].boxNo === rows[rowIndex].boxNo) return 0;

  let end = rowIndex + 1;
  while (end < rows.length && rows[end].boxNo === rows[rowIndex].boxNo) end += 1;
  return Math.max(0, end - rowIndex - 1);
}

function isVerticalMergeContinuation(rows, rowIndex, field) {
  if (!config.mergeBoxCells || config.mergeMixedNames) return false;
  if (field.key !== "boxNo" && field.key !== "totalBoxes") return false;
  return rowIndex > 0 && rows[rowIndex - 1].boxNo === rows[rowIndex].boxNo;
}

function buildStyledRows(rows, options = {}) {
  const output = [];
  if (options.includeHeader) {
    output.push(`<tr>${DECLARATION_FIELDS.map((field) => buildStyledCell(field.label, field, { header: true })).join("")}</tr>`);
  }
  rows.forEach((row, index) => {
    const cells = DECLARATION_FIELDS.map((field) => {
      if (isVerticalMergeContinuation(rows, index, field)) return "";
      let value = row[field.key];
      return buildStyledCell(value, field, {
        header: false,
        rowSpan: getVerticalMergeDown(rows, index, field) + 1,
      });
    }).join("");
    output.push(`<tr>${cells}</tr>`);
  });
  return output.join("");
}

function buildStyledCell(value, field, options = {}) {
  const span = getDeclarationLayout()[field.key];
  const colspan = span > 1 ? ` colspan="${span}"` : "";
  const rowspan = options.rowSpan > 1 ? ` rowspan="${options.rowSpan}"` : "";
  return `<td${colspan}${rowspan} style="${getDeclarationCellStyle(options.header ? "header" : field.key)}">${escapeHtml(value ?? "")}</td>`;
}

function getDeclarationCellStyle(fieldKey) {
  const isHeader = fieldKey === "header";
  const isBoxNo = fieldKey === "boxNo";
  const fontFamily = isBoxNo ? "Arial" : "SimSun, 宋体";
  const color = isHeader || isBoxNo ? "#000000" : "#FF0000";
  const fontSize = isBoxNo ? "10.5pt" : "9pt";
  const align = fieldKey === "boxNo" || fieldKey === "goodsName" ? "left" : "center";
  return [
    "border:1px solid #000000",
    "font-weight:normal",
    `font-family:${fontFamily}`,
    `font-size:${fontSize}`,
    `color:${color}`,
    `text-align:${align}`,
    "vertical-align:middle",
    "mso-number-format:'\\@'",
    "padding:0 4px",
  ].join(";");
}

function buildDeclarationHeaderRow() {
  const row = [];
  getLayoutRanges().forEach((field) => {
    row[field.start] = field.label;
    for (let c = field.start + 1; c <= field.end; c += 1) row[c] = "";
  });
  return row;
}

function buildDeclarationDataRow(sourceRow, continuation) {
  const row = [];
  getLayoutRanges().forEach((field) => {
    let value = sourceRow[field.key];
    if (continuation && (field.key === "boxNo" || field.key === "totalBoxes")) value = "";
    row[field.start] = value;
    for (let c = field.start + 1; c <= field.end; c += 1) row[c] = "";
  });
  return row;
}

function getLayoutColumns() {
  return getLayoutRanges().flatMap((field) => {
    const columns = [{ wch: field.width }];
    for (let i = 1; i < field.span; i += 1) columns.push({ wch: field.width });
    return columns;
  });
}

function getLayoutRanges() {
  const layout = getDeclarationLayout();
  let start = 0;
  return DECLARATION_FIELDS.map((field) => {
    const span = layout[field.key];
    const range = { ...field, span, start, end: start + span - 1 };
    start += span;
    return range;
  });
}

function getDeclarationLayout() {
  config.declarationLayout = normalizeDeclarationLayout(config.declarationLayout);
  return config.declarationLayout;
}

function normalizeDeclarationLayout(layout) {
  const result = { ...DEFAULT_DECLARATION_LAYOUT };
  DECLARATION_FIELDS.forEach((field) => {
    const value = Number(layout?.[field.key]);
    result[field.key] = Number.isFinite(value) ? Math.max(1, Math.min(12, Math.round(value))) : DEFAULT_DECLARATION_LAYOUT[field.key];
  });
  return result;
}

function exportArchiveExcel() {
  if (!window.XLSX) {
    setAlerts(["Excel 解析库没有加载成功，暂时无法导出档案 Excel。"]);
    return;
  }

  syncArchiveFromEditRows();
  const rows = [["MSKU", "默认品名", "货物名称及规格", "单件毛重", "更新时间"]];
  const exported = new Set();
  for (const editRow of els.editRows.querySelectorAll(".edit-row")) {
    const msku = mskuKey(editRow.dataset.msku);
    if (!msku) continue;
    exported.add(msku);
    const item = archive.items[msku] || {};
    rows.push([msku, editRow.dataset.defaultName || "", item.goodsName || "", item.unitGross || "", item.updatedAt || ""]);
  }
  Object.entries(archive.items || {})
    .filter(([msku]) => !exported.has(msku))
    .sort(([a], [b]) => a.localeCompare(b))
    .forEach(([msku, item]) => {
      rows.push([msku, "", item.goodsName || "", item.unitGross || "", item.updatedAt || ""]);
    });
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  sheet["!cols"] = [{ wch: 28 }, { wch: 36 }, { wch: 36 }, { wch: 12 }, { wch: 24 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "MSKU档案");
  XLSX.writeFile(workbook, `MSKU档案-${formatDate(new Date())}.xlsx`);
}

async function importArchiveExcel() {
  if (!window.XLSX) {
    setAlerts(["Excel 解析库没有加载成功，暂时无法导入档案 Excel。"]);
    return;
  }

  const file = els.archiveExcelImport.files[0];
  if (!file) return;
  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    const header = (rows[0] || []).map(normalize);
    const mskuIndex = header.indexOf("MSKU");
    const nameIndex = header.indexOf("货物名称及规格");
    const weightIndex = header.indexOf("单件毛重");
    if (mskuIndex < 0) throw new Error("第一行必须包含 MSKU 列");

    for (let i = 1; i < rows.length; i += 1) {
      const row = rows[i];
      const key = mskuKey(cell(row, mskuIndex));
      if (!key) continue;
      upsertArchive(key, {
        goodsName: nameIndex >= 0 ? cell(row, nameIndex) : "",
        unitGross: weightIndex >= 0 ? toNumber(cell(row, weightIndex)) || "" : "",
      }, true);
    }
    persistArchive();
    rebuildOutputs();
    setArchiveStatus(`已导入档案 Excel: ${file.name}`);
  } catch (error) {
    setArchiveStatus(`导入失败: ${error.message}`);
  } finally {
    els.archiveExcelImport.value = "";
  }
}

async function pickArchiveFile() {
  if (!window.showOpenFilePicker) {
    setArchiveStatus("当前浏览器不支持直接选择档案文件，已使用本地存储。");
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{ description: "JSON", accept: { "application/json": [".json"] } }],
      multiple: false,
    });
    archiveHandle = handle;
    const file = await handle.getFile();
    archive = JSON.parse(await file.text());
    persistArchive(false);
    rebuildOutputs();
    setArchiveStatus(`已连接档案: ${file.name}`);
  } catch (error) {
    if (error.name !== "AbortError") setArchiveStatus(error.message);
  }
}

async function saveArchive() {
  persistArchive();
  if (!archiveHandle || !window.showSaveFilePicker) {
    exportArchive();
    return;
  }
  try {
    const writable = await archiveHandle.createWritable();
    await writable.write(JSON.stringify(archive, null, 2));
    await writable.close();
    setArchiveStatus("档案已保存");
  } catch (error) {
    setArchiveStatus(error.message);
  }
}

async function importArchive() {
  const file = els.archiveImport.files[0];
  if (!file) return;
  try {
    archive = JSON.parse(await file.text());
    persistArchive(false);
    rebuildOutputs();
    setArchiveStatus(`已导入: ${file.name}`);
  } catch (error) {
    setArchiveStatus(`导入失败: ${error.message}`);
  } finally {
    els.archiveImport.value = "";
  }
}

function exportArchive() {
  persistArchive();
  const blob = new Blob([JSON.stringify(archive, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "packing-list-archive.json";
  a.click();
  URL.revokeObjectURL(url);
}

function saveConfigAndRebuild() {
  const declarationLayout = {};
  DECLARATION_FIELDS.forEach((field) => {
    declarationLayout[field.key] = toNumber(els.layoutInputs[field.key].value) || DEFAULT_DECLARATION_LAYOUT[field.key];
  });
  config = {
    ...config,
    netMode: els.netMode.value,
    netValue: toNumber(els.netValue.value),
    mergeBoxCells: els.mergeBoxCells.checked,
    mergeMixedNames: els.mergeMixedNames.checked,
    hideVisuals: els.hideVisuals.checked,
    declarationLayout: normalizeDeclarationLayout(declarationLayout),
  };
  persistConfig();
  DECLARATION_FIELDS.forEach((field) => {
    els.layoutInputs[field.key].value = config.declarationLayout[field.key];
  });
  rebuildOutputs();
}

function persistConfig() {
  config.customsGroups = normalizeCustomsGroups(config.customsGroups);
  localStorage.setItem(CONFIG_KEY, JSON.stringify(config));
}

function calcNetWeight(grossWeight) {
  const value = Number(config.netValue || 0);
  if (config.netMode === "percent") return Math.max(0, grossWeight * (1 - value / 100));
  return Math.max(0, grossWeight - value);
}

function upsertArchive(key, patch, overwrite) {
  const current = archive.items[key] || {};
  const next = { ...current };
  for (const [name, value] of Object.entries(patch)) {
    if (overwrite || current[name] == null || current[name] === "") next[name] = value;
  }
  next.updatedAt = new Date().toISOString();
  archive.items[key] = next;
}

function persistArchive(updateTime = true) {
  if (updateTime) archive.updatedAt = new Date().toISOString();
  localStorage.setItem(ARCHIVE_KEY, JSON.stringify(archive));
  updateArchiveStatus();
}

function loadJson(key, fallback) {
  try {
    return JSON.parse(localStorage.getItem(key)) || fallback;
  } catch {
    return fallback;
  }
}

function renderFileList(files) {
  els.fileList.innerHTML = "";
  files.forEach((file) => {
    const div = document.createElement("div");
    div.className = "file-pill";
    div.innerHTML = `<i data-lucide="file-spreadsheet"></i><span>${escapeHtml(file.name)}</span>`;
    els.fileList.appendChild(div);
  });
  refreshIcons();
}

function setAlerts(alerts) {
  els.alerts.innerHTML = "";
  alerts.forEach((message) => {
    const div = document.createElement("div");
    div.className = "alert";
    div.textContent = message;
    els.alerts.appendChild(div);
  });
}

function updateControls() {
  const enabled = shipments.length > 0;
  els.editNamesBtn.disabled = !enabled;
  els.downloadBtn.disabled = !enabled;
  els.editMixedNamesBtn.disabled = !enabled;
  els.editCustomsGroupsBtn.disabled = !enabled;
}

function updateArchiveStatus() {
  const count = Object.keys(archive.items || {}).length;
  setArchiveStatus(`已存 ${count} 个 MSKU`);
}

function setArchiveStatus(text) {
  els.archiveStatus.textContent = text;
}

function refreshIcons() {
  if (window.lucide?.createIcons) {
    window.lucide.createIcons();
  }
}

function findValueAfter(rows, label) {
  for (const row of rows) {
    for (let i = 0; i < row.length; i += 1) {
      if (normalize(row[i]) === label) return cell(row, i + 1);
    }
  }
  return "";
}

function normalize(value) {
  return String(value ?? "").trim();
}

function cell(row, index) {
  if (index == null || index < 0) return "";
  return normalize(row?.[index]);
}

function toNumber(value) {
  const number = Number(String(value ?? "").replace(/,/g, "").trim());
  return Number.isFinite(number) ? number : 0;
}

function round(value, digits = 2) {
  const factor = 10 ** digits;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function compareBoxNo(a, b) {
  return boxNumber(a) - boxNumber(b);
}

function boxNumber(value) {
  const match = String(value).match(/(\d+)\D*$/);
  return match ? Number(match[1]) : 0;
}

function safeSheetName(name) {
  return String(name || "Sheet").replace(/[\\/?*\[\]:]/g, " ").slice(0, 31) || "Sheet";
}

function mskuKey(msku) {
  return normalize(msku).toUpperCase();
}

function formatDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}${m}${d}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function xmlEscape(value) {
  return escapeHtml(value).replace(/'/g, "&apos;");
}
