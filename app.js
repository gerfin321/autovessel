const STORAGE_KEY = "autovessel-web-mappings-v1";

const state = {
  pendingDecision: null,
};

const elements = {
  fileInput: document.getElementById("xlsxFile"),
  fileName: document.getElementById("fileName"),
  addBlankColumn: document.getElementById("addBlankColumn"),
  processButton: document.getElementById("processButton"),
  clearMappingsButton: document.getElementById("clearMappingsButton"),
  statusText: document.getElementById("statusText"),
  mappingCount: document.getElementById("mappingCount"),
  logOutput: document.getElementById("logOutput"),
  modal: document.getElementById("decisionModal"),
  decisionBody: document.getElementById("decisionBody"),
  mergeOnceButton: document.getElementById("mergeOnceButton"),
  mergeRememberButton: document.getElementById("mergeRememberButton"),
  keepSeparateButton: document.getElementById("keepSeparateButton"),
};

const SimilarityDecision = {
  MERGE_ONCE: "merge_once",
  MERGE_AND_REMEMBER: "merge_and_remember",
  KEEP_SEPARATE: "keep_separate",
};

class RegistryStorage {
  constructor() {
    this.mappings = this.load();
  }

  load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        return {};
      }

      const parsed = JSON.parse(raw);
      return parsed && typeof parsed === "object" ? parsed : {};
    } catch {
      return {};
    }
  }

  save() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(this.mappings));
    updateMappingCount();
  }

  clear() {
    this.mappings = {};
    this.save();
  }

  resolve(aliasKey) {
    return this.mappings[aliasKey] || null;
  }

  setMapping(aliasKey, canonicalDisplay) {
    this.mappings[aliasKey] = canonicalDisplay;
  }

  allMappings() {
    return this.mappings;
  }
}

class SimilarityChecker {
  findPotentialMatch(candidate, existingCanonicals) {
    const candidateKey = collapseForCompare(candidate);
    if (candidateKey.length < 4) {
      return null;
    }

    let best = null;
    for (const existing of existingCanonicals) {
      const existingKey = collapseForCompare(existing);
      if (existingKey.length < 4 || existingKey === candidateKey) {
        continue;
      }

      const distance = levenshteinDistance(candidateKey, existingKey);
      const scale = Math.max(candidateKey.length, existingKey.length);
      const similarity = scale > 0 ? 1 - distance / scale : 0;
      const isContainment = candidateKey.includes(existingKey) || existingKey.includes(candidateKey);

      if (
        (similarity >= 0.84 && Math.abs(candidateKey.length - existingKey.length) <= 4) ||
        (isContainment && Math.min(candidateKey.length, existingKey.length) >= 6)
      ) {
        if (!best || similarity > best.score) {
          best = { candidate, existing, score: similarity };
        }
      }
    }

    return best;
  }
}

class AgencyNormalizer {
  constructor(storage, similarityChecker) {
    this.storage = storage;
    this.similarityChecker = similarityChecker;
    this.sessionMappings = {};
    this.canonicalDisplayByKey = {};
    this.reloadFromStorage();
  }

  reloadFromStorage() {
    this.sessionMappings = {};
    this.canonicalDisplayByKey = {};

    for (const [alias, canonical] of Object.entries(this.storage.allMappings())) {
      this.canonicalDisplayByKey[normalizeAgencyKey(canonical)] = canonical;
      this.sessionMappings[alias] = canonical;
    }
  }

  async resolve(rawAgency) {
    const cleanedDisplay = cleanAgencyDisplay(rawAgency);
    const aliasKey = normalizeAgencyKey(cleanedDisplay);
    if (!aliasKey) {
      return {
        canonicalKey: "",
        canonicalDisplay: "",
        mergedFrom: null,
        persisted: false,
      };
    }

    const resolvedAlias = this.resolveAlias(aliasKey);
    if (resolvedAlias) {
      this.ensureCanonicalKnown(resolvedAlias);
      const canonicalKey = normalizeAgencyKey(resolvedAlias);
      return {
        canonicalKey,
        canonicalDisplay: resolvedAlias,
        mergedFrom: canonicalKey !== aliasKey ? cleanedDisplay : null,
        persisted: canonicalKey !== aliasKey && Boolean(this.storage.resolve(aliasKey)),
      };
    }

    if (this.canonicalDisplayByKey[aliasKey]) {
      return {
        canonicalKey: aliasKey,
        canonicalDisplay: this.canonicalDisplayByKey[aliasKey],
        mergedFrom: null,
        persisted: false,
      };
    }

    const match = this.similarityChecker.findPotentialMatch(cleanedDisplay, this.knownCanonicals());
    if (match) {
      const decision = await askSimilarityDecision(match);
      if (decision === SimilarityDecision.MERGE_ONCE || decision === SimilarityDecision.MERGE_AND_REMEMBER) {
        const canonicalKey = normalizeAgencyKey(match.existing);
        this.sessionMappings[aliasKey] = match.existing;
        this.ensureCanonicalKnown(match.existing);

        if (decision === SimilarityDecision.MERGE_AND_REMEMBER) {
          this.storage.setMapping(aliasKey, match.existing);
          this.storage.save();
        }

        return {
          canonicalKey,
          canonicalDisplay: match.existing,
          mergedFrom: cleanedDisplay,
          persisted: decision === SimilarityDecision.MERGE_AND_REMEMBER,
        };
      }
    }

    this.sessionMappings[aliasKey] = cleanedDisplay;
    this.canonicalDisplayByKey[aliasKey] = cleanedDisplay;
    return {
      canonicalKey: aliasKey,
      canonicalDisplay: cleanedDisplay,
      mergedFrom: null,
      persisted: false,
    };
  }

  ensureCanonicalKnown(canonicalDisplay) {
    this.canonicalDisplayByKey[normalizeAgencyKey(canonicalDisplay)] = canonicalDisplay;
  }

  resolveAlias(aliasKey) {
    if (this.sessionMappings[aliasKey]) {
      return this.sessionMappings[aliasKey];
    }
    return this.storage.resolve(aliasKey);
  }

  knownCanonicals() {
    return Object.values(this.canonicalDisplayByKey);
  }
}

function boot() {
  if (typeof window.XLSX === "undefined" || typeof window.JSZip === "undefined") {
    setStatus("Не удалось загрузить библиотеки");
    log("Ошибка: внешние библиотеки XLSX или JSZip не загрузились.");
    elements.processButton.disabled = true;
    return;
  }

  updateMappingCount();
  log("Веб-версия готова. Выберите XLSX и нажмите «Собрать ZIP с DOCX».");

  elements.fileInput.addEventListener("change", handleFileSelection);
  elements.processButton.addEventListener("click", handleProcess);
  elements.clearMappingsButton.addEventListener("click", handleClearMappings);
  elements.mergeOnceButton.addEventListener("click", () => resolveDecision(SimilarityDecision.MERGE_ONCE));
  elements.mergeRememberButton.addEventListener("click", () => resolveDecision(SimilarityDecision.MERGE_AND_REMEMBER));
  elements.keepSeparateButton.addEventListener("click", () => resolveDecision(SimilarityDecision.KEEP_SEPARATE));
}

function handleFileSelection(event) {
  const file = event.target.files && event.target.files[0];
  elements.fileName.textContent = file ? file.name : "Файл ещё не выбран";
}

async function handleProcess() {
  const file = elements.fileInput.files && elements.fileInput.files[0];
  if (!file) {
    setStatus("Нужно выбрать XLSX");
    alert("Сначала выберите XLSX-файл.");
    return;
  }

  setBusy(true);
  setStatus("Идёт обработка...");
  log("");
  log("===== Новый запуск =====");
  log(`Входной файл: ${file.name}`);
  log(
    elements.addBlankColumn.checked
      ? "Опция включена: добавить третий пустой столбец."
      : "Опция выключена: не добавлять третий пустой столбец."
  );

  try {
    const report = await processWorkbook(file, elements.addBlankColumn.checked);
    log("Итог обработки");
    log(`Прочитано строк: ${report.totalRowsRead}`);
    log(`Успешно обработано: ${report.processedRows}`);
    log(`Создано документов: ${report.documentsCreated}`);
    log(`Архив скачан: ${report.archiveName}`);

    if (report.createdFiles.length) {
      log("Созданные файлы:");
      for (const fileName of report.createdFiles) {
        log(` - ${fileName}`);
      }
    }

    if (report.skippedRows.length) {
      log("Пропущенные строки:");
      for (const skipped of report.skippedRows) {
        let details = ` - Строка ${skipped.worksheetRow}: ${skipped.reason}`;
        if (skipped.vesselText) {
          details += ` [Судно: ${skipped.vesselText}]`;
        }
        if (skipped.agencyText) {
          details += ` [Компания: ${skipped.agencyText}]`;
        }
        log(details);
      }
    }

    if (report.merges.length) {
      log("Объединённые компании:");
      for (const merge of report.merges) {
        log(` - ${merge.fromName} -> ${merge.toName}${merge.persisted ? " (запомнено)" : ""}`);
      }
    }

    setStatus("Готово. ZIP скачан.");
  } catch (error) {
    const message = error instanceof Error ? error.message : "Неизвестная ошибка.";
    log(`Ошибка: ${message}`);
    setStatus("Ошибка обработки");
    alert("Во время обработки произошла ошибка. Подробности есть в журнале.");
  } finally {
    setBusy(false);
  }
}

async function processWorkbook(file, addBlankColumn) {
  const storage = new RegistryStorage();
  const similarityChecker = new SimilarityChecker();
  const agencyNormalizer = new AgencyNormalizer(storage, similarityChecker);

  log("Чтение XLSX...");
  const arrayBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("В XLSX не найдено ни одного листа.");
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const worksheetData = extractWorksheetData(worksheet);

  log("Поиск заголовков...");
  const headers = detectHeaders(worksheetData);

  const report = {
    totalRowsRead: 0,
    processedRows: 0,
    documentsCreated: 0,
    createdFiles: [],
    skippedRows: [],
    merges: [],
    archiveName: buildArchiveName(file.name),
  };

  log("Обработка строк...");
  const parsedRows = [];
  let orderCounter = 0;
  for (let rowIndex = headers.headerRowIndex + 1; rowIndex < worksheetData.rows.length; rowIndex += 1) {
    const row = worksheetData.rows[rowIndex];
    if (isRowEmpty(row)) {
      continue;
    }

    report.totalRowsRead += 1;

    const dateCell = findCell(row, headers.dateColumn);
    const vesselCell = findCell(row, headers.vesselColumn);
    const agencyCell = findCell(row, headers.agencyColumn);

    const dateText = safeCellText(dateCell);
    const vesselText = collapseWhitespace(safeCellText(vesselCell));
    const agencyText = cleanAgencyDisplay(safeCellText(agencyCell));

    if (!vesselText || !agencyText) {
      report.skippedRows.push({
        worksheetRow: row.rowNumber,
        reason: !vesselText ? "Пропущено: нет названия судна." : "Пропущено: нет агентирующей компании.",
        dateText,
        vesselText,
        agencyText,
      });
      continue;
    }

    const vesselParts = parseVesselName(vesselText);
    const agency = await agencyNormalizer.resolve(agencyText);
    if (agency.mergedFrom && agency.canonicalDisplay) {
      addMergeRecord(report, agency.mergedFrom, agency.canonicalDisplay, agency.persisted);
    }

    parsedRows.push({
      originalOrder: orderCounter,
      worksheetRow: row.rowNumber,
      dateText,
      dateValue: dateCell ? parseDate(dateCell) : null,
      vesselName: vesselParts.name,
      dwt: vesselParts.dwt,
      originalAgency: agencyText,
      canonicalAgencyKey: agency.canonicalKey,
      canonicalAgencyDisplay: agency.canonicalDisplay,
    });
    orderCounter += 1;
  }

  report.processedRows = parsedRows.length;

  log("Сортировка и группировка...");
  parsedRows.sort(compareParsedRows);
  const groups = new Map();
  for (const row of parsedRows) {
    if (!groups.has(row.canonicalAgencyKey)) {
      groups.set(row.canonicalAgencyKey, {
        key: row.canonicalAgencyKey,
        displayName: row.canonicalAgencyDisplay,
        rows: [],
      });
    }
    groups.get(row.canonicalAgencyKey).rows.push(row);
  }

  log("Создание DOCX...");
  const outerZip = new JSZip();
  const usedFileNames = new Set();
  const sortedGroups = Array.from(groups.values()).sort((left, right) =>
    left.displayName.localeCompare(right.displayName, "uk")
  );

  for (const group of sortedGroups) {
    const baseName = sanitizeFileName(group.displayName, "Компания");
    const uniqueName = ensureUniqueName(baseName, usedFileNames);
    const docxFileName = `${uniqueName}.docx`;
    const docxArrayBuffer = await buildDocxArrayBuffer(group, addBlankColumn);
    outerZip.file(docxFileName, docxArrayBuffer);
    report.createdFiles.push(docxFileName);
  }

  report.documentsCreated = report.createdFiles.length;
  outerZip.file("processing-report.txt", buildReportText(report));

  const zipBlob = await outerZip.generateAsync({ type: "blob", compression: "DEFLATE" });
  downloadBlob(zipBlob, report.archiveName);
  return report;
}

function extractWorksheetData(sheet) {
  const data = { rows: [] };
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = {
      rowNumber: rowIndex + 1,
      cells: {},
    };

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      const cell = sheet[address];
      if (!cell) {
        continue;
      }

      const formatted = typeof cell.w === "string" ? cell.w : formatFallbackCell(cell);
      row.cells[columnIndex] = {
        text: collapseWhitespace(formatted),
        empty: !formatted,
      };
    }

    data.rows.push(row);
  }

  return data;
}

function formatFallbackCell(cell) {
  if (cell == null || cell.v == null) {
    return "";
  }
  if (typeof cell.v === "string") {
    return cell.v;
  }
  return String(cell.v);
}

function detectHeaders(data) {
  let best = {
    headerRowIndex: 0,
    dateColumn: -1,
    vesselColumn: -1,
    agencyColumn: -1,
  };
  let bestScore = -1;

  const scanLimit = Math.min(data.rows.length, 15);
  for (let rowIndex = 0; rowIndex < scanLimit; rowIndex += 1) {
    const row = data.rows[rowIndex];
    const current = {
      headerRowIndex: rowIndex,
      dateColumn: -1,
      vesselColumn: -1,
      agencyColumn: -1,
    };

    let score = 0;
    for (const [columnKey, cell] of Object.entries(row.cells)) {
      const columnIndex = Number(columnKey);
      const normalized = normalizeHeader(cell.text);

      if (current.dateColumn < 0 && isDateHeader(normalized)) {
        current.dateColumn = columnIndex;
        score += 1;
      }
      if (current.vesselColumn < 0 && isVesselHeader(normalized)) {
        current.vesselColumn = columnIndex;
        score += 1;
      }
      if (current.agencyColumn < 0 && isAgencyHeader(normalized)) {
        current.agencyColumn = columnIndex;
        score += 1;
      }
    }

    if (score > bestScore) {
      best = current;
      bestScore = score;
    }

    if (score === 3) {
      break;
    }
  }

  if (best.dateColumn < 0 || best.vesselColumn < 0 || best.agencyColumn < 0) {
    throw new Error("Не удалось автоматически найти колонки даты, судна и компании.");
  }

  return best;
}

function isDateHeader(value) {
  return containsInsensitive(value, "дата");
}

function isVesselHeader(value) {
  return value === "Назва судна" || value === "Назва судна" ||
    (containsInsensitive(value, "наз") && containsInsensitive(value, "суд"));
}

function isAgencyHeader(value) {
  return containsInsensitive(value, "агент") &&
    (containsInsensitive(value, "фирм") || containsInsensitive(value, "компан"));
}

function parseVesselName(rawValue) {
  const cleaned = collapseWhitespace(rawValue);
  if (!cleaned) {
    return { name: "", dwt: "" };
  }

  const dwtMarker = /^(.*?)(?:\s+|\()DWT[\s:.-]*(\d{3,6}(?:[.,]\d{1,2})?)\)?\s*$/i;
  const parenthesizedNumber = /^(.*?)[\s]*\((\d{4,6}(?:[.,]\d{1,2})?)\)\s*$/i;
  const trailingNumber = /^(.*?\D)\s+(\d{4,6}(?:[.,]\d{1,2})?)\s*$/i;

  const dwtMatch = cleaned.match(dwtMarker);
  if (dwtMatch && containsLetter(trim(dwtMatch[1]))) {
    return {
      name: trim(dwtMatch[1]),
      dwt: normalizeRoundedDwt(dwtMatch[2]),
    };
  }

  const parenthesizedMatch = cleaned.match(parenthesizedNumber);
  if (parenthesizedMatch && containsLetter(trim(parenthesizedMatch[1]))) {
    return {
      name: trim(parenthesizedMatch[1]),
      dwt: normalizeRoundedDwt(parenthesizedMatch[2]),
    };
  }

  const trailingMatch = cleaned.match(trailingNumber);
  if (trailingMatch && containsLetter(trim(trailingMatch[1]))) {
    return {
      name: trim(trailingMatch[1]),
      dwt: normalizeRoundedDwt(trailingMatch[2]),
    };
  }

  return {
    name: cleaned,
    dwt: "",
  };
}

function normalizeRoundedDwt(value) {
  const normalized = value.replace(/\s+/g, "").replace(",", ".");
  return String(Math.round(Number(normalized)));
}

function parseDate(cell) {
  return parseDateText(cell.text);
}

function parseDateText(text) {
  const cleaned = trim(text);
  if (!cleaned) {
    return null;
  }

  let match = cleaned.match(/^(\d{4})[./-](\d{1,2})[./-](\d{1,2})$/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    return isValidDate(year, month, day) ? { year, month, day } : null;
  }

  match = cleaned.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (match) {
    let year = Number(match[3]);
    if (year < 100) {
      year += year >= 70 ? 1900 : 2000;
    }
    const month = Number(match[2]);
    const day = Number(match[1]);
    return isValidDate(year, month, day) ? { year, month, day } : null;
  }

  return null;
}

function isValidDate(year, month, day) {
  if (year < 1900 || year > 2100 || month < 1 || month > 12 || day < 1 || day > 31) {
    return false;
  }
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function compareParsedRows(left, right) {
  if (left.dateValue && right.dateValue) {
    const leftKey = left.dateValue.year * 10000 + left.dateValue.month * 100 + left.dateValue.day;
    const rightKey = right.dateValue.year * 10000 + right.dateValue.month * 100 + right.dateValue.day;
    if (leftKey !== rightKey) {
      return leftKey - rightKey;
    }
  } else if (left.dateValue && !right.dateValue) {
    return -1;
  } else if (!left.dateValue && right.dateValue) {
    return 1;
  }

  return left.originalOrder - right.originalOrder;
}

async function buildDocxArrayBuffer(group, addBlankColumn) {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", buildContentTypesXml());
  zip.file("_rels/.rels", buildRootRelsXml());
  zip.file("word/document.xml", buildDocumentXml(group, addBlankColumn));
  zip.file("word/styles.xml", buildStylesXml());
  zip.file("word/_rels/document.xml.rels", buildDocumentRelsXml());
  zip.file("docProps/core.xml", buildCoreXml());
  zip.file("docProps/app.xml", buildAppXml());
  return zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
}

function buildDocumentXml(group, addBlankColumn) {
  let xml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
  xml += "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>";
  xml += buildParagraph(group.displayName, true, true, 28);
  xml += "<w:tbl><w:tblPr><w:tblW w:w=\"0\" w:type=\"auto\"/>";
  xml += "<w:tblBorders>";
  xml += "<w:top w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "<w:left w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "<w:bottom w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "<w:right w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "<w:insideH w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "<w:insideV w:val=\"single\" w:sz=\"12\" w:space=\"0\" w:color=\"000000\"/>";
  xml += "</w:tblBorders></w:tblPr>";

  xml += "<w:tr>";
  xml += buildCell("Название", addBlankColumn ? 6200 : 7200, true, true);
  xml += buildCell("DWT", 1800, true, true);
  if (addBlankColumn) {
    xml += buildCell("", 1800, true, true);
  }
  xml += "</w:tr>";

  for (const row of group.rows) {
    xml += "<w:tr>";
    xml += buildCell(row.vesselName, addBlankColumn ? 6200 : 7200, false, false);
    xml += buildCell(row.dwt, 1800, false, true);
    if (addBlankColumn) {
      xml += buildCell("", 1800, false, true);
    }
    xml += "</w:tr>";
  }

  xml += "</w:tbl>";
  xml += "<w:sectPr>";
  xml += "<w:pgSz w:w=\"11906\" w:h=\"16838\"/>";
  xml += "<w:pgMar w:top=\"1134\" w:right=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:header=\"708\" w:footer=\"708\" w:gutter=\"0\"/>";
  xml += "</w:sectPr></w:body></w:document>";
  return xml;
}

function buildParagraph(text, bold, centered, sizeHalfPoints) {
  let xml = "<w:p>";
  if (centered) {
    xml += "<w:pPr><w:jc w:val=\"center\"/></w:pPr>";
  }
  xml += "<w:r><w:rPr>";
  if (bold) {
    xml += "<w:b/>";
  }
  xml += `<w:sz w:val="${sizeHalfPoints}"/>`;
  xml += "<w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/>";
  xml += `</w:rPr><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
  return xml;
}

function buildCell(text, width, header, center) {
  let xml = `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p>`;
  if (center) {
    xml += "<w:pPr><w:jc w:val=\"center\"/></w:pPr>";
  }
  xml += "<w:r><w:rPr>";
  if (header) {
    xml += "<w:b/>";
  }
  xml += "<w:sz w:val=\"24\"/><w:rFonts w:ascii=\"Times New Roman\" w:hAnsi=\"Times New Roman\" w:cs=\"Times New Roman\"/></w:rPr>";
  if (text) {
    xml += `<w:t xml:space="preserve">${escapeXml(text)}</w:t>`;
  }
  xml += "</w:r></w:p></w:tc>";
  return xml;
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
        <w:sz w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>`;
}

function buildContentTypesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
}

function buildRootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function buildDocumentRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function buildCoreXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Autovessel Web</dc:title>
  <dc:creator>gerfin / Andrii Morgun</dc:creator>
</cp:coreProperties>`;
}

function buildAppXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Autovessel Web</Application>
</Properties>`;
}

function buildReportText(report) {
  const lines = [
    "Autovessel Web - processing report",
    "",
    `Прочитано строк: ${report.totalRowsRead}`,
    `Успешно обработано: ${report.processedRows}`,
    `Создано документов: ${report.documentsCreated}`,
    "",
    "Файлы:",
    ...report.createdFiles.map((fileName) => `- ${fileName}`),
  ];
  return lines.join("\r\n");
}

function buildArchiveName(originalFileName) {
  const baseName = originalFileName.replace(/\.[^.]+$/, "");
  const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `${sanitizeFileName(baseName, "autovessel")}_${timestamp}.zip`;
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function askSimilarityDecision(match) {
  elements.decisionBody.textContent =
    `Найдены похожие названия компаний:\n\n"${match.candidate}" и "${match.existing}".\n\nЧто сделать?`;

  elements.modal.classList.remove("hidden");
  return new Promise((resolve) => {
    state.pendingDecision = resolve;
  });
}

function resolveDecision(decision) {
  if (typeof state.pendingDecision === "function") {
    state.pendingDecision(decision);
    state.pendingDecision = null;
  }
  elements.modal.classList.add("hidden");
}

function handleClearMappings() {
  if (!confirm("Очистить все запомненные объединения компаний в браузере?")) {
    return;
  }

  const storage = new RegistryStorage();
  storage.clear();
  log("Сохранённые объединения очищены.");
  setStatus("Память объединений очищена");
}

function updateMappingCount() {
  const storage = new RegistryStorage();
  elements.mappingCount.textContent = String(Object.keys(storage.allMappings()).length);
}

function setBusy(isBusy) {
  elements.processButton.disabled = isBusy;
  elements.clearMappingsButton.disabled = isBusy;
  elements.fileInput.disabled = isBusy;
  elements.addBlankColumn.disabled = isBusy;
}

function setStatus(text) {
  elements.statusText.textContent = text;
}

function log(message) {
  const line = message
    ? `[${new Date().toLocaleTimeString("ru-RU")}] ${message}`
    : "";
  elements.logOutput.textContent += `${line}\n`;
  elements.logOutput.scrollTop = elements.logOutput.scrollHeight;
}

function addMergeRecord(report, from, to, persisted) {
  const exists = report.merges.some(
    (item) => iEquals(item.fromName, from) && iEquals(item.toName, to) && item.persisted === persisted
  );
  if (!exists) {
    report.merges.push({ fromName: from, toName: to, persisted });
  }
}

function ensureUniqueName(baseName, usedNames) {
  let candidate = baseName;
  let suffix = 1;
  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${baseName}_${suffix}`;
    suffix += 1;
  }
  usedNames.add(candidate.toLowerCase());
  return candidate;
}

function findCell(row, columnIndex) {
  return row.cells[columnIndex] || null;
}

function safeCellText(cell) {
  return cell ? cell.text : "";
}

function isRowEmpty(row) {
  return Object.values(row.cells).every((cell) => isBlank(cell.text));
}

function collapseForCompare(value) {
  return normalizeAgencyKey(value).replace(/\s+/g, "");
}

function levenshteinDistance(left, right) {
  if (!left.length) {
    return right.length;
  }
  if (!right.length) {
    return left.length;
  }

  let previous = new Array(right.length + 1);
  let current = new Array(right.length + 1);

  for (let j = 0; j <= right.length; j += 1) {
    previous[j] = j;
  }

  for (let i = 1; i <= left.length; i += 1) {
    current[0] = i;
    for (let j = 1; j <= right.length; j += 1) {
      const substitution = previous[j - 1] + (left[i - 1] === right[j - 1] ? 0 : 1);
      const insertion = current[j - 1] + 1;
      const deletion = previous[j] + 1;
      current[j] = Math.min(substitution, insertion, deletion);
    }
    [previous, current] = [current, previous];
  }

  return previous[right.length];
}

function trim(value) {
  return (value || "").trim();
}

function collapseWhitespace(value) {
  return trim((value || "").replace(/\s+/g, " "));
}

function toLower(value) {
  return (value || "").toLocaleLowerCase("uk-UA");
}

function toUpper(value) {
  return (value || "").toLocaleUpperCase("uk-UA");
}

function stripEdgeSymbols(value) {
  return (value || "").replace(/^[^\p{L}\p{N}_-]+|[^\p{L}\p{N}_-]+$/gu, "");
}

function normalizeHeader(value) {
  return collapseWhitespace(
    toLower(collapseWhitespace(value)).replace(/[^\p{L}\p{N}]/gu, " ")
  );
}

function normalizeAgencyKey(value) {
  return collapseWhitespace(
    toUpper(collapseWhitespace(stripEdgeSymbols(value))).replace(/[^\p{L}\p{N}]/gu, " ")
  );
}

function cleanAgencyDisplay(value) {
  return collapseWhitespace(stripEdgeSymbols(value));
}

function sanitizeFileName(value, fallback) {
  let cleaned = trim(value);
  if (!cleaned) {
    cleaned = fallback;
  }

  cleaned = cleaned.replace(/[<>:"/\\|?*\u0000-\u001F]/g, "_").trim().replace(/[. ]+$/g, "");
  return cleaned || fallback;
}

function escapeXml(value) {
  return (value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function containsInsensitive(haystack, needle) {
  return toLower(haystack).includes(toLower(needle));
}

function iEquals(left, right) {
  return toLower(left) === toLower(right);
}

function isBlank(value) {
  return !trim(value);
}

function containsLetter(value) {
  return /\p{L}/u.test(value);
}

boot();
