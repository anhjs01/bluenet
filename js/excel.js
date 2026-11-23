export function rows(u) {
  return u.map(x => ({
    "N°": x.sequence,
    "Ticket": x.ticket || "",
    "ID": x.unitId || "",
    "Fabricante": x.manufacturer || "",
    "Fecha del lote": x.lotDate || "",
    "Tipo de lectura": x.readType || "",
    "Cantidad": 1,
    "Estado diagnóstico": x.diagnosis || "",
    "Reparable":
      x.repairable === true
        ? "Sí"
        : x.repairable === false
          ? "No"
          : "",
    "Motivo no reparable":
      (x.noRepairReasons || []).join(" + "),
    "Reparación / mantenimiento":
      (x.repairs || []).join(" + "),
    "Observaciones":
      x.observations || "",
    "Empaque":
      x.packaging || ""
  }));
}

function dateForExcel(value) {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(value + (String(value).length === 10 ? "T00:00:00" : ""));
  if (Number.isNaN(d.getTime())) return "";
  return [String(d.getDate()).padStart(2, "0"), String(d.getMonth() + 1).padStart(2, "0"), d.getFullYear()].join("/");
}

function delivered(unit) {
  const status = String(unit.deliveryStatus || unit.status || "").trim().toLocaleLowerCase("es");
  const observations = String(unit.observations || "");
  return unit.delivered === true
    || status === "entregada"
    || status === "entregado"
    || status === "delivered"
    || Boolean(unit.deliveryDate || unit.deliveredAt)
    || /(^|\s|·)ENTREGAD[OA](\s|·|$)/i.test(observations);
}

function warrantyExpiration(unit, isDelivered) {
  if (!isDelivered) return "";
  if (unit.warrantyUntil) return dateForExcel(unit.warrantyUntil);
  const deliveryDate = unit.deliveryDate || unit.deliveredAt || "";
  if (!deliveryDate) return "";
  const match = String(deliveryDate).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) return "";
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  if (Number.isNaN(date.getTime())) return "";
  const months = Math.max(0, Number(unit.warrantyMonths ?? unit.warrantyDuration ?? 2) || 2);
  date.setMonth(date.getMonth() + months);
  return dateForExcel([date.getFullYear(), String(date.getMonth() + 1).padStart(2, "0"), String(date.getDate()).padStart(2, "0")].join("-"));
}

function cloneCellStyle(source, target) {
  if (source?.s) target.s = JSON.parse(JSON.stringify(source.s));
  if (source?.z) target.z = source.z;
}

function xmlEscape(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function excelSerialDate(value) {
  if (!value) return null;
  const text = String(value).trim();
  const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const utc = Date.UTC(year, month - 1, day);
    const check = new Date(utc);
    if (check.getUTCFullYear() !== year || check.getUTCMonth() !== month - 1 || check.getUTCDate() !== day) return null;
    return utc / 86400000 + 25569;
  }
  const parsed = new Date(value);
  if (Number.isNaN(parsed.getTime())) return null;
  return Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()) / 86400000 + 25569;
}

function excelCellStyle(sheetXml, address) {
  const escaped = address.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const match = sheetXml.match(new RegExp(`<c\\b(?=[^>]*\\br="${escaped}")[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)`));
  if (!match) return "";
  const open = match[0].match(/^<c\b([^>]*)/);
  if (!open) return "";
  const style = open[1].match(/\bs="([^"]+)"/);
  return style ? ` s="${style[1]}"` : "";
}

function setExcelCell(sheetXml, address, value, styleFallback = "") {
  const rowNumber = Number(address.match(/\d+$/)?.[0]);
  const column = address.match(/^[A-Z]+/)?.[0] || "A";
  const columnNumber = letters => [...letters].reduce((n, letter) => n * 26 + letter.charCodeAt(0) - 64, 0);
  const rowMatches = [...sheetXml.matchAll(/<row\b[^>]*>[\s\S]*?<\/row>/g)];
  const rowMatch = rowMatches.find(match => {
    const rowId = match[0].match(/^<row\b[^>]*\br="(\d+)"/);
    return rowId && Number(rowId[1]) === rowNumber;
  });
  if (!rowMatch) return sheetXml;

  const rowXml = rowMatch[0];
  const cellMatches = [...rowXml.matchAll(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g)];
  const existing = cellMatches.find(match => {
    const ref = match[0].match(/^<c\b[^>]*\br="([A-Z]+\d+)"/);
    return ref && ref[1] === address;
  });
  let attributes = existing
    ? (existing[0].match(/^<c\b([^>]*)/) || [null, ""])[1]
    : ` r="${address}"${styleFallback}`;
  attributes = attributes
    .replace(/\s+t="[^"]*"/g, "")
    .replace(/\s+cm="[^"]*"/g, "")
    .replace(/\s+vm="[^"]*"/g, "")
    .replace(/\/\s*$/, "");

  let cellXml;
  if (value === null || value === undefined || value === "") {
    cellXml = `<c${attributes}/>`;
  } else if (typeof value === "number" && Number.isFinite(value)) {
    cellXml = `<c${attributes}><v>${value}</v></c>`;
  } else {
    cellXml = `<c${attributes} t="inlineStr"><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
  }

  let updatedRow;
  if (existing) {
    const start = existing.index;
    updatedRow = rowXml.slice(0, start) + cellXml + rowXml.slice(start + existing[0].length);
  } else {
    const cells = cellMatches.map(match => match[0]);
    cells.push(cellXml);
    cells.sort((left, right) => {
      const leftRef = left.match(/^<c\b[^>]*\br="([A-Z]+\d+)"/)?.[1] || "A1";
      const rightRef = right.match(/^<c\b[^>]*\br="([A-Z]+\d+)"/)?.[1] || "A1";
      return columnNumber(leftRef.match(/^[A-Z]+/)?.[0] || "A") - columnNumber(rightRef.match(/^[A-Z]+/)?.[0] || "A");
    });
    const openTag = rowXml.match(/^<row\b[^>]*>/)?.[0] || `<row r="${rowNumber}">`;
    const closeTag = "</row>";
    updatedRow = openTag + cells.join("") + closeTag;
  }

  return sheetXml.slice(0, rowMatch.index) + updatedRow + sheetXml.slice(rowMatch.index + rowXml.length);
}

function safeExcelFileName(value) {
  const cleaned = String(value ?? "")
    .normalize("NFC")
    .replace(/[\\/:*?"<>|\u0000-\u001F]/g, "-")
    .replace(/\s+/g, " ")
    .replace(/[. ]+$/g, "")
    .trim();
  return (cleaned || "conteo-rapido-remision") + ".xlsx";
}

function shiftWorksheetRows(xml, startRow, offset) {
  if (!offset) return xml;
  const shiftRefs = value => value.replace(/(\$?[A-Z]{1,3}\$?)(\d+)/g, (match, column, row) => {
    const number = Number(row);
    return column + (number >= startRow ? number + offset : number);
  });
  xml = xml.replace(/\br="(\d+)"/g, (match, row) => {
    const number = Number(row);
    return number >= startRow ? 'r="' + (number + offset) + '"' : match;
  });
  xml = xml.replace(/\bref="([^"]+)"/g, (match, ref) => 'ref="' + shiftRefs(ref) + '"');
  xml = xml.replace(/<f\b([^>]*)>([\s\S]*?)<\/f>/g, (match, attrs, formula) => '<f' + attrs + '>' + shiftRefs(formula) + '</f>');
  return xml;
}

export async function exportExcel(u, lotName = "") {
  if (!window.JSZip || !window.XLSX) return alert("No están disponibles los módulos necesarios para conservar la plantilla Excel original.");

  let sourceZip;
  let headerColumns;
  try {
    const response = await fetch("./guias/prueba01.xlsx", { cache: "no-store" });
    if (!response.ok) throw new Error("No se pudo cargar la nueva plantilla Prueba 01.");
    const sourceBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(sourceBuffer, { type: "array" });
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: "", blankrows: true });
    const normalize = value => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
    const required = ["article", "serial", "brand", "request", "maintenance", "delivery"];
    const identifyHeaders = row => {
      const found = {};
      (row || []).slice(0, 8).forEach((value, index) => {
        const h = normalize(value);
        if (/articulo|article|artic/.test(h)) found.article = index;
        else if (/serie|serial|s n|nro.*serie|numero.*de.*serie/.test(h)) found.serial = index;
        else if (/marca|fabricante|brand/.test(h)) found.brand = index;
        else if (/solicitud.*reparacion|solicitud|ticket|request/.test(h)) found.request = index;
        else if (/tipo.*mantenimiento|mantenimiento|maintenance/.test(h)) found.maintenance = index;
        else if (/fecha.*entrega|fecha entrega|delivery date|delivery/.test(h)) found.delivery = index;
      });
      return found;
    };
    // Priorizar la fila 5 y aceptar encabezados abreviados como "Artículo" y "Serial".
    headerColumns = identifyHeaders(matrix[4]);
    const countHeaders = found => required.filter(key => found[key] !== undefined).length;
    // Algunas copias de la plantilla desplazan los títulos por celdas combinadas.
    if (countHeaders(headerColumns) < required.length) {
      const alternatives = matrix.slice(0, 8).map((row, index) => ({ row: index + 1, found: identifyHeaders(row) }));
      alternatives.sort((a, b) => countHeaders(b.found) - countHeaders(a.found));
      if (countHeaders(alternatives[0]?.found || {}) > countHeaders(headerColumns)) headerColumns = alternatives[0].found;
    }
    // Si la plantilla usa un título no estándar para el número de serie,
    // inferir la columna inmediatamente posterior a "Artículo" solo si está libre.
    if (headerColumns.serial === undefined && headerColumns.article !== undefined) {
      const inferredSerialColumn = headerColumns.article + 1;
      const occupied = Object.entries(headerColumns).some(([key, value]) => key !== "serial" && value === inferredSerialColumn);
      if (inferredSerialColumn < 8 && !occupied) headerColumns.serial = inferredSerialColumn;
    }
    const missing = required.filter(key => headerColumns[key] === undefined || headerColumns[key] > 7);
    if (missing.length) {
      const labels = { article: "Artículo sin etiqueta", serial: "Número de serie", brand: "Marca", request: "Solicitud de reparación", maintenance: "Tipo de mantenimiento", delivery: "Fecha de entrega" };
      throw new Error("No se encontraron en A:H los encabezados: " + missing.map(key => labels[key]).join(", ") + ". Se revisó la fila 5 y las filas cercanas. No se descargó un Excel incompleto.");
    }
    sourceZip = await JSZip.loadAsync(sourceBuffer);
  } catch (error) {
    alert("No se pudo preparar la nueva plantilla Prueba 01.\n" + (error.message || ""));
    return;
  }

  try {
    const sheetPath = "xl/worksheets/sheet1.xml";
    const sheetFile = sourceZip.file(sheetPath);
    if (!sheetFile) throw new Error("La plantilla no contiene la primera hoja Excel esperada.");
    let sheetXml = await sheetFile.async("string");

    // Prueba 01: filas 1–5 y 66–68 quedan reservadas; los registros empiezan en la fila 6.
    const firstDataRow = 6;
    const lastExistingRow = 65;
    const firstInsertRow = 66;
    const lastDataCol = 8;
    const colName = number => {
      let n = number, name = "";
      while (n > 0) {
        const rem = (n - 1) % 26;
        name = String.fromCharCode(65 + rem) + name;
        n = Math.floor((n - 1) / 26);
      }
      return name;
    };

    const fallbackStyles = {};
    for (let c = 1; c <= lastDataCol; c++) {
      const col = colName(c);
      fallbackStyles[col] = excelCellStyle(sheetXml, col + firstDataRow);
    }

    const templateRowMatch = sheetXml.match(new RegExp('<row\\b(?=[^>]*\\br="' + firstDataRow + '")[^>]*>[\\s\\S]*?<\\/row>'));
    const templateRowXml = templateRowMatch?.[0] || "";
    if (!templateRowXml) throw new Error("No se encontró la fila 6 para extender el formato de la plantilla.");

    // Si hay más de 60 unidades, insertar filas antes del bloque reservado 66–68.
    // El contenido y el formato del bloque inferior se desplazan juntos, sin sobrescribirlos.
    const extraRowsCount = Math.max(0, u.length - (lastExistingRow - firstDataRow + 1));
    if (extraRowsCount) {
      sheetXml = shiftWorksheetRows(sheetXml, firstInsertRow, extraRowsCount);
      const extraRows = [];
      for (let row = firstInsertRow; row < firstInsertRow + extraRowsCount; row++) {
        let clone = templateRowXml.replace(/\br="6"/g, 'r="' + row + '"');
        for (let c = 1; c <= lastDataCol; c++) {
          const col = colName(c);
          clone = setExcelCell(clone, col + row, null, fallbackStyles[col]);
        }
        extraRows.push(clone);
      }
      const shiftedFooterRow = firstInsertRow + extraRowsCount;
      const footerMatch = new RegExp('<row\\b(?=[^>]*\\br="' + shiftedFooterRow + '")[^>]*>');
      const footerIndex = sheetXml.search(footerMatch);
      const insertAt = footerIndex >= 0 ? footerIndex : sheetXml.indexOf("</sheetData>");
      if (insertAt < 0) throw new Error("No se encontró el área de filas de la hoja Excel.");
      sheetXml = sheetXml.slice(0, insertAt) + extraRows.join("") + sheetXml.slice(insertAt);
    }

    // Limpiar únicamente A6:H65. El resto del libro conserva sus valores y formato.
    for (let row = firstDataRow; row <= lastExistingRow; row++) {
      for (let c = 1; c <= lastDataCol; c++) {
        const col = colName(c);
        sheetXml = setExcelCell(sheetXml, col + row, null, fallbackStyles[col]);
      }
    }

    const deliveryDate = unit => excelSerialDate(unit.deliveryDate || unit.deliveredAt || "");
    const valuesForUnit = (unit, index) => ({
      article: unit.sequence ?? index + 1,
      serial: String(unit.unitId || ""),
      brand: unit.manufacturer || "",
      request: unit.ticket || unit.observations || "",
      maintenance: (unit.repairs || []).join(" + ") || unit.diagnosis || "",
      delivery: deliveryDate(unit)
    });
    const keyToValue = { article: "article", serial: "serial", brand: "brand", request: "request", maintenance: "maintenance", delivery: "delivery" };

    for (let i = 0; i < u.length; i++) {
      const row = firstDataRow + i;
      const data = valuesForUnit(u[i], i);
      for (const key of Object.keys(keyToValue)) {
        const col = colName(headerColumns[key] + 1);
        sheetXml = setExcelCell(sheetXml, col + row, data[keyToValue[key]], fallbackStyles[col]);
      }
    }

    sourceZip.file(sheetPath, sheetXml);
    const finalBlob = await sourceZip.generateAsync({ type: "blob", compression: "DEFLATE" });
    const url = URL.createObjectURL(finalBlob);
    const link = document.createElement("a");
    link.href = url;
    link.download = safeExcelFileName(lotName);
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (error) {
    console.error("Error al generar el Excel con Prueba 01:", error);
    alert("No se pudo generar el Excel conservando intacto el diseño de Prueba 01. No se descargó un archivo incompleto.\n" + (error.message || ""));
  }
}

/*
  ============================================================
  IMPORTAR EXCEL
  ============================================================
*/

function clean(value) {
  return String(value ?? "").trim();
}

function splitValues(value) {
  return clean(value)
    .split("+")
    .map(x => x.trim())
    .filter(Boolean);
}

function stripDeliveryMarker(value) {
  return clean(value).replace(/^ENTREGADO(?:\s*·\s*|\s+)?/i, "").trim();
}

function normalizeImportedDate(value) {
  if (!value) return "";
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  const text = clean(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const match = text.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4})$/);
  if (match) return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
  return "";
}

function parseRepairable(value, diagnosis) {
  const v =
    clean(value).toLowerCase();

  if (
    v === "sí" ||
    v === "si" ||
    v === "yes" ||
    v === "true"
  ) {
    return true;
  }

  if (
    v === "no" ||
    v === "false"
  ) {
    return false;
  }

  return diagnosis !== "No reparable";
}

function normalizeDiagnosis(value) {
  const v =
    clean(value).toLowerCase();

  if (
    v === "no reparable" ||
    v === "no reparable "
  ) {
    return "No reparable";
  }

  return "Reparable";
}

/**
 * Convierte una fila del Excel exportado por
 * Conteo Rápido a una unidad de inventario.
 */
export function rowToUnit(row, index, lot) {
  const diagnosis =
    normalizeDiagnosis(
      row["Estado diagnóstico"] ?? row["Diagnóstico"]
    );

  const unitId =
    clean(row["ID"]);

  if (!unitId) {
    throw new Error(
      `La fila ${index + 2} no tiene ID.`
    );
  }

  const repairs =
    diagnosis === "No reparable"
      ? []
      : splitValues(
          row["Reparación / mantenimiento"] ?? row["Reparación / proceso"]
        );

  const reasons =
    diagnosis === "No reparable"
      ? splitValues(
          row[
            "Motivo no reparable"
          ]
        )
      : [];

  return {
    id:
      "unit-" +
      crypto.randomUUID(),

    unitId,

    sequence:
      index + 1,

    ticket:
      clean(row["Ticket"]),

    manufacturer:
      clean(row["Fabricante"]) ||
      "Pendiente",

    lotDate:
      normalizeImportedDate(row["Fecha del lote"] ?? row["Fecha recepción"]) || lot.date,

    readType:
      clean(row["Tipo de lectura"]) ||
      "Importado desde Excel",

    diagnosis,

    repairable:
      parseRepairable(
        row["Reparable"],
        diagnosis
      ),

    noRepairReasons:
      reasons,

    repairs,

    observations:
      stripDeliveryMarker(row["Observaciones"]),

    delivered: /(^|\s|·)ENTREGADO(\s|·|$)/i.test(clean(row["Observaciones"])) || Boolean(normalizeImportedDate(row["Fecha entrega"])),
    deliveryStatus: /(^|\s|·)ENTREGADO(\s|·|$)/i.test(clean(row["Observaciones"])) || Boolean(normalizeImportedDate(row["Fecha entrega"])) ? "Entregada" : "Pendiente",
    deliveryDate: normalizeImportedDate(row["Fecha entrega"]),
    warrantyMonths: Number((clean(row["Garantía"]).match(/\d+/) || [2])[0]),
    warrantyUntil: normalizeImportedDate(row["Garantía hasta"]),

    packaging:
      clean(row["Empaque"]) ||
      "Listo para empacar",

    lotId:
      lot.id,

    createdAt:
      new Date().toISOString(),

    updatedAt:
      new Date().toISOString(),

    importedFromExcel: true
  };
}

/**
 * Importa un Excel y devuelve los datos preparados.
 *
 * Esta función NO escribe directamente en IndexedDB.
 * app.js se encarga de crear el lote y guardar los registros.
 */
export async function readExcelFile(file) {
  if (!window.XLSX) {
    throw new Error(
      "La biblioteca de Excel todavía no está disponible."
    );
  }

  if (!file) {
    throw new Error(
      "No se seleccionó ningún archivo."
    );
  }

  const valid =
    /\.(xlsx|xls)$/i.test(
      file.name
    );

  if (!valid) {
    throw new Error(
      "Selecciona un archivo Excel (.xlsx o .xls)."
    );
  }

  const buffer =
    await file.arrayBuffer();

  const workbook =
    XLSX.read(buffer, {
      type: "array"
    });

  if (
    !workbook.SheetNames.length
  ) {
    throw new Error(
      "El archivo Excel no contiene hojas."
    );
  }

  /*
    Utilizamos la primera hoja.
    Es la hoja "Inventario" que genera
    Conteo Rápido.
  */
  const sheet =
    workbook.Sheets[
      workbook.SheetNames[0]
    ];

  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false, raw: false });
  const headerIndex = matrix.findIndex(row => row.some(cell => clean(cell).toLowerCase() === "id"));
  if (headerIndex < 0) throw new Error('No se encontró una fila de encabezados con la columna "ID".');
  const headers = matrix[headerIndex].map(clean);
  const rows = matrix.slice(headerIndex + 1)
    .filter(row => row.some(value => clean(value) !== ""))
    .map(row => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ""])));

  if (!rows.length) {
    throw new Error(
      "El Excel está vacío."
    );
  }

  const requiredColumns = [
    "ID"
  ];

  const columns =
    Object.keys(rows[0]);

  for (const column of requiredColumns) {
    if (!columns.includes(column)) {
      throw new Error(
        `El Excel no tiene la columna "${column}".`
      );
    }
  }

  return {
    fileName: file.name,
    sheetName:
      workbook.SheetNames[0],
    rows
  };
}

export function exportCsv(u) {
  const d = rows(u);

  const k =
    Object.keys(
      d[0] || {
        "N°": ""
      }
    );

  const e = value =>
    String(value ?? "")
      .replaceAll('"', '""');

  const c =
    "\ufeff" +
    k
      .map(
        x =>
          '"' +
          e(x) +
          '"'
      )
      .join(",") +
    "\n" +
    d
      .map(r =>
        k
          .map(
            x =>
              '"' +
              e(r[x]) +
              '"'
          )
          .join(",")
      )
      .join("\n");

  download(
    new Blob(
      [c],
      {
        type:
          "text/csv;charset=utf-8"
      }
    ),
    "conteo-rapido.csv"
  );
}

export function exportJson(x) {
  download(
    new Blob(
      [
        JSON.stringify(
          x,
          null,
          2
        )
      ],
      {
        type:
          "application/json"
      }
    ),
    "conteo-rapido-backup.json"
  );
}

function download(blob, name) {
  const a =
    document.createElement("a");

  a.href =
    URL.createObjectURL(blob);

  a.download = name;

  a.click();

  setTimeout(
    () =>
      URL.revokeObjectURL(
        a.href
      ),
    1000
  );
}