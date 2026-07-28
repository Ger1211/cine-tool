/**
 * Parser del archivo Roster (Excel con múltiples pestañas)
 * Usa ExcelJS para leer y procesar el XLSX.
 */
const RosterParser = {
  /**
   * Parsea el archivo XLSX del roster
   * @param {ArrayBuffer} data - Datos del archivo
   * @param {string} targetMonth - Mes objetivo "YYYY-MM"
   * @returns {Array} Array de empleados {nombre, tipo, dias:[...]}
   */
  async parse(data, targetMonth) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);

    const sectoresIgnorar = [
      "pedidos de horario",
      "jornadas reducci",
      "hoja1",
      "hoja2",
      "vacaciones",
    ];

    const sheetsToProcess = wb.worksheets.filter((ws) => {
      const lower = ws.name.toLowerCase();
      return !sectoresIgnorar.some((s) => lower.includes(s));
    });

    const all = [];
    for (const ws of sheetsToProcess) {
      if (ws.rowCount < 5) continue;
      const parsed = this.parseSheet(ws, targetMonth);
      all.push(...parsed);
    }

    return this.mergeEmpleados(all);
  },

  /**
   * Parsea una hoja del roster usando ExcelJS
   */
  parseSheet(ws, targetMonth) {
    const weeks = this.findWeekBlocks(ws);
    if (weeks.length === 0) return [];

    const resultado = [];

    for (const week of weeks) {
      const { empStartRow, empEndRow, dateCols } = week;

      for (let r = empStartRow; r <= empEndRow; r++) {
        const row = ws.getRow(r);
        const nombre = this.normalizeName(this.getCellValue(row, 1));
        const tipo = String(this.getCellValue(row, 2) || "").trim().toUpperCase();
        if (!nombre || (tipo !== "FT" && tipo !== "PT")) continue;

        const empleado = { nombre, tipo, dias: [] };

        for (const { col_entry, col_exit, date } of dateCols) {
          if (!date) continue;

          const dateStr = this.formatDate(date);
          if (targetMonth && !dateStr.startsWith(targetMonth)) continue;

          const entry = this.getCellValue(row, col_entry);
          const exit = this.getCellValue(row, col_exit);

          const entryStr = this.cellToString(entry);
          const exitStr = this.cellToString(exit);

          if (!entryStr && !exitStr) continue;

          const isFranco = /^FR\d*$/i.test(entryStr);
          const isLicencia = entryStr.toUpperCase() === "LIC";
          const isVacaciones = entryStr.toUpperCase() === "VAC";

          empleado.dias.push({
            fecha: dateStr,
            entrada: isFranco || isLicencia || isVacaciones ? null : (this.isTime(entryStr) ? entryStr : null),
            salida: isFranco || isLicencia || isVacaciones ? null : (this.isTime(exitStr) ? exitStr : null),
            franco: isFranco,
            licencia: isLicencia,
            vacaciones: isVacaciones,
          });
        }

        if (empleado.dias.length > 0) {
          resultado.push(empleado);
        }
      }
    }

    return resultado;
  },

  /**
   * Obtiene el valor efectivo de una celda (resuelve fórmulas)
   */
  getCellValue(row, col) {
    const cell = row.getCell(col);
    const v = cell.value;
    if (v === null || v === undefined) return null;
    // Fórmulas: ExcelJS devuelve { formula, result }
    if (typeof v === "object" && v !== null && "result" in v) {
      return v.result;
    }
    return v;
  },

  /**
   * Convierte un valor de celda a string usable como horario.
   * Las fechas reales (año > 1900) se ignoran (no son horarios).
   * Las fechas epoch (1899/1900) son valores de tiempo de Excel.
   */
  cellToString(val) {
    if (val === null || val === undefined) return "";
    if (val instanceof Date && !isNaN(val.getTime())) {
      // Fecha real (año >= 2000): no es un horario
      if (val.getFullYear() >= 2000) return "";
      // Epoch de Excel (1899/1900): es un valor de tiempo
      const h = val.getUTCHours();
      const m = val.getUTCMinutes();
      return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:00`;
    }
    if (typeof val === "number") {
      // Fracción de tiempo de Excel (ej: 0.375 = 9:00)
      if (val > 0 && val < 1) {
        const totalMinutes = Math.round(val * 1440);
        const h = Math.floor(totalMinutes / 60);
        const m = totalMinutes % 60;
        return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:00`;
      }
      return "";
    }
    return String(val).trim();
  },

  /**
   * Encuentra bloques de semanas escaneando TODAS las filas en busca de STAFF|FT/PT.
   * Los marcadores pueden estar en múltiples filas (el roster los reparte).
   */
  findWeekBlocks(ws) {
    const weekBlocks = [];
    const colCount = ws.columnCount || 80;

    // Escanear todas las filas (típicamente filas 3-5 contienen los STAFF)
    for (let rowIdx = 1; rowIdx <= Math.min(ws.rowCount, 6); rowIdx++) {
      const row = ws.getRow(rowIdx);

      for (let c = 1; c <= colCount; c++) {
        const cellVal = String(this.getCellValue(row, c) || "").trim().toUpperCase();
        if (cellVal !== "STAFF") continue;
        if (c + 1 > colCount) continue;
        const nextVal = String(this.getCellValue(row, c + 1) || "").trim().toUpperCase();
        if (nextVal !== "FT/PT" && nextVal !== "FTPT") continue;

        // STAFF|FT/PT encontrado en (rowIdx, c). Leer 7 fechas a partir de c+2.
        const dateCols = [];
        for (let dc = c + 2, day = 0; day < 7 && dc + 1 <= colCount; dc += 2, day++) {
          const rawVal = this.getCellValue(row, dc);
          const dateVal = this.parseExcelDate(rawVal);
          if (dateVal) {
            dateCols.push({ col_entry: dc, col_exit: dc + 1, date: dateVal });
          }
        }

        if (dateCols.length === 0) continue;

        let empStartRow = Math.max(rowIdx + 1, 5);
        let empEndRow = ws.rowCount;

        for (let r = empStartRow; r <= ws.rowCount; r++) {
          const nxtRow = ws.getRow(r);
          const f = String(this.getCellValue(nxtRow, 1) || "").trim().toUpperCase();
          if (f === "STAFF") { empEndRow = r - 1; break; }
        }

        if (empEndRow < empStartRow) continue;

        weekBlocks.push({ headerRow: rowIdx, empStartRow, empEndRow, dateCols });
      }
    }

    return weekBlocks;
  },

  parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) return val;
    const str = String(val).trim();
    if (!str) return null;

    const m = str.match(/(\d{2})[\/\-](\d{2})[\/\-](\d{2,4})/);
    if (m) {
      const d = parseInt(m[1], 10);
      const mo = parseInt(m[2], 10) - 1;
      const y = parseInt(m[3], 10);
      return new Date(y < 100 ? 2000 + y : y, mo, d);
    }

    return null;
  },

  formatDate(date) {
    if (!date) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  },

  isTime(str) {
    return /^\d{1,2}:\d{2}(:\d{2})?$/.test(str);
  },

  normalizeName(name) {
    if (!name) return "";
    return String(name).replace(/\s+/g, " ").trim();
  },

  mergeEmpleados(all) {
    const map = new Map();
    for (const emp of all) {
      const key = emp.nombre.toLowerCase();
      if (map.has(key)) {
        map.get(key).dias.push(...emp.dias);
      } else {
        map.set(key, { nombre: emp.nombre, tipo: emp.tipo, dias: [...emp.dias] });
      }
    }
    const result = Array.from(map.values());
    for (const emp of result) {
      emp.dias.sort((a, b) => a.fecha.localeCompare(b.fecha));
    }
    return result;
  },
};
