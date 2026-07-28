/**
 * Generador de archivos Excel finales para RRHH usando ExcelJS.
 * Usa los templates de ejemplo como base y solo rellena datos.
 * Si no encuentra los templates, construye desde cero con formato exacto.
 */
const Generator = {
  _templatesLoaded: false,
  _templateAusenciasWb: null,
  _templateTrasnochesWb: null,

  COLORS: {
    red: "FFC00000",
    yellow: "FFFFFF00",
    black: "FF000000",
    white: "FFFFFFFF",
    none: "00000000",
  },

  /**
   * Carga los templates. Solo se llama una vez.
   */
  async loadTemplates() {
    if (this._templatesLoaded) return;
    // En file:// no se puede hacer fetch (CORS), usar fallback directo
    if (window.location.protocol === "file:") {
      this._templatesLoaded = true;
      return;
    }
    try {
      const [ausBuf, trasBuf] = await Promise.all([
        this.fetchTemplate("Planilla de Ausencias Julio 2026.xlsx"),
        this.fetchTemplate("Trasnoches AVEL JUL 26.xlsx"),
      ]);
      const wb1 = new ExcelJS.Workbook();
      await wb1.xlsx.load(ausBuf);
      this._templateAusenciasWb = wb1;

      const wb2 = new ExcelJS.Workbook();
      await wb2.xlsx.load(trasBuf);
      this._templateTrasnochesWb = wb2;
    } catch (e) {
      // Silencioso: usar fallback manual
    }
    this._templatesLoaded = true;
  },

  async fetchTemplate(filename) {
    // Intenta cargar desde ejemplos/ o desde la misma carpeta
    const paths = [`ejemplos/${filename}`, filename];
    for (const p of paths) {
      try {
        const resp = await fetch(p);
        if (resp.ok) return await resp.arrayBuffer();
      } catch (_) {}
    }
    throw new Error(`Template ${filename} not found`);
  },

  // ═══════════════════════════════════════════════════
  // Planilla de Ausencias
  // ═══════════════════════════════════════════════════

  async generarPlanillaAusencias(data) {
    await this.loadTemplates();

    if (this._templateAusenciasWb) {
      return await this.fillTemplateAusencias(data);
    }
    return await this.buildFromScratchAusencias(data);
  },

  async fillTemplateAusencias(data) {
    const { novedades, horasExtras, puntualidad, cuartaJornada, mes, anio } = data;
    // Clonar workbook recreando desde el buffer original
    const origBuf = await this._templateAusenciasWb.xlsx.writeBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(origBuf);

    // ── Novedades ──
    const wsNov = wb.getWorksheet("Novedades");
    if (wsNov) {
      // Limpiar filas de datos (desde fila 9 en adelante)
      this.clearDataRows(wsNov, 9);
      // Rellenar novedades
      if (novedades && novedades.length > 0) {
        const dataRow = wsNov.getRow(9);
        novedades.forEach((n, idx) => {
          const r = 9 + idx;
          this.copyRowStyle(dataRow, wsNov.getRow(r));
          this.setVal(wsNov, r, 1, n.legajo || "");
          this.setVal(wsNov, r, 2, n.nombre || "");
          this.setVal(wsNov, r, 3, n.fecha ? new Date(n.fecha + "T00:00:00") : "");
          this.setVal(wsNov, r, 5, n.conAviso ? "x" : "");
          this.setVal(wsNov, r, 8, n.motivo || "");
          this.setVal(wsNov, r, 10, n.certificado ? "x" : "");
          this.setVal(wsNov, r, 15, n.sancion || "");
        });
      }
    }

    // ── 4ta Jornada ──
    const ws4ta = wb.getWorksheet("4ta jornada");
    if (ws4ta) {
      this.clearDataRows(ws4ta, 5);
      const fechaMes = `${this.nombreMes(mes)} ${anio}`;
      ws4ta.getCell("A2").value = fechaMes;
      if (cuartaJornada && cuartaJornada.length > 0) {
        cuartaJornada.forEach((item, idx) => {
          const r = 5 + idx;
          this.setVal(ws4ta, r, 1, item.legajo || "");
          this.setVal(ws4ta, r, 2, item.nombre || "");
          this.setVal(ws4ta, r, 4, item.cantidad || 1);
          this.setVal(ws4ta, r, 5, item.motivo || "");
        });
      }
    }

    // ── Horas Extras ──
    const wsHE = wb.getWorksheet("Horas Extras");
    if (wsHE) {
      this.clearDataRows(wsHE, 4);
      wsHE.getCell("B1").value = this.nombreMes(mes);
      wsHE.getCell("C1").value = anio;
      if (horasExtras && horasExtras.length > 0) {
        horasExtras.forEach((h, idx) => {
          const r = 4 + idx;
          this.setVal(wsHE, r, 1, h.legajo || "");
          this.setVal(wsHE, r, 2, h.nombre || "");
          this.setVal(wsHE, r, 3, h.fecha ? new Date(h.fecha + "T00:00:00") : "");
          this.setVal(wsHE, r, 4, this.timeToExcel(h.desde));
          this.setVal(wsHE, r, 5, "Hs");
          this.setVal(wsHE, r, 6, this.timeToExcel(h.hasta));
          this.setVal(wsHE, r, 7, "Hs");
          this.setVal(wsHE, r, 8, this.timeToExcel(h.heDesde));
          this.setVal(wsHE, r, 9, "Hs");
          this.setVal(wsHE, r, 10, this.timeToExcel(h.heHasta));
          this.setVal(wsHE, r, 11, "Hs");
          this.setVal(wsHE, r, 12, parseInt(h.cantidad) || 2);
        });
      }
    }

    // ── Puntualidad ──
    const wsPunt = wb.getWorksheet("Puntualidad");
    if (wsPunt) {
      this.clearDataRows(wsPunt, 6);
      if (puntualidad && puntualidad.length > 0) {
        puntualidad.forEach((p, idx) => {
          const r = 6 + idx;
          this.setVal(wsPunt, r, 1, p.legajo || "");
          this.setVal(wsPunt, r, 2, p.nombre || "");
        });
      }
    }

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  },

  // ═══════════════════════════════════════════════════
  // Trasnoches AVEL
  // ═══════════════════════════════════════════════════

  async generarTrasnoches(data) {
    await this.loadTemplates();

    if (this._templateTrasnochesWb) {
      return await this.fillTemplateTrasnoches(data);
    }
    return await this.buildFromScratchTrasnoches(data);
  },

  async fillTemplateTrasnoches(data) {
    const { empleados, viernes, sabados, mes, anio } = data;

    const origBuf = await this._templateTrasnochesWb.xlsx.writeBuffer();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(origBuf);

    const ws = wb.getWorksheet("TRASNOCHES");
    if (!ws) {
      const fallback = await this.buildFromScratchTrasnoches(data);
      return fallback;
    }

    // Limpiar datos de empleados (filas 4 a 67)
    this.clearDataRows(ws, 4, 67);

    // Actualizar fechas en fila 2 (base date en C2, luego fórmulas)
    if (viernes.length > 0) {
      const baseCol = 3;
      ws.getCell(2, baseCol).value = viernes[0] ? new Date(viernes[0] + "T00:00:00") : null;
    }

    // Rellenar datos por empleado
    empleados.forEach((emp, idx) => {
      const r = 4 + idx;
      if (r > ws.rowCount) return;
      this.setVal(ws, r, 1, emp.legajo || "");
      this.setVal(ws, r, 2, emp.nombre || "");

      let totalDias = 0;
      for (let i = 0; i < viernes.length; i++) {
        const baseCol = 3 + i * 6;
        const v = emp.viernes && emp.viernes[i] ? emp.viernes[i] : null;
        const s = emp.sabados && emp.sabados[i] ? emp.sabados[i] : null;

        this.setVal(ws, r, baseCol, v && v.horas > 0 ? "si" : "");
        this.setVal(ws, r, baseCol + 1, v && v.horas > 0 && v.salida ? this.timeToExcel(v.salida) : this.timeToExcel("00:00:00"));
        if (v && v.horas > 0 && baseCol + 2 <= 100) {
          this.setVal(ws, r, baseCol + 2, this.timeToExcel(this.horasToTimeStr(v.horas)));
          totalDias++;
        }
        this.setVal(ws, r, baseCol + 3, s && s.horas > 0 ? "si" : "");
        this.setVal(ws, r, baseCol + 4, s && s.horas > 0 && s.salida ? this.timeToExcel(s.salida) : this.timeToExcel("00:00:00"));
        if (s && s.horas > 0 && baseCol + 5 <= 100) {
          this.setVal(ws, r, baseCol + 5, this.timeToExcel(this.horasToTimeStr(s.horas)));
          totalDias++;
        }
      }
      // DIAS column
      const diasCol = 3 + viernes.length * 6;
      this.setVal(ws, r, diasCol, totalDias);
    });

    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  },

  // ═══════════════════════════════════════════════════
  // Helpers para template
  // ═══════════════════════════════════════════════════

  clearDataRows(ws, startRow, endRow) {
    const end = endRow || ws.rowCount;
    for (let r = startRow; r <= end; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= (ws.columnCount || 50); c++) {
        const cell = row.getCell(c);
        cell.value = null;
      }
    }
  },

  setVal(ws, r, c, val) {
    if (val === null || val === undefined || val === "") return;
    ws.getCell(r, c).value = val;
  },

  copyRowStyle(srcRow, destRow) {
    // ExcelJS preserves styles when loading template, so cells already have styles
    // We just need to clear old data and set new values
  },

  // ═══════════════════════════════════════════════════
  // Fallback: build from scratch (misma lógica que antes)
  // ═══════════════════════════════════════════════════

  async buildFromScratchAusencias(data) {
    const wb = new ExcelJS.Workbook();
    const { novedades, horasExtras, puntualidad, cuartaJornada, mes, anio } = data;
    this._buildSheetNovedades(wb, novedades, mes, anio);
    this._buildSheet4taJornada(wb, cuartaJornada, mes, anio);
    this._buildSheetHorasExtras(wb, horasExtras, mes, anio);
    this._buildSheetPuntualidad(wb, puntualidad);
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  },

  async buildFromScratchTrasnoches(data) {
    const wb = new ExcelJS.Workbook();
    this._buildSheetTrasnoches(wb, data.empleados, data.viernes, data.sabados, data.mes, data.anio);
    const buf = await wb.xlsx.writeBuffer();
    return new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  },

  // ── Fallback: font helpers ──
  fontArial(size, bold) { return { name: "Arial", size, bold, color: { argb: this.COLORS.black } }; },
  fontArialNarrow(size, bold) { return { name: "Arial Narrow", size, bold, color: { argb: this.COLORS.black } }; },
  fontCalibri(size, bold, white) { return { name: "Calibri", size, bold, color: { argb: white ? this.COLORS.white : this.COLORS.black } }; },
  fillSolid(argb) { return { type: "pattern", pattern: "solid", fgColor: { argb } }; },
  alignC() { return { vertical: "middle", horizontal: "center" }; },
  alignCW() { return { vertical: "middle", horizontal: "center", wrapText: true }; },
  alignLW() { return { vertical: "middle", horizontal: "left", wrapText: true }; },
  borderAll(s) { const b = { style: s }; return { top: b, bottom: b, left: b, right: b }; },
  borderM() { return this.borderAll("medium"); },
  borderT() { return this.borderAll("thin"); },
  borders(t, b, l, r) { const o = {}; if (t) o.top = { style: t }; if (b) o.bottom = { style: b }; if (l) o.left = { style: l }; if (r) o.right = { style: r }; return o; },

  style(ws, r, c, val, font, fill, alignment, border, numFmt) {
    const cell = ws.getCell(r, c);
    cell.value = val;
    if (font) cell.font = font;
    if (fill) cell.fill = fill;
    if (alignment) cell.alignment = alignment;
    if (border) cell.border = border;
    if (numFmt) cell.numFmt = numFmt;
    return cell;
  },

  _buildSheetNovedades(wb, novedades, mes, anio) {
    const ws = wb.addWorksheet("Novedades");
    const F = this.fontArial.bind(this);
    const widths = [10.86, 24, 43.14, 0.71, 9, 8.43, 0.71, 48.29, 0.71, 13.14, 13.86, 0.71, 2, 8, 3, 20, 5];
    widths.forEach((w, i) => ws.getColumn(i + 1).width = w);

    this.style(ws, 3, 1, "Planilla de Novedades - Ausencias", F(12, true), null, this.alignC(), this.borders("medium", "medium", "medium", null), "mm-dd-yy");
    ws.mergeCells("A3:C3");
    this.style(ws, 3, 8, "Avellaneda", F(12, true), null, this.alignC(), this.borderM(), "mm-dd-yy");

    const fechaMes = new Date(anio, mes - 1, 1);
    this.style(ws, 4, 1, fechaMes, F(12, true), null, this.alignC(), this.borderM(), "mmm-yy");
    ws.mergeCells("A4:Q4");

    // Row 7 headers
    this.style(ws, 7, 1, "Datos del Empleado", F(9, true), null, this.alignC(), this.borderM());
    ws.mergeCells("A7:C7");
    this.style(ws, 7, 5, "Ausente", F(9, true), null, this.alignCW(), this.borders("medium", "medium", "medium", "thin"));
    ws.mergeCells("E7:F7");
    this.style(ws, 7, 8, "Ausencia por enf. / examen / vs,", F(9, true), null, this.alignCW(), this.borders("medium", "thin", "medium", "medium"));
    this.style(ws, 7, 10, "Entrega de Certificado", F(9, true), null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    ws.mergeCells("J7:K7");
    this.style(ws, 7, 13, "Llegada tarde en minutos", F(9, true), null, this.alignCW(), this.borders("medium", "thin", "medium", "medium"));
    ws.mergeCells("M7:N7");
    this.style(ws, 7, 15, "Tipo de sanción a imponer", F(9, true), null, this.alignCW(), this.borders("medium", "medium", "medium", null));
    this.style(ws, 7, 16, "2hs", F(11, true), this.fillSolid("FFFF0000"), this.alignC(), this.borders("medium", null, "medium", null));

    // Row 8 sub-headers
    this.style(ws, 8, 1, "Nº Leg.", F(9, true), null, this.alignC(), this.borders("medium", "medium", "medium", null));
    this.style(ws, 8, 2, "Nombre y Apellido", F(9, true), null, this.alignC(), this.borders("medium", "medium", "medium", "medium"));
    this.style(ws, 8, 3, "Día", F(9, true), null, this.alignC(), this.borders("medium", "medium", "medium", "medium"));
    this.style(ws, 8, 4, " ", null, null, { vertical: "middle" }, this.borders(null, "medium", null, null));
    this.style(ws, 8, 5, "con aviso", F(9, true), null, this.alignCW(), this.borders("medium", "medium", "medium", "medium"));
    this.style(ws, 8, 6, "Sin Aviso", F(9, true), null, this.alignCW(), this.borders(null, "medium", null, "medium"));
    this.style(ws, 8, 7, " ", null, null, { vertical: "middle" }, this.borders(null, "medium", null, null));
    this.style(ws, 8, 10, "SI", F(9, true), null, this.alignC(), this.borders(null, "medium", "medium", null));
    this.style(ws, 8, 11, "NO", F(10, true), null, this.alignC(), this.borders(null, "medium", null, "medium"));
    this.style(ws, 8, 14, " ", F(9, true), null, { vertical: "middle" }, this.borders(null, "medium", null, null));
    this.style(ws, 8, 16, "De aviso", F(14, true), this.fillSolid("FFFF0000"), this.alignC(), this.borders(null, "thin", "medium", null));

    // Data rows
    if (novedades && novedades.length > 0) {
      novedades.forEach((n, idx) => {
        const r = 9 + idx;
        ws.getRow(r).height = 35.25;
        this.style(ws, r, 1, n.legajo || "", F(9, true), null, this.alignC(), this.borderT());
        this.style(ws, r, 2, n.nombre || "", F(10, true), null, this.alignC(), this.borderT());
        this.style(ws, r, 3, n.fecha ? new Date(n.fecha + "T00:00:00") : "", F(10, true), null, this.alignC(), this.borders(null, "thin", "thin", "thin"), "mm-dd-yy");
        this.style(ws, r, 5, n.conAviso ? "x" : "", F(9, true), null, this.alignCW(), this.borders(null, "thin", "thin", "thin"));
        this.style(ws, r, 8, n.motivo || "", F(10, true), null, this.alignCW(), this.borderT());
        this.style(ws, r, 10, n.certificado ? "x" : "", F(9, true), null, this.alignC(), this.borders(null, "thin", "thin", "thin"));
        this.style(ws, r, 15, n.sancion || "", F(10, true), null, this.alignCW(), this.borders(null, "thin", "medium", null));
      });
    }
  },

  _buildSheet4taJornada(wb, cuartaJornada, mes, anio) {
    const ws = wb.addWorksheet("4ta jornada");
    ws.getColumn(1).width = 26;
    ws.getColumn(2).width = 33;
    ws.getColumn(3).width = 3.29;
    ws.getColumn(4).width = 21.43;
    ws.getColumn(5).width = 40.14;

    this.style(ws, 1, 1, "4tos Días", this.fontCalibri(14, true, true), this.fillSolid(this.COLORS.red), this.alignC(), this.borders(null, null, "medium", null));
    ws.mergeCells("A1:E1");
    this.style(ws, 2, 1, `${this.nombreMes(mes)} ${anio}`, this.fontCalibri(14, true), null, this.alignC(), this.borders(null, "medium", "medium", null));
    ws.mergeCells("A2:E2");

    const hdrF = this.fontCalibri(12, false);
    const hdrB = this.borders("medium", null, "medium", "medium");
    this.style(ws, 3, 1, "Legajo", hdrF, null, this.alignC(), hdrB);
    this.style(ws, 3, 2, "Nombre y Apellido", hdrF, null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 3, 4, "Cantidad de 4tos días", hdrF, null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 3, 5, "Motivo", hdrF, null, this.alignC(), this.borders("medium", null, "medium", "medium"));

    if (cuartaJornada && cuartaJornada.length > 0) {
      cuartaJornada.forEach((item, idx) => {
        const r = 5 + idx;
        this.style(ws, r, 1, item.legajo || "", null, null, this.alignC());
        this.style(ws, r, 2, item.nombre || "", null, null, this.alignC());
        this.style(ws, r, 4, item.cantidad || 1, null, null, this.alignC());
        this.style(ws, r, 5, item.motivo || "", null, null, this.alignC());
      });
    }
  },

  _buildSheetHorasExtras(wb, horasExtras, mes, anio) {
    const ws = wb.addWorksheet("Horas Extras");
    const F = this.fontArial.bind(this);
    const FN = this.fontArialNarrow.bind(this);
    [11.14, 34.14, 31.71, 10.71, 4.14, 11.43, 4.14, 10.71, 4.14, 10.71, 4.14, 10.71].forEach((w, i) => ws.getColumn(i + 1).width = w);

    const hb = F(10, true);
    this.style(ws, 1, 1, "PERIODO", hb, null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 1, 2, this.nombreMes(mes), hb, null, this.alignC(), this.borders("medium", null, null, "medium"));
    this.style(ws, 1, 3, anio, hb, null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 1, 4, "Jornada Laboral", hb, null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 1, 8, "Cantidad de Horas Extras", hb, null, this.alignC(), this.borders("medium", null, null, "medium"));
    ws.mergeCells("A1:A2"); ws.mergeCells("B1:B2"); ws.mergeCells("C1:C2");
    ws.mergeCells("D1:G2"); ws.mergeCells("H1:N2");

    this.style(ws, 3, 1, "Legajo", hb, null, this.alignC(), this.borderM());
    this.style(ws, 3, 2, "Apellidos y Nombres Completos", hb, null, this.alignC(), this.borders("medium", "medium", null, "medium"));
    this.style(ws, 3, 3, "Fecha", hb, null, this.alignC(), this.borderM());
    this.style(ws, 3, 4, "Desde", hb, null, this.alignC(), this.borderM()); ws.mergeCells("D3:E3");
    this.style(ws, 3, 6, "Hasta", hb, null, this.alignC(), this.borderM()); ws.mergeCells("F3:G3");
    this.style(ws, 3, 8, "Desde", hb, null, this.alignC(), this.borderM()); ws.mergeCells("H3:I3");
    this.style(ws, 3, 10, "Hasta", hb, null, this.alignC(), this.borderM()); ws.mergeCells("J3:K3");
    this.style(ws, 3, 12, "Cantidad", hb, null, this.alignC(), this.borders("medium", "medium", null, "medium"));
    this.style(ws, 3, 13, 0.5, hb, null, this.alignC(), this.borders("medium", "medium", null, "medium"), "0%");
    this.style(ws, 3, 14, 1, hb, null, this.alignC(), this.borderM(), "0%");

    if (horasExtras && horasExtras.length > 0) {
      const lf = FN(10, true), nf = FN(10, false), tf = F(10, false), hs = F(10, true);
      horasExtras.forEach((h, idx) => {
        const r = 4 + idx;
        this.style(ws, r, 1, h.legajo || "", lf, null, this.alignCW(), this.borders("thin", "thin", "medium", "medium"));
        this.style(ws, r, 2, h.nombre || "", nf, null, this.alignLW(), this.borders("thin", "thin", null, "medium"));
        this.style(ws, r, 3, h.fecha ? new Date(h.fecha + "T00:00:00") : "", tf, null, this.alignC(), this.borders(null, "thin", "medium", "medium"), "[$-F800]dddd, mmmm dd, yyyy");
        this.style(ws, r, 4, this.timeToExcel(h.desde), tf, null, this.alignC(), this.borders(null, "thin", "medium", null), "h:mm");
        this.style(ws, r, 5, "Hs", hs, null, this.alignC(), this.borders(null, "thin", "medium", "medium"));
        this.style(ws, r, 6, this.timeToExcel(h.hasta), tf, null, this.alignC(), this.borders(null, "thin", null, null), "h:mm");
        this.style(ws, r, 7, "Hs", hs, null, this.alignC(), this.borders(null, "thin", "medium", "medium"));
        this.style(ws, r, 8, this.timeToExcel(h.heDesde), tf, null, this.alignC(), this.borders(null, "thin", null, null), "h:mm");
        this.style(ws, r, 9, "Hs", hs, null, this.alignC(), this.borders(null, "thin", "medium", "medium"));
        this.style(ws, r, 10, this.timeToExcel(h.heHasta), tf, null, this.alignC(), this.borders(null, "thin", null, null), "h:mm");
        this.style(ws, r, 11, "Hs", hs, null, this.alignC(), this.borders(null, "thin", "medium", "medium"));
        this.style(ws, r, 12, parseInt(h.cantidad) || 2, hs, null, this.alignC(), this.borders(null, "thin", null, "medium"));
      });
    }
  },

  _buildSheetPuntualidad(wb, puntualidad) {
    const ws = wb.addWorksheet("Puntualidad");
    const FN = this.fontArialNarrow.bind(this);
    ws.getColumn(1).width = 15.86;
    ws.getColumn(2).width = 30.43;

    this.style(ws, 2, 1, "CINEPOLIS  -  COMPLEJO  AVELLANEDA", FN(14, true), null, this.alignC(), this.borderM());
    ws.mergeCells("A2:B2");
    this.style(ws, 4, 1, "PERDIERON LA PUNTUALIDAD", FN(14, true), null, this.alignC(), this.borderM());
    ws.mergeCells("A4:B4");
    this.style(ws, 5, 1, "Leg.", FN(14, true), null, this.alignC(), this.borders("medium", null, "medium", "medium"));
    this.style(ws, 5, 2, "Apellidos y Nombres", FN(14, true), null, this.alignC(), this.borders("medium", null, null, "medium"));

    if (puntualidad && puntualidad.length > 0) {
      puntualidad.forEach((p, idx) => {
        const r = 6 + idx;
        this.style(ws, r, 1, p.legajo || "", FN(10, true), null, this.alignC(), this.borderT());
        this.style(ws, r, 2, p.nombre || "", FN(10, false), null, this.alignC(), this.borderT());
      });
    }
  },

  _buildSheetTrasnoches(wb, empleados, viernes, sabados, mes, anio) {
    const ws = wb.addWorksheet("TRASNOCHES");
    const F = this.fontArial.bind(this);
    const FN = this.fontArialNarrow.bind(this);

    ws.getColumn(1).width = 10;
    ws.getColumn(2).width = 48.57;

    const hb = F(10, true);
    this.style(ws, 1, 1, "AVELLANEDA", hb, null, this.alignC(), this.borderM());
    ws.mergeCells("A1:B2");

    const numFins = viernes.length;
    const totalCol = 3 + numFins * 6;

    for (let i = 0; i < numFins; i++) {
      const bc = 3 + i * 6;
      this.style(ws, 1, bc, "VIERNES", hb, null, null, this.borders("medium", "medium", "medium", null));
      this.style(ws, 1, bc + 3, "SABADO", hb, null, null, this.borders("medium", "medium", "medium", null));
      const vDate = viernes[i] ? new Date(viernes[i] + "T00:00:00") : null;
      const sDate = sabados[i] ? new Date(sabados[i] + "T00:00:00") : null;
      this.style(ws, 2, bc, vDate, hb, this.fillSolid(this.COLORS.yellow), null, this.borders("medium", "medium", "medium", null), "d-mmm");
      this.style(ws, 2, bc + 3, sDate, hb, this.fillSolid(this.COLORS.yellow), null, this.borders("medium", "medium", "medium", null), "d-mmm");
    }

    ws.mergeCells(`${this.colLetter(totalCol)}1:${this.colLetter(totalCol)}2`);
    this.style(ws, 1, totalCol, "TOTALES", hb, null, this.alignC(), this.borderM());

    this.style(ws, 3, 1, "Leg.", hb, null, this.alignC(), this.borderM());
    this.style(ws, 3, 2, "Empleado", hb, null, this.alignC(), this.borders("medium", "medium", null, "medium"));

    for (let i = 0; i < numFins; i++) {
      const bc = 3 + i * 6;
      this.style(ws, 3, bc, "si", hb, null, this.alignC(), this.borders(null, "medium", null, "thin"));
      this.style(ws, 3, bc + 1, "Salida", hb, null, this.alignC(), this.borders(null, "medium", "thin", null));
      this.style(ws, 3, bc + 2, "Trasnoche", hb, null, this.alignC(), this.borders(null, "medium", "medium", "medium"));
      this.style(ws, 3, bc + 3, "si", hb, null, this.alignC(), this.borders(null, "medium", null, "thin"));
      this.style(ws, 3, bc + 4, "Salida", hb, null, this.alignC(), this.borders(null, "medium", "thin", null));
      this.style(ws, 3, bc + 5, "Trasnoche", hb, null, this.alignC(), this.borders(null, "medium", "medium", "medium"));
    }

    // DIAS column
    const diasCol = totalCol;
    this.style(ws, 3, diasCol, "DIAS", hb, null, this.alignC(), this.borderM());
    ws.getColumn(diasCol).width = 7;

    for (let i = 0; i < numFins; i++) {
      const bc = 3 + i * 6;
      [bc, bc + 1, bc + 2, bc + 3, bc + 4, bc + 5].forEach((col, j) => {
        ws.getColumn(col).width = j % 3 === 1 ? 9.57 : 9.71;
      });
    }

    const dt = this.timeToExcel("00:00:00");
    const tf = F(10, false);

    empleados.forEach((emp, idx) => {
      const r = 4 + idx;
      this.style(ws, r, 1, emp.legajo || "", FN(10, true), null, this.alignCW(), this.borders("thin", "thin", "medium", "medium"));
      this.style(ws, r, 2, emp.nombre || "", FN(10, false), null, this.alignLW(), this.borders("thin", "thin", null, "medium"));

      let totalDias = 0;
      for (let i = 0; i < numFins; i++) {
        const bc = 3 + i * 6;
        const v = emp.viernes && emp.viernes[i] ? emp.viernes[i] : null;
        const s = emp.sabados && emp.sabados[i] ? emp.sabados[i] : null;

        this.style(ws, r, bc, v && v.horas > 0 ? "si" : "", null, null, this.alignC(), this.borders(null, "thin", null, "thin"));
        this.style(ws, r, bc + 1, v && v.horas > 0 && v.salida ? this.timeToExcel(v.salida) : dt, tf, null, this.alignC(), this.borders(null, "thin", "thin", null), "h:mm");
        if (v && v.horas > 0) {
          this.style(ws, r, bc + 2, this.timeToExcel(this.horasToTimeStr(v.horas)), tf, null, this.alignC(), this.borders(null, "thin", "medium", "medium"), "h:mm");
          totalDias++;
        } else {
          this.style(ws, r, bc + 2, dt, tf, null, this.alignC(), this.borders(null, "thin", "medium", "medium"), "h:mm");
        }

        this.style(ws, r, bc + 3, s && s.horas > 0 ? "si" : "", null, null, this.alignC(), this.borders(null, "thin", null, "thin"));
        this.style(ws, r, bc + 4, s && s.horas > 0 && s.salida ? this.timeToExcel(s.salida) : dt, tf, null, this.alignC(), this.borders(null, "thin", "thin", null), "h:mm");
        if (s && s.horas > 0) {
          this.style(ws, r, bc + 5, this.timeToExcel(this.horasToTimeStr(s.horas)), tf, null, this.alignC(), this.borders(null, "thin", "medium", "medium"), "h:mm");
          totalDias++;
        } else {
          this.style(ws, r, bc + 5, dt, tf, null, this.alignC(), this.borders(null, "thin", "medium", "medium"), "h:mm");
        }
      }
      this.style(ws, r, diasCol, totalDias, FN(10, true), null, this.alignC());
    });

    // Total row
    const totalRow = empleados.length + 4;
    this.style(ws, totalRow, 1, empleados.length, hb, null, this.alignC());
    this.style(ws, totalRow, 2, "TRASNOCHES POR DIA", hb);

    for (let i = 0; i < numFins; i++) {
      const bc = 3 + i * 6;
      let vc = 0, sc = 0;
      for (const emp of empleados) {
        if (emp.viernes && emp.viernes[i] && emp.viernes[i].horas > 0) vc++;
        if (emp.sabados && emp.sabados[i] && emp.sabados[i].horas > 0) sc++;
      }
      this.style(ws, totalRow, bc, vc, null, null, this.alignC());
      this.style(ws, totalRow, bc + 3, sc, null, null, this.alignC());
    }

    // Instructions
    const instr = [
      "Cargar solamente los datos de los turnos de las personas que tuvieron trasnoche.",
      "Si los empleados están dados de baja, no borren la fila, déjenla con datos en 0 para que no se rompan las fórmulas.",
      "Si hay empleados que no están porque son nuevos ingresos, agregarlos al final.",
    ];
    instr.forEach((txt, i) => {
      const r = totalRow + 2 + i;
      this.style(ws, r, 1, txt, F(10, true));
      ws.mergeCells(`${this.colLetter(1)}${r}:${this.colLetter(totalCol - 1)}${r}`);
    });
  },

  // ═══════════════════════════════════════════════════
  // Helpers
  // ═══════════════════════════════════════════════════

  nombreMes(m) {
    const meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
    return meses[m - 1] || "";
  },

  nombreMesAbbr(m) {
    return ["ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC"][m - 1] || "";
  },

  horasToTimeStr(h) {
    const hh = Math.floor(h);
    const mm = Math.round((h - hh) * 60);
    return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}:00`;
  },

  timeToExcel(ts) {
    if (!ts) return null;
    const p = String(ts).split(":");
    return (parseInt(p[0]) * 60 + parseInt(p[1])) / 1440;
  },

  colLetter(n) {
    let s = "";
    while (n > 0) { n--; s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26); }
    return s;
  },
};
