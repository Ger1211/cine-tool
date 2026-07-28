/**
 * Controlador principal de la aplicación
 */
(function () {
  "use strict";

  // ── Estado ──
  const state = {
    nelclockFile: null,
    rosterFile: null,
    fichadas: [],
    rosterData: [],
    resPuntualidad: null,
    resTrasnoches: null,
    res4taJornada: null,
    novedades: [],
    horasExtras: [],
    mes: 7,
    anio: 2026,
  };

  // ── Legajos (extraídos de andy ayuda rrhh.xlsx) ──
  const LEGAJOS = {
    "alfonzo, facundo carlos javier": "12636",
    "avalos, maria de los angeles": "11864",
    "ayala munoz, fernando daniel": "12631",
    "bruno, juan pablo": "11580",
    "caceres arias, victoria": "12642",
    "caceres, gabino tomas": "12142",
    "carballo, sofia itati": "12641",
    "celiz, galo valentino": "11549",
    "cerveny, kiara candela": "11852",
    "chavez imfeld, thiago joaquin": "12635",
    "de zan, nicolas alberto": "7395",
    "dip brain, natali daniela": "11578",
    "estevez, samanta michelle": "12644",
    "fernandez, gimena noemi": "8823",
    "fernandez, lucia valentina": "12526",
    "fuentes, camilo benjamin": "12643",
    "gomez, federico daniel": "8011",
    "gomez, sofia celeste": "11134",
    "gonzalez rodas, araceli abigail": "12638",
    "gonzalez, melany morena": "11045",
    "gutierrez, betiana lara": "8274",
    "herrera, alejandro ismael": "10177",
    "insaurralde, jano leonel": "12421",
    "lagos, juan manuel": "12634",
    "leguizamon, santiago valentin": "11454",
    "lescano, federico nahuel": "5847",
    "lezcano matto, dana luciana": "12645",
    "llanos, gonzalo damian": "7184",
    "lopez, tomas santino": "12633",
    "lucena, juanita ludmila": "12630",
    "mari, florencia gabriela": "1468",
    "monzon, brian maximiliano": "12640",
    "nicola almiron, sofia daniela": "11413",
    "ortiz, andres sebastian": "6176",
    "otero, gustavo damian": "8262",
    "ovejero, pamela deborah": "5240",
    "pedrozo coronel, johana ester": "11128",
    "pedrozo coronel, yohana ester": "11128",
    "ramirez contreras, marilyn marlene": "12527",
    "rodriguez, enzo gustavo": "12271",
    "ruiz, matilde carolina": "8855",
    "santana, juan cruz": "12629",
    "segovia coronel, candela morena": "12637",
    "sosa, noelia elizabeth": "12628",
    "sotelo, vanina micaela": "11125",
    "terrio, julieta macarena": "12484",
    "travaglini, matias agustin": "12639",
    "tur gimenez, carlos agustin": "11501",
    "vandecaveye, marcelo alejandro": "11046",
    "vega, barbara alejandra": "12409",
    "zaza, thiago stefano": "12524",
    "zisuela, josefina lilen": "12632",
  };

  function buscarLegajo(nombre) {
    if (!nombre) return "";
    const key = nombre.toLowerCase().trim().replace(/\s+/g, " ");
    if (LEGAJOS[key]) return LEGAJOS[key];
    // Buscar por coincidencia parcial
    const palabras = key.split(/\s+/).filter((w) => w.length > 1);
    for (const [k, v] of Object.entries(LEGAJOS)) {
      if (palabras.every((p) => k.includes(p))) return v;
    }
    // Búsqueda más laxa: al menos 2 palabras coinciden
    if (palabras.length >= 2) {
      for (const [k, v] of Object.entries(LEGAJOS)) {
        const matches = palabras.filter((p) => k.includes(p));
        if (matches.length >= 2 && matches.length >= palabras.length - 1) return v;
      }
    }
    // Coincidencia por una palabra larga y única
    if (palabras.length === 1 && palabras[0].length > 4) {
      for (const [k, v] of Object.entries(LEGAJOS)) {
        if (k.includes(palabras[0])) return v;
      }
    }
    // Fallback: generar legajo secuencial para empleados sin mapeo
    if (!buscarLegajo._fallback) buscarLegajo._fallback = {};
    if (!buscarLegajo._fallback[key]) {
      buscarLegajo._counter = (buscarLegajo._counter || 90000) + 1;
      buscarLegajo._fallback[key] = String(buscarLegajo._counter);
    }
    return buscarLegajo._fallback[key] || "";
  }

  // ── Elementos DOM ──
  const $ = (s) => document.querySelector(s);
  const $$ = (s) => document.querySelectorAll(s);

  const dom = {
    dropNelclock: $("#drop-nelclock"),
    dropRoster: $("#drop-roster"),
    inputNelclock: $("#input-nelclock"),
    inputRoster: $("#input-roster"),
    fileNelclockName: $("#file-nelclock-name"),
    fileRosterName: $("#file-roster-name"),
    btnProcesar: $("#btn-procesar"),
    selectMes: $("#select-mes"),
    inputAnio: $("#input-anio"),
    resultadosAuto: $("#resultados-auto"),
    cargaManual: $("#carga-manual"),
    descarga: $("#descarga"),
    numPerdieron: $("#num-perdieron"),
    numTrasnoches: $("#num-trasnoches"),
    num4ta: $("#num-4ta"),
    tbodyPuntualidad: $("#tbody-puntualidad"),
    tbodyTrasnoches: $("#tbody-trasnoches"),
    tbody4ta: $("#tbody-4ta"),
    novEmpleado: $("#nov-empleado"),
    heEmpleado: $("#he-empleado"),
    tbodyNovedades: $("#tbody-novedades"),
    tbodyHorasExtras: $("#tbody-horas-extras"),
    btnDescargarAusencias: $("#btn-descargar-ausencias"),
    btnDescargarTrasnoches: $("#btn-descargar-trasnoches"),
    statusDot: $("#status-dot"),
    statusText: $("#status-text"),
  };

  // ── Inicialización ──
  function init() {
    setupDropZones();
    setupTabs();
    setupButtons();
    setDefaultDates();

    // Cargar datos de localStorage si existen
    loadState();
    // Precargar templates en background (silent fail si CORS)
    Generator.loadTemplates().catch(() => {});
  }

  // ── Drop zones ──
  function setupDropZones() {
    [dom.dropNelclock, dom.dropRoster].forEach((zone) => {
      zone.addEventListener("dragover", (e) => {
        e.preventDefault();
        zone.classList.add("drag-over");
      });
      zone.addEventListener("dragleave", () => zone.classList.remove("drag-over"));
      zone.addEventListener("drop", (e) => {
        e.preventDefault();
        zone.classList.remove("drag-over");
        const file = e.dataTransfer.files[0];
        if (file) handleFile(zone, file);
      });
      zone.addEventListener("click", () => {
        const input = zone.querySelector("input[type=file]");
        if (input) input.click();
      });
    });

    dom.inputNelclock.addEventListener("change", (e) => {
      if (e.target.files[0]) handleFile(dom.dropNelclock, e.target.files[0]);
    });
    dom.inputRoster.addEventListener("change", (e) => {
      if (e.target.files[0]) handleFile(dom.dropRoster, e.target.files[0]);
    });
  }

  function handleFile(zone, file) {
    const isNelclock = zone === dom.dropNelclock;
    const nameEl = isNelclock ? dom.fileNelclockName : dom.fileRosterName;

    if (isNelclock) {
      state.nelclockFile = file;
    } else {
      state.rosterFile = file;
    }

    nameEl.textContent = file.name;
    zone.classList.add("loaded");
    updateProcesarButton();
  }

  function updateProcesarButton() {
    dom.btnProcesar.disabled = !(state.nelclockFile && state.rosterFile);
    if (state.nelclockFile && state.rosterFile) {
      dom.btnProcesar.textContent = "Procesar datos";
    }
  }

  // ── Tabs ──
  function setupTabs() {
    document.addEventListener("click", (e) => {
      const tab = e.target.closest(".tab");
      if (!tab) return;

      const container = tab.closest(".card, section");
      if (!container) return;
      const tabId = tab.dataset.tab;
      if (!tabId) return;

      container.querySelectorAll(".tab").forEach((t) => t.classList.remove("active"));
      container.querySelectorAll(".tab-content").forEach((c) => c.classList.remove("active"));

      tab.classList.add("active");
      const content = document.getElementById(tabId);
      if (content) content.classList.add("active");

      // Re-renderizar si es un tab de resultados
      if (content && content.closest("#resultados-auto")) {
        renderTab(tabId);
      }
    });
  }

  // ── Buttons ──
  function setupButtons() {
    dom.btnProcesar.addEventListener("click", procesarDatos);

    $("#btn-agregar-novedad").addEventListener("click", agregarNovedad);
    $("#btn-agregar-hora-extra").addEventListener("click", agregarHoraExtra);

    dom.btnDescargarAusencias.addEventListener("click", descargarPlanillaAusencias);
    dom.btnDescargarTrasnoches.addEventListener("click", descargarTrasnoches);

    dom.selectMes.addEventListener("change", () => {
      state.mes = parseInt(dom.selectMes.value);
      saveState();
    });
    dom.inputAnio.addEventListener("change", () => {
      state.anio = parseInt(dom.inputAnio.value);
      saveState();
    });
  }

  function setDefaultDates() {
    const now = new Date();
    dom.selectMes.value = String(now.getMonth() + 1);
    dom.inputAnio.value = String(now.getFullYear());
    state.mes = now.getMonth() + 1;
    state.anio = now.getFullYear();

    document.getElementById("nov-fecha").valueAsDate = now;
    document.getElementById("he-fecha").valueAsDate = now;
  }

  // ── Procesamiento ──
  async function procesarDatos() {
    const mes = state.mes;
    const anio = state.anio;
    const targetMonth = `${anio}-${String(mes).padStart(2, "0")}`;

    dom.btnProcesar.textContent = "Procesando...";
    dom.btnProcesar.disabled = true;

    try {
      // 1. Parsear Nelclock
      const txtContent = await readFileAsText(state.nelclockFile);
      const nelclockData = NelclockParser.parse(txtContent);
      state.fichadas = nelclockData.fichadas;

      // 2. Parsear Roster (ExcelJS es async)
      const rosterBuf = await readFileAsBuffer(state.rosterFile);
      const rosterData = await RosterParser.parse(rosterBuf, targetMonth);
      state.rosterData = rosterData;

      // 3. Ejecutar motores
      state.resPuntualidad = Engines.puntualidad(rosterData, state.fichadas);
      state.resTrasnoches = Engines.trasnoches(rosterData, state.fichadas);
      state.res4taJornada = Engines.cuartaJornada(rosterData, state.fichadas);

      // 4. Mostrar resultados
      mostrarResultados();
      llenarDropdownsEmpleados(rosterData);

      dom.resultadosAuto.classList.remove("hidden");
      dom.descarga.classList.remove("hidden");
      setStatus("ok", "Datos procesados correctamente");
    } catch (err) {
      console.error(err);
      setStatus("warn", "Error al procesar: " + err.message);
    } finally {
      dom.btnProcesar.textContent = "Procesar datos";
      dom.btnProcesar.disabled = false;
    }
  }

  // ── Mostrar resultados con búsqueda y paginación ──
  const PAGE_SIZE = 15;
  const tabData = {};

  function mostrarResultados() {
    // Puntualidad
    const { perdieron, resultados } = state.resPuntualidad;
    dom.numPerdieron.textContent = perdieron.length;
    tabData["tab-puntualidad"] = {
      data: resultados,
      cols: 7,
      _page: 0,
      _filter: "",
      renderRow(r) {
        const estado = r.sinFichada
          ? '<span class="sin-fichada">Sin fichada</span>'
          : r.perdio ? '<span class="perdio">PERDIÓ</span>' : '<span class="conserva">OK</span>';
        return `<td>${esc(r.nombre)}</td><td>${esc(r.tipo)}</td><td>${esc(r.fecha)}</td><td>${esc(r.horaProgramada || "")}</td><td>${esc(r.horaReal || "—")}</td><td>${r.diferencia !== null ? r.diferencia : "—"}</td><td>${estado}</td>`;
      },
      searchFields(r) { return r.nombre; },
    };

    // Trasnoches
    const trasnoches = state.resTrasnoches.porEmpleado;
    const totalTrasnoches = trasnoches.reduce((sum, e) => sum + e.trasnoches.length, 0);
    dom.numTrasnoches.textContent = totalTrasnoches;
    const trasRegs = [];
    for (const emp of trasnoches) {
      for (const t of emp.trasnoches) {
        trasRegs.push({ ...t, nombre: emp.nombre });
      }
    }
    tabData["tab-trasnoches"] = {
      data: trasRegs,
      cols: 5,
      _page: 0,
      _filter: "",
      renderRow(r) {
        return `<td>${esc(r.nombre)}</td><td>${esc(r.fecha)}</td><td>${esc(r.horaEntrada)}</td><td>${esc(r.horaSalidaReal || r.horaSalidaProgramada)}</td><td>${r.horas.toFixed(2)}h</td>`;
      },
      searchFields(r) { return r.nombre + " " + r.fecha; },
    };

    // 4ta Jornada
    const cuarta = state.res4taJornada;
    dom.num4ta.textContent = cuarta.reduce((s, e) => s + e.cantidad, 0);
    tabData["tab-4ta"] = {
      data: cuarta,
      cols: 5,
      _page: 0,
      _filter: "",
      renderRow(item) {
        return `<td>${esc(item.nombre)}</td><td>${item.cantidad}</td><td>${item.semanas.map(s => s.semana).join(", ")}</td><td>${item.semanas.map(s => s.cantidadDias + "d").join(", ")}</td><td>${esc(item.motivo || "—")}</td>`;
      },
      searchFields(item) { return item.nombre; },
    };

    // Inyectar barras de búsqueda + paginación en cada tab
    setupSearchAndPagination();
    // Renderizar la tab activa
    renderCurrentTab();
  }

  function setupSearchAndPagination() {
    ["tab-puntualidad", "tab-trasnoches", "tab-4ta"].forEach((tabId) => {
      const container = document.getElementById(tabId);
      if (!container) return;

      // Buscar o crear barra de búsqueda
      let searchBar = container.querySelector(".search-bar");
      if (!searchBar) {
        searchBar = document.createElement("div");
        searchBar.className = "search-bar";
        searchBar.innerHTML = `<span class="search-icon">🔍</span><input type="text" placeholder="Buscar..." />`;
        container.insertBefore(searchBar, container.firstChild);
        const input = searchBar.querySelector("input");
        input.addEventListener("input", () => {
          const td = tabData[tabId];
          if (!td) return;
          td._page = 0;
          td._filter = input.value.toLowerCase();
          renderTab(tabId);
        });
      } else {
        // Resetear filtro si ya existía
        searchBar.querySelector("input").value = "";
      }

      // Buscar o crear paginación
      let pagDiv = container.querySelector(".pagination");
      if (!pagDiv) {
        pagDiv = document.createElement("div");
        pagDiv.className = "pagination";
        pagDiv.innerHTML = `<span class="page-info"></span><div class="page-buttons"></div>`;
        container.appendChild(pagDiv);
      }
    });
  }

  function renderCurrentTab() {
    const active = document.querySelector("#resultados-auto .tab-content.active");
    if (active) renderTab(active.id);
  }

  function renderTab(tabId) {
    const td = tabData[tabId];
    if (!td) return;

    const container = document.getElementById(tabId);
    if (!container) return;

    const filter = td._filter || "";
    let filtered = td.data;
    if (filter) {
      filtered = td.data.filter((r) => td.searchFields(r).toLowerCase().includes(filter));
    }

    const page = td._page || 0;
    const totalPages = Math.max(1, Math.ceil(filtered.length / PAGE_SIZE));
    if (page >= totalPages) td._page = totalPages - 1;
    const start = td._page * PAGE_SIZE;
    const slice = filtered.slice(start, start + PAGE_SIZE);

    const tbody = container.querySelector("tbody");
    if (tbody) {
      tbody.innerHTML = slice.length
        ? slice.map((r) => "<tr>" + td.renderRow(r) + "</tr>").join("")
        : `<tr><td colspan="${td.cols}">Sin resultados</td></tr>`;
    }

    // Actualizar paginación
    const pagDiv = container.querySelector(".pagination");
    if (pagDiv) {
      pagDiv.querySelector(".page-info").textContent = `${filtered.length} resultados · Pág ${td._page + 1} de ${totalPages}`;
      const btnDiv = pagDiv.querySelector(".page-buttons");
      btnDiv.innerHTML = "";
      const prevBtn = document.createElement("button");
      prevBtn.className = "btn btn-outline btn-page";
      prevBtn.textContent = "←";
      prevBtn.disabled = td._page === 0;
      prevBtn.onclick = () => { if (td._page > 0) { td._page--; renderTab(tabId); } };
      btnDiv.appendChild(prevBtn);

      const nextBtn = document.createElement("button");
      nextBtn.className = "btn btn-outline btn-page";
      nextBtn.textContent = "→";
      nextBtn.disabled = td._page >= totalPages - 1;
      nextBtn.onclick = () => { if (td._page < totalPages - 1) { td._page++; renderTab(tabId); } };
      btnDiv.appendChild(nextBtn);
    }
  }

  // ── Dropdowns de empleados ──
  function llenarDropdownsEmpleados(rosterData) {
    const nombres = rosterData
      .map((e) => e.nombre)
      .filter((v, i, a) => a.indexOf(v) === i)
      .sort((a, b) => a.localeCompare(b));

    const options = nombres.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join("");
    dom.novEmpleado.innerHTML = `<option value="">Seleccionar...</option>${options}`;
    dom.heEmpleado.innerHTML = `<option value="">Seleccionar...</option>${options}`;
  }

  // ── Novedades ──
  function agregarNovedad() {
    const empleado = dom.novEmpleado.value;
    const fecha = document.getElementById("nov-fecha").value;
    const motivo = document.getElementById("nov-motivo").value;
    const motivoCustom = document.getElementById("nov-motivo-custom").value;
    const aviso = document.getElementById("nov-aviso").value;
    const certificado = document.getElementById("nov-certificado").value;
    const llegadaTarde = document.getElementById("nov-llegada-tarde").value;
    const sancion = document.getElementById("nov-sancion").value;

    if (!empleado) return alert("Seleccioná un empleado");
    if (!fecha) return alert("Seleccioná una fecha");

    const motivoFinal = motivo === "Otro" ? motivoCustom : motivo;
    if (!motivoFinal) return alert("Seleccioná un motivo");

    state.novedades.push({
      nombre: empleado,
      fecha: fecha,
      motivo: motivoFinal,
      conAviso: aviso === "si",
      sinAviso: aviso === "no",
      certificado: certificado === "si",
      llegadaTarde: llegadaTarde || "",
      sancion: sancion,
    });

    renderNovedades();
    saveState();
    limpiarFormNovedad();
  }

  function renderNovedades() {
    dom.tbodyNovedades.innerHTML = state.novedades
      .map(
        (n, i) => `<tr>
          <td>${esc(n.nombre)}</td>
          <td>${esc(n.fecha)}</td>
          <td>${esc(n.motivo)}</td>
          <td>${n.certificado ? "Sí" : "No"}</td>
          <td>${esc(n.sancion || "—")}</td>
          <td><button class="btn btn-outline btn-sm" data-action="del-novedad" data-idx="${i}">✕</button></td>
        </tr>`
      )
      .join("");
  }

  function limpiarFormNovedad() {
    document.getElementById("nov-fecha").valueAsDate = new Date();
    document.getElementById("nov-motivo").value = "";
    document.getElementById("nov-motivo-custom").value = "";
    document.getElementById("nov-llegada-tarde").value = "";
    document.getElementById("nov-sancion").value = "";
  }

  // ── Horas Extras ──
  function agregarHoraExtra() {
    const empleado = dom.heEmpleado.value;
    const fecha = document.getElementById("he-fecha").value;
    const desde = document.getElementById("he-desde").value;
    const hasta = document.getElementById("he-hasta").value;
    const heDesde = document.getElementById("he-he-desde").value;
    const heHasta = document.getElementById("he-he-hasta").value;
    const cantidad = document.getElementById("he-cantidad").value;

    if (!empleado) return alert("Seleccioná un empleado");
    if (!fecha) return alert("Seleccioná una fecha");

    state.horasExtras.push({
      nombre: empleado,
      fecha: fecha,
      desde: desde,
      hasta: hasta,
      heDesde: heDesde,
      heHasta: heHasta,
      cantidad: cantidad || "2",
    });

    renderHorasExtras();
    saveState();
    limpiarFormHE();
  }

  function renderHorasExtras() {
    dom.tbodyHorasExtras.innerHTML = state.horasExtras
      .map(
        (h, i) => `<tr>
          <td>${esc(h.nombre)}</td>
          <td>${esc(h.fecha)}</td>
          <td>${esc(h.desde)}</td>
          <td>${esc(h.hasta)}</td>
          <td>${esc(h.cantidad)}</td>
          <td><button class="btn btn-outline btn-sm" data-action="del-he" data-idx="${i}">✕</button></td>
        </tr>`
      )
      .join("");
  }

  function limpiarFormHE() {
    document.getElementById("he-fecha").valueAsDate = new Date();
    document.getElementById("he-desde").value = "";
    document.getElementById("he-hasta").value = "";
    document.getElementById("he-he-desde").value = "";
    document.getElementById("he-he-hasta").value = "";
    document.getElementById("he-cantidad").value = "";
  }

  // ── Eliminar items ──
  document.addEventListener("click", (e) => {
    const btn = e.target.closest("button[data-action]");
    if (!btn) return;

    const action = btn.dataset.action;
    const idx = parseInt(btn.dataset.idx);

    if (action === "del-novedad") {
      state.novedades.splice(idx, 1);
      renderNovedades();
      saveState();
    }
    if (action === "del-he") {
      state.horasExtras.splice(idx, 1);
      renderHorasExtras();
      saveState();
    }
  });

  // ── Generar planillas ──
  async function descargarPlanillaAusencias() {
    try {
      const perdieronLista = state.resPuntualidad
        ? state.resPuntualidad.perdieron.map((p) => ({
            nombre: p.nombre,
            legajo: buscarLegajo(p.nombre),
          }))
        : [];

      const cuartaConLegajos = (state.res4taJornada || []).map((c) => ({
        ...c,
        legajo: buscarLegajo(c.nombre),
      }));

      const blob = await Generator.generarPlanillaAusencias({
        novedades: state.novedades,
        horasExtras: state.horasExtras,
        puntualidad: perdieronLista,
        cuartaJornada: cuartaConLegajos,
        mes: state.mes,
        año: state.anio,
      });

      triggerDownload(blob, `Planilla de Ausencias ${Generator.nombreMes(state.mes)} ${state.anio}.xlsx`);
      setStatus("ok", "Planilla de Ausencias descargada");
    } catch (err) {
      console.error(err);
      setStatus("warn", "Error: " + err.message);
    }
  }

  async function descargarTrasnoches() {
    try {
      const trasnochesData = state.resTrasnoches;
      if (!trasnochesData) return;

      // Construir estructura para el generador: incluir todos los empleados del roster
      const allEmpsFromRoster = state.rosterData.map(e => ({ nombre: e.nombre }));
      const porEmpleadoT = trasnochesData.porEmpleado;
      
      // Merge: empleados del roster que NO están en porEmpleado (sin trasnoche)
      const merged = [...porEmpleadoT];
      for (const re of allEmpsFromRoster) {
        if (!porEmpleadoT.some(pe => pe.nombre.toLowerCase() === re.nombre.toLowerCase())) {
          merged.push({ nombre: re.nombre, tipo: "FT", trasnoches: [] });
        }
      }

      // Detectar viernes y sábados del mes
      const viernes = [];
      const sabados = [];
      const diasMes = new Date(state.anio, state.mes, 0).getDate();
      for (let d = 1; d <= diasMes; d++) {
        const fecha = new Date(state.anio, state.mes - 1, d);
        if (fecha.getDay() === 5) viernes.push(fecha);
        if (fecha.getDay() === 6) sabados.push(fecha);
      }

      const viernesStr = viernes.map((d) => d.toISOString().slice(0, 10));
      const sabadosStr = sabados.map((d) => d.toISOString().slice(0, 10));

      // Mapear empleados con sus trasnoches por viernes/sábado
      const empleadosData = merged.map((emp) => {
        const e = { nombre: emp.nombre, legajo: buscarLegajo(emp.nombre), viernes: [], sabados: [] };
        for (let i = 0; i < viernes.length; i++) {
          const vKey = viernesStr[i];
          const t = (emp.trasnoches || []).find((t) => t.fecha === vKey);
          e.viernes.push(t ? { horas: t.horas, salida: t.horaSalidaReal || t.horaSalidaProgramada } : { horas: 0, salida: null });
        }
        for (let i = 0; i < sabados.length; i++) {
          const sKey = sabadosStr[i];
          const t = (emp.trasnoches || []).find((t) => t.fecha === sKey);
          e.sabados.push(t ? { horas: t.horas, salida: t.horaSalidaReal || t.horaSalidaProgramada } : { horas: 0, salida: null });
        }
        return e;
      });

      const blob = await Generator.generarTrasnoches({
        empleados: empleadosData,
        viernes: viernesStr,
        sabados: sabadosStr,
        mes: state.mes,
        año: state.anio,
      });

      triggerDownload(blob, `Trasnoches AVEL ${Generator.nombreMesAbbr(state.mes)} ${String(state.anio).slice(-2)}.xlsx`);
      setStatus("ok", "Trasnoches AVEL descargado");
    } catch (err) {
      console.error(err);
      setStatus("warn", "Error: " + err.message);
    }
  }

  // ── Helpers ──
  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsText(file, "ISO-8859-1");
    });
  }

  function readFileAsBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  function setStatus(type, text) {
    dom.statusDot.className = "dot " + type;
    dom.statusText.textContent = text;
  }

  function esc(str) {
    if (!str) return "";
    const s = String(str);
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  // ── Persistencia en localStorage ──
  function saveState() {
    try {
      const data = {
        novedades: state.novedades,
        horasExtras: state.horasExtras,
        mes: state.mes,
        anio: state.anio,
      };
      localStorage.setItem("cine-rrhh", JSON.stringify(data));
    } catch (_) {}
  }

  function loadState() {
    try {
      const raw = localStorage.getItem("cine-rrhh");
      if (!raw) return;
      const data = JSON.parse(raw);
      if (data.mes) {
        state.mes = data.mes;
        dom.selectMes.value = String(data.mes);
      }
      if (data.anio) {
        state.anio = data.anio;
        dom.inputAnio.value = String(data.anio);
      }
      if (data.novedades) {
        state.novedades = data.novedades;
        renderNovedades();
      }
      if (data.horasExtras) {
        state.horasExtras = data.horasExtras;
        renderHorasExtras();
      }
    } catch (_) {}
  }

  // ── Init ──
  init();
})();
