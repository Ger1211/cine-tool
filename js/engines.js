/**
 * Motores de procesamiento: puntualidad, trasnoches, 4ta jornada
 */
const Engines = {
  /**
   * Calcula quién perdió la puntualidad.
   * Compara la hora de entrada programada (roster) vs. la real (Nelclock).
   * @param {Array} rosterData - Array de empleados con sus días [empleado, ...]
   * @param {Array} fichadas - Array de fichadas de Nelclock [{legajo, entrada, salida, nombre, apellido}, ...]
   * @param {number} toleranciaMinutos - Minutos de tolerancia (default 5)
   * @returns {{perdieron: Array, conservan: Array, todos: Array}}
   */
  puntualidad(rosterData, fichadas, toleranciaMinutos = 5) {
    const fichadasByName = this.indexarFichadas(fichadas);
    const resultados = [];

    for (const emp of rosterData) {
      const fichas = this.buscarFichadas(emp.nombre, fichadasByName);

      for (const dia of emp.dias) {
        if (!dia.entrada || dia.franco || dia.licencia || dia.vacaciones) continue;

        const fecha = dia.fecha;
        const horaProgramada = dia.entrada;
        const fichadaDia = fichas.find((f) => this.mismaFecha(f.entrada, fecha));

        if (!fichadaDia) {
          resultados.push({
            nombre: emp.nombre,
            tipo: emp.tipo,
            fecha,
            horaProgramada,
            horaReal: null,
            diferencia: null,
            perdio: false,
            sinFichada: true,
          });
          continue;
        }

        const horaReal = fichadaDia.entrada;
        const programadaMin = this.timeToMinutes(horaProgramada);
        const realMin = this.timeToMinutes(horaReal);
        const diff = realMin - programadaMin;
        const perdio = diff > toleranciaMinutos;

        resultados.push({
          nombre: emp.nombre,
          tipo: emp.tipo,
          fecha,
          horaProgramada,
          horaReal: this.formatMinutes(realMin),
          diferencia: diff,
          perdio,
          sinFichada: false,
        });
      }
    }

    // Agrupar por empleado: alguien perdió si perdió al menos una vez en el mes
    const perdieron = [];
    const conservan = [];
    const procesados = new Set();

    for (const r of resultados) {
      if (procesados.has(r.nombre)) continue;
      procesados.add(r.nombre);

      const todasLasFallas = resultados.filter((x) => x.nombre === r.nombre && x.perdio);
      if (todasLasFallas.length > 0) {
        perdieron.push({
          nombre: r.nombre,
          tipo: r.tipo,
          fallas: todasLasFallas.map((x) => ({
            fecha: x.fecha,
            horaProgramada: x.horaProgramada,
            horaReal: x.horaReal,
            diferencia: x.diferencia,
          })),
        });
      }
    }

    return { perdieron, resultados };
  },

  /**
   * Detecta trasnoches: turnos con salida posterior a medianoche.
   * Compara lo programado en el roster con lo registrado en Nelclock.
   * @param {Array} rosterData
   * @param {Array} fichadas
   * @returns {{porFecha: Object, porEmpleado: Array}}
   */
  trasnoches(rosterData, fichadas) {
    const fichadasByName = this.indexarFichadas(fichadas);
    const porEmpleado = new Map();

    for (const emp of rosterData) {
      const fichas = this.buscarFichadas(emp.nombre, fichadasByName);

      for (const dia of emp.dias) {
        if (!dia.entrada || !dia.salida || dia.franco || dia.licencia || dia.vacaciones) continue;

        // Trasnoche = salida programada es después de las 00:00 (ej: "01:00:00")
        const salidaMin = this.timeToMinutes(dia.salida);
        const entradaMin = this.timeToMinutes(dia.entrada);

        // Si la salida es menor que la entrada, es porque cruza medianoche
        const esTrasnocheProgramada = salidaMin < entradaMin && salidaMin < 360; // salida antes de las 6am

        if (!esTrasnocheProgramada) continue;

        const fecha = dia.fecha;
        const fichadaDia = fichas.find((f) => this.mismaFecha(f.entrada, fecha));

        let horas = 0;
        let salidaReal = null;

        if (fichadaDia && fichadaDia.salida) {
          const salidaRealMin = this.timeToMinutes(fichadaDia.salida);
          const entradaRealMin = this.timeToMinutes(fichadaDia.entrada);
          salidaReal = fichadaDia.salida;

          if (salidaRealMin < entradaRealMin) {
            // También cruza medianoche
            horas = (1440 - entradaRealMin + salidaRealMin) / 60;
          } else {
            horas = (salidaRealMin - entradaRealMin) / 60;
          }
        } else {
          // Usamos lo programado como estimación
          horas = (1440 - entradaMin + salidaMin) / 60;
        }

        const key = emp.nombre;
        if (!porEmpleado.has(key)) {
          porEmpleado.set(key, {
            nombre: emp.nombre,
            tipo: emp.tipo,
            trasnoches: [],
          });
        }

        porEmpleado.get(key).trasnoches.push({
          fecha,
          horaEntrada: dia.entrada,
          horaSalidaProgramada: dia.salida,
          horaSalidaReal: salidaReal || dia.salida,
          horas,
        });
      }
    }

    const resultado = Array.from(porEmpleado.values());

    // Por fecha (para el formato de la planilla de trasnoches por día)
    const porFecha = {};
    for (const emp of resultado) {
      for (const t of emp.trasnoches) {
        if (!porFecha[t.fecha]) porFecha[t.fecha] = [];
        porFecha[t.fecha].push({
          nombre: emp.nombre,
          salida: t.horaSalidaReal,
          horas: t.horas,
        });
      }
    }

    return { porEmpleado: resultado, porFecha };
  },

  /**
   * Detecta 4ta jornada para empleados Part Time.
   * Cuenta cuántas semanas tuvieron 4+ días laborales.
   * @param {Array} rosterData
   * @returns {Array} [{nombre, cantidad, semanas: [...]}]
   */
  cuartaJornada(rosterData, fichadas) {
    const pt = rosterData.filter((e) => e.tipo === "PT");
    const fichadasByName = this.indexarFichadas(fichadas);
    const porEmpleado = new Map();

    for (const emp of pt) {
      const semanas = new Map();
      const fichas = this.buscarFichadas(emp.nombre, fichadasByName);

      for (const dia of emp.dias) {
        if (dia.franco || dia.vacaciones) continue;
        // Contamos también licencias para 4ta jornada
        if (!dia.entrada && !dia.salida && !dia.licencia) continue;

        const fecha = new Date(dia.fecha + "T00:00:00");
        const semanaKey = this.getWeekKey(fecha);

        if (!semanas.has(semanaKey)) semanas.set(semanaKey, []);
        semanas.get(semanaKey).push(dia);
      }

      const semanasCuatro = [];
      for (const [semanaKey, dias] of semanas) {
        if (dias.length < 4) continue;

        // Determinar motivo: si algún día fue licencia o ausencia
        let motivo = "";
        for (const d of dias) {
          if (d.licencia) {
            motivo = motivo ? motivo : "LICENCIA";
          } else if (d.entrada && !d.franco) {
            // Verificar si fichó (ausencia)
            const fichadaDia = fichas.find((f) => this.mismaFecha(f.entrada, d.fecha));
            if (!fichadaDia) {
              motivo = motivo ? motivo : "AUSENCIA";
            }
          }
        }

        semanasCuatro.push({
          semana: semanaKey,
          cantidadDias: dias.length,
          fechas: dias.map((d) => d.fecha),
          motivo,
        });
      }

      if (semanasCuatro.length > 0) {
        // Motivo consolidado: el más grave entre todas las semanas
        const motivosUnicos = [...new Set(semanasCuatro.map((s) => s.motivo).filter(Boolean))];
        porEmpleado.set(emp.nombre, {
          nombre: emp.nombre,
          cantidad: semanasCuatro.length,
          semanas: semanasCuatro,
          motivo: motivosUnicos.join(", "),
        });
      }
    }

    return Array.from(porEmpleado.values());
  },

  // ── Helpers ──

  /**
   * Indexa fichadas de Nelclock para búsqueda flexible por nombre.
   * Genera múltiples variantes de nombre como clave (apellido+nombre, nombre+apellido).
   */
  indexarFichadas(fichadas) {
    const map = new Map();
    for (const f of fichadas) {
      const apellido = (f.apellido || "").toLowerCase().trim();
      const nombre = (f.nombre || "").toLowerCase().trim();
      const claves = new Set();

      claves.add(`${apellido} ${nombre}`.replace(/\s+/g, " ").trim());
      claves.add(`${nombre} ${apellido}`.replace(/\s+/g, " ").trim());

      for (const clave of claves) {
        if (!map.has(clave)) map.set(clave, []);
        map.get(clave).push(f);
      }
    }
    return map;
  },

  /**
   * Busca fichadas para un empleado del roster intentando múltiples variantes.
   */
  buscarFichadas(nombreRoster, fichadasMap) {
    const nombre = nombreRoster.toLowerCase().replace(/\s+/g, " ").trim();
    const palabras = nombre.split(/\s+/).filter((w) => w.length > 1);

    // Búsqueda exacta
    if (fichadasMap.has(nombre)) return fichadasMap.get(nombre);

    // Búsqueda: probar cada clave contra las palabras del nombre del roster
    for (const [clave, fichas] of fichadasMap) {
      const todasPresentes = palabras.every((p) => clave.includes(p));
      if (todasPresentes) return fichas;
    }

    // Búsqueda inversa: palabras de la clave contra el nombre del roster
    for (const [clave, fichas] of fichadasMap) {
      const palabrasClave = clave.split(/\s+/).filter((w) => w.length > 1);
      const todasPresentes = palabrasClave.every((p) => nombre.includes(p));
      if (todasPresentes) return fichas;
    }

    return [];
  },

  mismaFecha(d1, fechaStr) {
    if (!d1) return false;
    const date = d1 instanceof Date ? d1 : new Date(d1);
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}` === fechaStr;
  },

  timeToMinutes(timeInput) {
    if (timeInput instanceof Date) {
      return timeInput.getHours() * 60 + timeInput.getMinutes();
    }
    const str = String(timeInput).trim();
    const parts = str.split(":");
    return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
  },

  formatMinutes(mins) {
    const sign = mins < 0 ? "-" : "";
    const abs = Math.abs(mins);
    const h = Math.floor(abs / 60);
    const m = abs % 60;
    return `${sign}${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
  },

  getWeekKey(date) {
    const d = new Date(date);
    const dayOfWeek = d.getDay();
    const diff = d.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1);
    const monday = new Date(d.setDate(diff));
    const y = monday.getFullYear();
    const mo = String(monday.getMonth() + 1).padStart(2, "0");
    const day = String(monday.getDate()).padStart(2, "0");
    return `${y}-${mo}-${day}`;
  },
};
