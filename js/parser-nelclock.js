/**
 * Parser del archivo TXT exportado de Nelclock (Reporte Municipal)
 * Formato de ancho fijo con separadores de línea "---"
 */
const NelclockParser = {
  /**
   * Parsea el contenido del TXT de Nelclock
   * @param {string} text - Contenido del archivo TXT
   * @returns {{fecha_reporte: string, fichadas: Array}}
   */
  parse(text) {
    const lines = text.split(/\r?\n/);
    const fichadas = [];

    const dataLineRegex =
      /^\s+(\d{6,8})\s+(DNI|LE|LC|PAS)\s+(\d{6,10})\s+(\d{11})\s+(.+?)\s{2,}(.+?)\s{2,}(\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2})(?:\s+(\d{2}\/\d{2}\/\d{2}\s+\d{2}:\d{2}:\d{2}))?/;

    for (const line of lines) {
      const m = line.match(dataLineRegex);
      if (!m) continue;

      const [, legajo, , , , apellido, nombre, entradaStr, salidaStr] = m;
      const entrada = this.parseDatetime(entradaStr);
      const salida = salidaStr ? this.parseDatetime(salidaStr) : null;

      fichadas.push({
        legajo: legajo.replace(/^0+/, ""),
        apellido: apellido.trim(),
        nombre: nombre.trim(),
        entrada,
        salida,
      });
    }

    return { fichadas };
  },

  parseDatetime(str) {
    const [datePart, timePart] = str.split(/\s+/);
    const [d, m, y] = datePart.split("/");
    const [h, min, s] = timePart.split(":");
    const year = 2000 + parseInt(y, 10);
    return new Date(year, parseInt(m, 10) - 1, parseInt(d, 10), parseInt(h, 10), parseInt(min, 10), parseInt(s, 10));
  },
};
