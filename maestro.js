// BenitoStock · constructor del MAESTRO: las hojas con fórmulas (Dashboard «A ejecutivo», TIEMPOS, LÍNEA y TOMAS) y las columnas de ayuda de las hojas de datos.
// Lo usan DOS sitios con el MISMO código: la app (app.js, «Descargar MAESTRO») y el generador de ejemplos de analisis/ (toma-maestro-ejemplo.mjs), así lo que se prueba allí es lo que baja la app.
// No dibuja las hojas de datos (POSICIONES, UNIDADES, INCIDENCIAS, POR CÓDIGO, FRANJAS, PAUSAS), la hoja LÉEME ni la marca de agua: eso lo pone quien llama; aquí solo se fijan las columnas que las fórmulas esperan.
// Las fórmulas miran las hojas de datos por la LETRA de la columna: COLS_ESPERADAS dice qué nombre debe tener cada columna que se usa y verificar() lo comprueba antes de armar nada.
// Solo funciones de Excel antiguo (COUNTIFS, SUMIFS, IFERROR, INDEX, MATCH, MAX, MIN, IF, AND, OR, ROUND, COUNT, COUNTA, COUNTIF, ISNUMBER, LEFT, LEN).
(function (raiz, fabrica) {
  if (typeof module === 'object' && module.exports) module.exports = fabrica(); else raiz.Maestro = fabrica();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  const AMARILLO = 'FFFFC21A', CARBON = 'FF14161B', GRIS = 'FFF1EFE8', GRIS_TXT = 'FF52514E', ROJO = 'FFA32D2D';
  const F_DIF = '[Blue]+0;[Red]-0;0';   // secciones: positivo;negativo;cero. En la del negativo Excel muestra el valor ABSOLUTO, por eso el «-» se escribe a mano. Rojo = menos, azul = más.
  const F_DELTA_PTS = '[Color10]"▲ +"0.0" pts vs anterior";[Red]"▼ −"0.0" pts vs anterior";"igual que la anterior"';   // sube = bien
  const relleno = (argb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb }, bgColor: { argb } });
  const BORDE = { top: { style: 'thin' }, bottom: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' } };
  const ESCALA = (ref) => ({ ref, rules: [{ type: 'colorScale', cfvo: [{ type: 'num', value: 0.91 }, { type: 'num', value: 0.95 }, { type: 'num', value: 0.98 }], color: [{ argb: 'FFF4B6B6' }, { argb: 'FFFBE3A6' }, { argb: 'FFB9DFB0' }] }] });
  const letra = (n) => { let s = ''; for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };   // 1 → A
  const numLetra = (l) => [...l].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);                                                              // A → 1
  const f = (x) => ({ formula: x });
  const rg = (h, c, n) => `${h}!$${c}$2:$${c}$${n + 1}`;   // n = filas de DATOS: de la 2 a la n+1 (la 1 son los títulos)
  function encabezado(ws, fila, nombres, desdeCol) {
    nombres.forEach((n, i) => { const c = ws.getCell(fila, (desdeCol || 1) + i); c.value = n; c.font = { bold: true, color: { argb: CARBON } }; c.fill = relleno(AMARILLO); c.alignment = { wrapText: true, vertical: 'middle' }; });
  }

  // Hasta dónde llegan las fórmulas. NP, NU, NPA, NI, NPC y NF = FILAS DE DATOS que miran en POSICIONES, UNIDADES, PAUSAS, INCIDENCIAS, POR CÓDIGO y FRANJAS; NT = tomas que lista TOMAS;
  // NTI = unidades que listan TIEMPOS y LÍNEA; NPL = pausas que lista TIEMPOS; NM = pasillos y torres del mapa de calor.
  const TAM_EJEMPLO = { NP: 10000, NU: 80, NPA: 150, NI: 500, NPC: 3500, NF: 800, NT: 12, NTI: 12, NPL: 15, NM: 12 };   // chico a propósito: HyperFormula, que calcula las pruebas, se ahoga con rangos grandes
  const TAM_REAL = { NP: 60000, NU: 600, NPA: 2000, NI: 6000, NPC: 24000, NF: 12000, NT: 24, NTI: 80, NPL: 40, NM: 120 };   // varias tomas completas (≈ 8.000 posiciones, 64 unidades y 3.000 códigos cada una)

  // Las columnas de las hojas de datos que usan las fórmulas, con el nombre que deben tener. Si el contrato de la app cambia (se mueve o se renombra una), verificar() falla ANTES de armar un MAESTRO que calcularía mal.
  const COLS_ESPERADAS = {
    POSICIONES: { A: 'ID toma', L: 'Stock sistema', M: 'Contado', N: 'Diferencia' },
    UNIDADES: { A: 'ID toma', B: 'Fecha toma', C: 'Unidad', D: 'Tipo', I: 'Contador', N: 'Contadas', O: 'Reconteo de', Z: 'Min. abierta', AC: 'Mayor hueco (min)', AD: 'Ritmo mediano (s)', AE: 'Con diferencia' },
    INCIDENCIAS: { A: 'ID toma', B: 'Tipo' },
    'POR CÓDIGO': { A: 'ID toma', B: 'Código', C: 'Descripción', E: 'Diferencia neta' },
    FRANJAS: { A: 'ID toma', C: 'Unidad', E: 'Franja n.º', F: 'Guardados' },
    PAUSAS: { A: 'ID toma', B: 'Contador', C: 'Unidad', D: 'Tipo', E: 'Desde', F: 'Hasta', G: 'Min', J: 'Confirma servidor', K: 'Antes (UDC·POS)', L: 'Después (UDC·POS)' },
  };
  function verificar(contratos) {   // contratos = { hoja: [nombres de columna en orden] }
    if (!contratos || typeof contratos !== 'object') throw new Error('MAESTRO: falta el contrato de columnas de las hojas de datos');
    for (const [hoja, cols] of Object.entries(COLS_ESPERADAS)) {
      const reales = contratos[hoja];
      if (!Array.isArray(reales)) throw new Error(`MAESTRO: falta la hoja ${hoja} en el contrato de columnas`);
      for (const [l, nombre] of Object.entries(cols)) {
        const real = reales[numLetra(l) - 1];
        if (real !== nombre) throw new Error(`MAESTRO: en ${hoja} la columna ${l} debe ser «${nombre}» y es «${real === undefined ? '(no existe)' : real}». Las fórmulas miran las columnas por su letra: hay que volver a poner la columna en su lugar o ajustar las fórmulas.`);
      }
    }
  }

  function crear(tam, opciones) {
    opciones = opciones || {};
    const { NP, NU, NPA, NI, NPC, NF, NT } = tam, NTI = tam.NTI ?? TAM_EJEMPLO.NTI, NPL = tam.NPL ?? TAM_EJEMPLO.NPL, NM = tam.NM ?? TAM_EJEMPLO.NM;
    const EJ = opciones.ejemplo ? 'EJEMPLO con datos ficticios · ' : '';
    // dónde está cada cosa (las pruebas y el generador leen estas direcciones; todo lo que depende del tamaño se deriva de él)
    const TITULO = 8 + NTI + 1, L0 = TITULO + 3, M0 = 50;
    const TIEM = { selector: 'B4', mostrada: 'D4', umbral: 'B5', aviso: 'A6', tabla: { fila0: 8, filas: NTI }, titulo: TITULO, nota: TITULO + 1, cab: TITULO + 2, lista: { fila0: L0, filas: NPL }, ultimaFila: L0 + NPL - 1 };
    const LINEA = { slots: [24, 87], fila0: 7, filas: NTI, ultimaFila: 7 + NTI - 1 };   // franjas de 15 min: 24 = 06:00 … 87 = 21:45
    // el mapa de calor, que puede ser largo, va al final del Dashboard
    const DASH = { banner: 'A3', tiles: { eri: 'A5', eru: 'D5', cob: 'G5', neto: 'J5' }, metas: { eri: 'C9', eru: 'E9', cob: 'G9' }, seccion: { tomas: 11, mirar: 26, codigos: 31, incid: 44, mapa: M0 - 2 },
      tomas: { fila0: 13, filas: 12 }, mirar: [27, 28, 29], codigos: { fila0: 33, filas: 10 }, incid: { fila: 46 }, mapa: { fila0: M0, filas: NM, cols: 8, ult: 'J' }, ultimaFila: M0 + NM - 1 };
    // la tabla «Capacidad de las fórmulas» (a la derecha de TOMAS): cuántas filas hay en cada hoja de datos y cuántas alcanzan a mirar las fórmulas
    const CAP = { fila0: 4, hojas: ['Tomas', 'POSICIONES', 'UNIDADES', 'INCIDENCIAS', 'POR CÓDIGO', 'FRANJAS', 'PAUSAS'], usadas: 'U', caben: 'V', estado: 'W', llenas: 'U11' };
    const P = (c) => rg('POSICIONES', c, NP), U = (c) => rg('UNIDADES', c, NU), PA = (c) => rg('PAUSAS', c, NPA), IN = (c) => rg('INCIDENCIAS', c, NI);
    const PC = (c) => rg("'POR CÓDIGO'", c, NPC), FR = (c) => rg('FRANJAS', c, NF), TO = (c) => rg('TOMAS', c, NT), NN = 'TOMAS!$U$1';

    // columnas de ayuda a la derecha de lo que se pega: numeran la primera aparición de cada toma, unidad o código y marcan lo duplicado
    function ayudas(ws, desdeCol, cabs, ultima, fn) {
      encabezado(ws, 1, cabs, desdeCol); cabs.forEach((_, i) => { ws.getColumn(desdeCol + i).width = 16; ws.getCell(1, desdeCol + i).fill = relleno('FFD9D9D9'); });
      for (let r = 2; r <= ultima; r++) fn(r).forEach((x, i) => { if (x) ws.getCell(r, desdeCol + i).value = f(x); });
    }

    function hojaTomas(wb) {
      const ws = wb.addWorksheet('TOMAS', { views: [{ state: 'frozen', ySplit: 1 }], properties: { tabColor: { argb: 'FF7F7F7F' } } });
      const cols = ['ID toma', 'Fecha', 'Carga', 'Posic.', 'Contadas', 'Menos (pos.)', 'Más (pos.)', 'ERI posiciones', 'Inesperadas', 'Halladas vacías', 'ERU aprox.', 'Cobertura', 'Falta (u)', 'Sobra (u)', 'Σ |dif| (u)', 'Stock contado (u)', 'ERI unidades', 'Neto (u)'];
      [20, 12, 36, 9, 10, 10, 10, 11, 11, 11, 10, 10, 10, 10, 11, 12, 11, 10, 3, 22, 14, 16, 10].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
      encabezado(ws, 1, cols); ws.getRow(1).height = 30;
      for (let r = 2; r <= NT + 1; r++) {
        const a = `$A${r}`, i = r - 1, c = (e) => f(`IF(${a}="","",${e})`), at = `MATCH(${i},${U('AG')},0)`;
        ws.getCell(r, 1).value = f(`IFERROR(INDEX(${U('A')},${at}),"")`); ws.getCell(r, 2).value = c(`INDEX(${U('B')},${at})`);
        // la fecha de cada toma no puede ser anterior a la de arriba: si se pegó una toma vieja después de una nueva, «vs anterior» compararía mal
        const orden = r > 2 ? `IF(B${r}<B${r - 1},"Revisar: fecha anterior a la toma de arriba (pega las tomas en orden de fecha)","OK")` : '"OK"';
        ws.getCell(r, 3).value = c(`IF(D${r}=0,"Falta POSICIONES",IF(COUNTIFS(${U('A')},${a},${U('AI')},"DUPLICADA")>0,"Revisar: unidades duplicadas",IF(SUMIFS(${U('N')},${U('A')},${a},${U('O')},"")<>E${r},"Revisar: POSICIONES no cuadra con UNIDADES",IF(COUNTIFS(${FR('A')},${a})=0,"Sin tiempos (falta FRANJAS)",${orden}))))`);
        const PA_ = P('A'), N_ = P('N');
        [`COUNTIFS(${PA_},${a})`, `COUNTIFS(${PA_},${a},${P('M')},"<>")`, `COUNTIFS(${PA_},${a},${N_},"<0")`, `COUNTIFS(${PA_},${a},${N_},">0")`, `IFERROR(1-(F${r}+G${r})/E${r},"")`,
          `COUNTIFS(${PA_},${a},${P('L')},0,${P('M')},">0")`, `COUNTIFS(${PA_},${a},${P('L')},">0",${P('M')},0)`, `IFERROR(1-(I${r}+J${r})/E${r},"")`, `IFERROR(E${r}/D${r},"")`,
          `SUMIFS(${N_},${PA_},${a},${N_},"<0")`, `SUMIFS(${N_},${PA_},${a},${N_},">0")`, `N${r}-M${r}`, `SUMIFS(${P('L')},${PA_},${a},${P('M')},"<>")`, `IFERROR(1-O${r}/P${r},"")`, `M${r}+N${r}`]
          .forEach((e, j) => { const cel = ws.getCell(r, 4 + j); cel.value = c(e); cel.numFmt = [4, 7, 8, 11].includes(j) ? '0.0%' : [9, 10, 14].includes(j) ? F_DIF : '#,##0'; });
        ws.getCell(r, 2).numFmt = 'dd/mm/yyyy';
      }
      // «Tomas cargadas» con tope: si hay más tomas de las que caben en la tabla, las cuentas siguen mirando solo las NT primeras (el aviso del Dashboard lo dice)
      ws.getCell('T1').value = 'Tomas cargadas'; ws.getCell('U1').value = f(`MIN(${NT},MAX(${U('AG')}))`); ws.getCell('T1').font = { bold: true };
      ws.autoFilter = { from: 'A1', to: 'R1' };
      // capacidad de las fórmulas: se cuenta hasta 2·n + 5000 filas para ver cuántas hay aunque sobrepasen lo que las fórmulas alcanzan a mirar
      ws.getCell('T2').value = 'Capacidad de las fórmulas'; ws.getCell('T2').font = { bold: true, color: { argb: CARBON } };
      encabezado(ws, 3, ['Hoja', 'Filas usadas', 'Filas que caben', 'Estado'], 20); ws.getRow(3).height = 30;
      const caben = { Tomas: NT, POSICIONES: NP, UNIDADES: NU, INCIDENCIAS: NI, 'POR CÓDIGO': NPC, FRANJAS: NF, PAUSAS: NPA };
      CAP.hojas.forEach((h, i) => {
        const r = CAP.fila0 + i, n = caben[h];
        ws.getCell(`T${r}`).value = h;
        ws.getCell(`U${r}`).value = f(h === 'Tomas' ? `MAX(${U('AG')})` : `COUNTA(${h === 'POR CÓDIGO' ? "'POR CÓDIGO'" : h}!$A$2:$A$${2 * n + 5000})`);
        ws.getCell(`V${r}`).value = n; ws.getCell(`V${r}`).numFmt = '#,##0'; ws.getCell(`U${r}`).numFmt = '#,##0';
        ws.getCell(`W${r}`).value = f(`IF(U${r}>V${r},"LLENA","")`); ws.getCell(`W${r}`).alignment = { horizontal: 'center' };
      });
      const FIN = CAP.fila0 + CAP.hojas.length;
      ws.getCell(`T${FIN}`).value = 'Hojas llenas:'; ws.getCell(`T${FIN}`).font = { italic: true, color: { argb: GRIS_TXT } };
      ws.getCell(CAP.llenas).value = f(CAP.hojas.map((_, i) => `IF(W${CAP.fila0 + i}="LLENA",T${CAP.fila0 + i}&", ","")`).join('&'));
      ws.addConditionalFormatting({ ref: `W${CAP.fila0}:W${FIN - 1}`, rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"LLENA"'], style: { fill: relleno('FFF9D2D2'), font: { bold: true, color: { argb: 'FF791F1F' } } } }] });
    }

    function hojaTiempos(wb) {
      const d = wb.addWorksheet('TIEMPOS', { views: [{ showGridLines: false }], properties: { tabColor: { argb: 'FFA32D2D' } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      [12, 14, 12, 14, 12, 13, 12, 26, 22, 12, 12, 12, 3, 9].forEach((w, i) => { d.getColumn(i + 1).width = w; });
      d.getColumn(14).hidden = true;
      d.getCell('A1').value = 'TIEMPOS por contador'; d.getCell('A1').font = { size: 16, bold: true, color: { argb: CARBON } };
      d.mergeCells('A2:L2'); d.getCell('A2').value = EJ + 'CONFIDENCIAL (Ley 29733): sirve para mejorar el proceso, no para sancionar por sí solo. Un hueco prueba que no hubo guardados, no el porqué.'; d.getCell('A2').font = { italic: true, size: 10, color: { argb: ROJO } };
      d.getCell('A4').value = 'Toma (vacío = la última) →'; d.getCell('A4').font = { bold: true }; d.getCell('A5').value = 'Umbral de pausa (min) →'; d.getCell('A5').font = { bold: true };
      for (const a of [TIEM.selector, TIEM.umbral]) { const c = d.getCell(a); c.fill = relleno('FFFFF2B3'); c.border = BORDE; c.alignment = { horizontal: 'center' }; c.font = { bold: true }; }
      d.getCell(TIEM.umbral).value = 5;
      d.getCell(TIEM.selector).dataValidation = { type: 'list', allowBlank: true, formulae: ['ListaTomas'], showErrorMessage: true, errorTitle: 'Toma', error: 'Elige una toma de la lista o deja la casilla vacía.' };
      d.getCell(TIEM.umbral).dataValidation = { type: 'decimal', operator: 'greaterThanOrEqual', allowBlank: false, formulae: [3], showErrorMessage: true, errorTitle: 'Umbral', error: 'Escribe minutos (3 o más: las pausas más cortas no se guardan).' };
      d.getCell('C4').value = 'Mostrando:'; d.getCell('C4').font = { italic: true }; d.getCell('C4').alignment = { horizontal: 'right' };
      d.mergeCells('D4:F4'); d.getCell(TIEM.mostrada).value = f(`IF($B$4="",IF(${NN}<1,"",INDEX(${TO('A')},${NN})),$B$4)`); d.getCell(TIEM.mostrada).font = { bold: true };
      // si la toma tiene más unidades de las que caben en la tabla, las demás no se ven: se dice
      d.mergeCells('A6:L6'); d.getCell(TIEM.aviso).value = f(`IF(MAX(${U('AJ')})>${NTI},"⚠ Solo caben ${NTI} filas y esta toma tiene "&MAX(${U('AJ')})&" unidades: las demás no se listan aquí ni en la hoja LÍNEA.","")`); d.getCell(TIEM.aviso).font = { bold: true, size: 10, color: { argb: ROJO } };
      encabezado(d, 7, ['Contador', 'Unidad', 'Posiciones', 'Unidad abierta (min)', 'Sin registros (min)', '% sin registros', 'Huecos > umbral', 'Mayor hueco (min)', 'Ritmo mediano (s)', 'Pos/h neto', '% con diferencia', 'Reconteos (pos.)'], 1); d.getRow(7).height = 32;
      const T0 = TIEM.tabla.fila0, D4 = '$D$4', B5 = '$B$5';
      for (let i = 1; i <= NTI; i++) {
        const r = T0 + i - 1, n = `$N${r}`, c = (e) => f(`IF(${n}="","",${e})`), ix = (col) => `INDEX(${U(col)},${n})`;
        d.getCell(r, 14).value = f(`IFERROR(MATCH(${i},${U('AJ')},0),"")`);
        [ix('I'), ix('C'), ix('N'), ix('Z'), `SUMIFS(${PA('G')},${PA('A')},${D4},${PA('C')},$B${r},${PA('G')},">"&${B5})`, `IFERROR(E${r}/D${r},"")`,
          `COUNTIFS(${PA('A')},${D4},${PA('C')},$B${r},${PA('D')},"Pausa",${PA('G')},">"&${B5})`, ix('AC'), ix('AD'), `IFERROR(C${r}/((D${r}-E${r})/60),"")`, `IFERROR(${ix('AE')}/C${r},"")`,
          `SUMIFS(${U('N')},${U('A')},${D4},${U('O')},$B${r})`].forEach((e, j) => { const cel = d.getCell(r, 1 + j); cel.value = c(e); cel.numFmt = [5, 10].includes(j) ? '0.0%' : [3, 4, 7, 8, 9].includes(j) ? '0.0' : '#,##0'; });
      }
      d.addConditionalFormatting({ ref: `F${T0}:F${T0 + NTI - 1}`, rules: [{ type: 'colorScale', cfvo: [{ type: 'num', value: 0 }, { type: 'num', value: 0.25 }], color: [{ argb: 'FFFFFFFF' }, { argb: 'FFF4A5A5' }] }] });
      // pausas por encima del umbral (el motivo se escribe en la hoja PAUSAS, junto a cada pausa)
      d.getCell(`A${TIEM.titulo}`).value = f(`"Pausas de más de "&${B5}&" min en la toma mostrada"`); d.getCell(`A${TIEM.titulo}`).font = { bold: true, size: 12, color: { argb: CARBON } };
      d.getCell(`A${TIEM.nota}`).value = f(`"Mostrando "&MIN(${NPL},MAX(${PA('Q')}))&" de "&MAX(${PA('Q')})&" · el motivo se anota en la hoja PAUSAS (columna Motivo), en la fila que dice la última columna"`); d.getCell(`A${TIEM.nota}`).font = { size: 9, color: { argb: GRIS_TXT } };
      encabezado(d, TIEM.cab, ['Contador', 'Unidad', 'Tipo', 'Desde', 'Hasta', 'Min', '¿Confirma servidor?', 'Entre posiciones', 'Motivo', 'Fila en PAUSAS'], 1);
      for (let i = 1; i <= NPL; i++) {
        const r = L0 + i - 1, n = `$N${r}`, c = (e) => f(`IF(${n}="","",${e})`), ix = (col) => `INDEX(${PA(col)},${n})`;
        d.getCell(r, 14).value = f(`IFERROR(MATCH(${i},${PA('Q')},0),"")`);
        [ix('B'), ix('C'), ix('D'), ix('E'), ix('F'), ix('G'), ix('J'), `${ix('K')}&" → "&${ix('L')}`, `IF(${ix('N')}="","Sin explicar",${ix('N')})`, `${n}+1`].forEach((e, j) => { const cel = d.getCell(r, 1 + j); cel.value = c(e); if (j === 3 || j === 4) cel.numFmt = 'hh:mm'; if (j === 5) cel.numFmt = '0.0'; });
      }
      d.addConditionalFormatting({ ref: `G${L0}:G${TIEM.ultimaFila}`, rules: [{ type: 'cellIs', operator: 'equal', formulae: ['"No"'], style: { fill: relleno('FFF9D2D2'), font: { color: { argb: 'FF791F1F' } } } }] });
    }

    function hojaLinea(wb) {
      const d = wb.addWorksheet('LÍNEA', { views: [{ showGridLines: false }], properties: { tabColor: { argb: 'FFA32D2D' } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      const [s0, s1] = LINEA.slots, n = s1 - s0 + 1;
      d.getColumn(1).width = 24; for (let c = 2; c <= 1 + n; c++) d.getColumn(c).width = 3.4;
      d.getCell('A1').value = 'Línea de tiempo · guardados por franja de 15 minutos (de 06:00 a 22:00)'; d.getCell('A1').font = { size: 14, bold: true, color: { argb: CARBON } };
      d.getCell('A2').value = f(`"Toma: "&TIEMPOS!$D$4&" · número = guardados en esa franja · ROJO = la unidad estaba abierta y no hubo ningún guardado · ${EJ}CONFIDENCIAL"`); d.getCell('A2').font = { italic: true, size: 9, color: { argb: ROJO } };
      d.getCell('A5').value = 'franja n.º'; d.getCell('A5').font = { size: 8, color: { argb: GRIS_TXT } }; d.getCell('A6').value = 'hora'; d.getCell('A6').font = { size: 9, bold: true };
      for (let s = s0; s <= s1; s++) {
        const c = 2 + s - s0; d.getCell(5, c).value = s; d.getCell(5, c).font = { size: 7, color: { argb: GRIS_TXT } };
        const h = d.getCell(6, c); h.value = f(`${letra(c)}5/96`); h.numFmt = 'hh:mm'; h.font = { size: 8, bold: true }; h.alignment = { textRotation: 90, horizontal: 'center' };
      }
      d.getRow(6).height = 36;
      for (let i = 0; i < LINEA.filas; i++) {
        const r = LINEA.fila0 + i, rt = TIEM.tabla.fila0 + i, vacia = `TIEMPOS!$N${rt}=""`;
        d.getCell(r, 1).value = f(`IF(${vacia},"",TIEMPOS!$A${rt}&" · "&TIEMPOS!$B${rt})`); d.getCell(r, 1).font = { size: 10 };
        for (let s = s0; s <= s1; s++) {
          const c = 2 + s - s0, L = letra(c), crit = `${FR('A')},TIEMPOS!$D$4,${FR('C')},TIEMPOS!$B${rt},${FR('E')},${L}$5`, cel = d.getCell(r, c);
          cel.value = f(`IF(${vacia},"",IF(COUNTIFS(${crit})=0,"",SUMIFS(${FR('F')},${crit})))`); cel.font = { size: 8 }; cel.alignment = { horizontal: 'center' }; cel.border = { left: { style: 'hair' }, right: { style: 'hair' }, top: { style: 'hair' }, bottom: { style: 'hair' } };
        }
      }
      const rango = `B${LINEA.fila0}:${letra(1 + n)}${LINEA.ultimaFila}`;
      d.addConditionalFormatting({ ref: rango, rules: [
        { type: 'cellIs', operator: 'equal', formulae: [0], style: { fill: relleno('FFF4A5A5'), font: { color: { argb: 'FF791F1F' }, bold: true } } },
        { type: 'colorScale', cfvo: [{ type: 'num', value: 1 }, { type: 'num', value: 15 }], color: [{ argb: 'FFEAF4E6' }, { argb: 'FF8FC987' }] }] });
    }

    function hojaDashboard(wb) {
      const d = wb.addWorksheet('Dashboard', { views: [{ showGridLines: false }], properties: { tabColor: { argb: AMARILLO } }, pageSetup: { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 } });
      d.getColumn(1).width = 22; for (let c = 2; c <= 12; c++) d.getColumn(c).width = 12;
      d.mergeCells('A1:L1'); d.getCell('A1').value = 'Informe de inventario · Dashboard'; d.getCell('A1').font = { size: 16, bold: true, color: { argb: CARBON } };
      d.mergeCells('A2:L2'); d.getCell('A2').value = EJ + 'BenitoStock · se actualiza al pegar una toma nueva en las hojas del MAESTRO · metas orientativas: cámbialas en la fila 9';
      d.getCell('A2').font = { italic: true, size: 10, color: { argb: GRIS_TXT } };
      // el aviso de arriba (vacío = todo en orden): fórmulas llenas, problemas de Carga o, en «Vista protegida», que hay que habilitar la edición para que Excel calcule
      d.mergeCells('A3:L3');
      const rEstado = `TOMAS!$W$${CAP.fila0}:$W$${CAP.fila0 + CAP.hojas.length - 1}`, LL = `TOMAS!$U$${CAP.fila0 + CAP.hojas.length}`;
      const aviso = `IF(COUNTIF(${rEstado},"LLENA")>0,"⚠ Se llenaron las fórmulas de: "&LEFT(${LL},LEN(${LL})-2)&". Lo que sobra no se cuenta: amplía las fórmulas (ver LÉEME).",IF(${NN}-COUNTIF(${TO('C')},"OK")>0,"⚠ Mira la columna Carga de la hoja TOMAS: hay tomas con datos que faltan, repetidos o fuera de orden.",""))`;
      // sin calcular (Excel en «Vista protegida» no recalcula) solo se ve el resultado guardado: ese texto dice qué hacer; al calcular desaparece
      d.getCell(DASH.banner).value = opciones.vistaProtegida ? { formula: aviso, result: '⚠ Para ver las cifras, abre este archivo en Excel y pulsa «Habilitar edición» (arriba). Hasta entonces el Dashboard se ve vacío.' } : f(aviso);
      d.getCell(DASH.banner).font = { bold: true, size: 11, color: { argb: 'FF791F1F' } }; d.getCell(DASH.banner).alignment = { indent: 1, vertical: 'middle', wrapText: true };
      d.addConditionalFormatting({ ref: 'A3:L3', rules: [{ type: 'expression', formulae: ['LEN($A$3)>0'], style: { fill: relleno('FFFCEBC0') } }] });
      const ult = (c, atras) => `INDEX(${TO(c)},${NN}${atras ? '-' + atras : ''})`, plano = '0.0%';
      const tile = (c1, c2, etiqueta, valor, fmt, delta, fmtD, meta) => {
        for (let r = 4; r <= 7; r++) { d.mergeCells(`${c1}${r}:${c2}${r}`); for (let c = c1.charCodeAt(0); c <= c2.charCodeAt(0); c++) d.getCell(`${String.fromCharCode(c)}${r}`).fill = relleno(GRIS); }
        const a = d.getCell(`${c1}4`), v = d.getCell(`${c1}5`), x = d.getCell(`${c1}6`), m = d.getCell(`${c1}7`);
        a.value = etiqueta; a.font = { size: 10, color: { argb: GRIS_TXT } }; a.alignment = { indent: 1 };
        v.value = valor; v.numFmt = fmt; v.font = { size: 22, bold: true, color: { argb: CARBON } }; v.alignment = { horizontal: 'left', indent: 1 };
        x.value = delta; x.numFmt = fmtD; x.font = { size: 10, bold: true }; x.alignment = { horizontal: 'left', indent: 1 };
        m.value = meta; m.font = { size: 9, color: { argb: GRIS_TXT } }; m.alignment = { indent: 1 }; d.getRow(5).height = 32;
      };
      const dpts = (col, atras) => f(`IF(${NN}<2,"",IFERROR((${ult(col)}-${ult(col, atras || 1)})*100,""))`);
      tile('A', 'C', 'ERI · exactitud de posiciones · última toma', f(`IF(${NN}<1,"",IFERROR(${ult('H')},""))`), plano, dpts('H'), F_DELTA_PTS, f(`"meta "&ROUND($C$9*100,1)&" %"`));
      tile('D', 'F', 'ERU aprox. · exactitud de ubicación', f(`IF(${NN}<1,"",IFERROR(${ult('K')},""))`), plano, dpts('K'), F_DELTA_PTS, f(`"meta "&ROUND($E$9*100,1)&" % · aproximado"`));
      tile('G', 'I', 'Cobertura · posiciones contadas', f(`IF(${NN}<1,"",IFERROR(${ult('L')},""))`), plano, dpts('L'), F_DELTA_PTS, f(`"meta "&ROUND($G$9*100,1)&" %"`));
      tile('J', 'L', 'Diferencia neta · unidades', f(`IF(${NN}<1,"",IFERROR(${ult('R')},""))`), F_DIF, f(`IF(${NN}<1,"",IFERROR("falta "&-${ult('M')}&" · sobra "&${ult('N')}&" u",""))`), '@', f(`IF(${NN}<1,"",IFERROR("Σ |dif| = "&${ult('O')}&" u",""))`));
      for (const [celda, meta] of [['$A$5', '$C$9'], ['$D$5', '$E$9'], ['$G$5', '$G$9']]) d.addConditionalFormatting({ ref: celda.replace(/\$/g, ''), rules: [
        { type: 'expression', formulae: [`AND(ISNUMBER(${celda}),${celda}>=${meta})`], style: { fill: relleno('FFD6EDCF') } },
        { type: 'expression', formulae: [`AND(ISNUMBER(${celda}),${celda}>=${meta}-0.03)`], style: { fill: relleno('FFFCEBC0') } },
        { type: 'expression', formulae: [`ISNUMBER(${celda})`], style: { fill: relleno('FFF9D2D2') } }] });
      d.getCell('A9').value = 'Metas (puedes cambiarlas) →'; d.getCell('A9').font = { bold: true };
      [['B9', 'ERI', 'C9', 0.98], ['D9', 'ERU', 'E9', 0.985], ['F9', 'Cobertura', 'G9', 0.95]].forEach(([l, t, c, val]) => { d.getCell(l).value = t; d.getCell(l).alignment = { horizontal: 'right' }; const x = d.getCell(c); x.value = val; x.numFmt = plano; x.fill = relleno('FFFFF2B3'); x.border = BORDE; x.alignment = { horizontal: 'center' }; x.font = { bold: true }; });
      const seccion = (fila, titulo, nota) => { d.getCell(`A${fila}`).value = titulo; d.getCell(`A${fila}`).font = { bold: true, size: 12, color: { argb: CARBON } }; if (nota) { d.getCell(`D${fila}`).value = nota; d.getCell(`D${fila}`).font = { size: 9, color: { argb: GRIS_TXT } }; } };
      // 1) ERI por toma (las últimas 12)
      seccion(DASH.seccion.tomas, 'Exactitud por toma', 'rojo suave < 95 % < verde desde 98 % · ERI de posiciones con el primer conteo');
      encabezado(d, DASH.seccion.tomas + 1, ['Toma', 'Fecha', 'Posic.', 'Contadas', 'ERI posic.', 'ERI unid.', 'ERU aprox.', 'Cobertura', 'ERI vs anterior'], 1); d.getRow(DASH.seccion.tomas + 1).height = 30;
      for (let i = 1; i <= DASH.tomas.filas; i++) {
        const r = DASH.tomas.fila0 + i - 1, x = `(MAX(0,${NN}-12)+${i})`, fuera = `${x}>${NN}`;
        [['A', 'A'], ['B', 'B'], ['C', 'D'], ['D', 'E'], ['E', 'H'], ['F', 'Q'], ['G', 'K'], ['H', 'L']].forEach(([dst, src]) => { d.getCell(`${dst}${r}`).value = f(`IF(${fuera},"",IFERROR(INDEX(${TO(src)},${x}),""))`); });
        d.getCell(`I${r}`).value = f(`IF(OR(${fuera},${x}<2),"",IFERROR((E${r}-INDEX(${TO('H')},${x}-1))*100,""))`);
        d.getCell(`B${r}`).numFmt = 'dd/mm/yyyy'; d.getCell(`C${r}`).numFmt = '#,##0'; d.getCell(`D${r}`).numFmt = '#,##0'; ['E', 'F', 'G', 'H'].forEach((c) => { d.getCell(`${c}${r}`).numFmt = plano; }); d.getCell(`I${r}`).numFmt = F_DELTA_PTS;
      }
      d.addConditionalFormatting(ESCALA(`E${DASH.tomas.fila0}:E${DASH.tomas.fila0 + DASH.tomas.filas - 1}`)); d.addConditionalFormatting(ESCALA(`G${DASH.tomas.fila0}:G${DASH.tomas.fila0 + DASH.tomas.filas - 1}`));
      // 2) qué mirar primero
      seccion(DASH.seccion.mirar, 'Qué mirar primero');
      const MA0 = DASH.mapa.fila0, MA1 = DASH.ultimaFila, C0 = DASH.codigos.fila0, C1 = C0 + DASH.codigos.filas - 1, J = `$J$${MA0}:$J$${MA1}`, H = `$H$${C0}:$H$${C1}`;
      const linea = (fila, formula) => { d.mergeCells(`A${fila}:L${fila}`); const c = d.getCell(`A${fila}`); c.value = f(formula); c.fill = relleno(GRIS); c.alignment = { indent: 1, vertical: 'middle' }; d.getRow(fila).height = 22; };
      linea(DASH.mirar[0], `IF(COUNT(${J})=0,"Sin datos por pasillo o torre todavía.","Pasillo/torre más flojo en la última toma: "&INDEX($A$${MA0}:$A$${MA1},MATCH(MIN(${J}),${J},0))&" ("&ROUND(MIN(${J})*100,1)&" % de exactitud).")`);
      linea(DASH.mirar[1], `IF(COUNT(${H})=0,"Ningún código repite falta o sobra todavía.","Código que más repite: "&INDEX($A$${C0}:$A$${C1},MATCH(MAX(${H}),${H},0))&" ("&INDEX($F$${C0}:$F$${C1},MATCH(MAX(${H}),${H},0))&", "&MAX(${H})&" tomas).")`);
      linea(DASH.mirar[2], `IF(${NN}<1,"Sin tomas todavía.","Pausas de más de "&TIEMPOS!$B$5&" min en la última toma: "&COUNTIFS(${PA('A')},${ult('A')},${PA('D')},"Pausa",${PA('G')},">"&TIEMPOS!$B$5)&" (detalle en la hoja TIEMPOS).")`);
      // 3) códigos que repiten falta o sobra (la lista sale de las columnas de ayuda de POR CÓDIGO)
      seccion(DASH.seccion.codigos, 'Códigos que repiten falta o sobra', 'rojo = falta, azul = sobra · se listan los 10 primeros que aparecieron');
      encabezado(d, DASH.seccion.codigos + 1, ['Código', 'Tomas contadas', 'Con falta', 'Con sobra', 'Neto', 'Patrón', 'Descripción'], 1); encabezado(d, DASH.seccion.codigos + 1, ['Veces'], 8); d.getRow(DASH.seccion.codigos + 1).height = 30;
      for (let i = 1; i <= DASH.codigos.filas; i++) {
        const r = C0 + i - 1, a = `$A${r}`, lk = `MATCH(${i},${PC('I')},0)`, e = (x) => f(`IF(${a}="","",${x})`);
        d.getCell(`A${r}`).value = f(`IFERROR(INDEX(${PC('B')},${lk}),"")`);
        [['B', `COUNTIFS(${PC('B')},${a})`], ['C', `COUNTIFS(${PC('B')},${a},${PC('E')},"<0")`], ['D', `COUNTIFS(${PC('B')},${a},${PC('E')},">0")`], ['E', `SUMIFS(${PC('E')},${PC('B')},${a})`],
          ['F', `IF(C${r}+D${r}<2,"",IF(D${r}=0,IF(C${r}=B${r},"Falta siempre","Falta repetida"),IF(C${r}=0,IF(D${r}=B${r},"Sobra siempre","Sobra repetida"),"Mixto")))`], ['H', `C${r}+D${r}`]].forEach(([c, x]) => { d.getCell(`${c}${r}`).value = e(x); });
        d.getCell(`G${r}`).value = f(`IFERROR(INDEX(${PC('C')},${lk}),"")`); d.getCell(`E${r}`).numFmt = F_DIF; d.getCell(`H${r}`).font = { color: { argb: GRIS_TXT } };
      }
      d.addConditionalFormatting({ ref: `F${C0}:F${C1}`, rules: [
        { type: 'expression', formulae: [`LEFT($F${C0},5)="Falta"`], style: { fill: relleno('FFF7C1C1'), font: { color: { argb: 'FF791F1F' } } } },
        { type: 'expression', formulae: [`LEFT($F${C0},5)="Sobra"`], style: { fill: relleno('FFB5D4F4'), font: { color: { argb: 'FF0C447C' } } } },
        { type: 'expression', formulae: [`LEFT($F${C0},5)="Mixto"`], style: { fill: relleno('FFE8E6DF'), font: { color: { argb: GRIS_TXT } } } }] });
      // 4) incidencias de la última toma
      seccion(DASH.seccion.incid, 'Incidencias de la última toma', 'detalle en la hoja INCIDENCIAS');
      encabezado(d, DASH.seccion.incid + 1, ['Diferencia grande ⚠', 'Reconteos pedidos', 'Correcciones', 'Lote o vcto. distinto', 'No ajustadas', 'Códigos que repiten'], 1); d.getRow(DASH.seccion.incid + 1).height = 30;
      ['c', 'd', 'e', 'f', 'h'].forEach((t, i) => { const c = d.getCell(DASH.incid.fila, 1 + i); c.value = f(`IF(${NN}<1,"",COUNTIFS(${IN('A')},${ult('A')},${IN('B')},"${t}"))`); c.alignment = { horizontal: 'center' }; });
      d.getCell(DASH.incid.fila, 6).value = f(`COUNT(${H})`); d.getCell(DASH.incid.fila, 6).alignment = { horizontal: 'center' };
      // 5) mapa de calor, al final porque tiene una fila por pasillo o torre: ERI de cada una en las últimas 8 tomas (los inventarios puntuales no entran: no son un pasillo)
      const SM = DASH.seccion.mapa;
      seccion(SM, 'Mapa de calor · ERI por pasillo / torre', f(`"cada columna es una toma (las últimas 8); vacío = esa unidad no se contó"&IF(MAX(${U('AH')})>${NM}," · ⚠ mostrando ${NM} de "&MAX(${U('AH')})&" pasillos/torres","")`));
      encabezado(d, SM + 1, ['Pasillo / torre'], 1);
      for (let j = 1; j <= DASH.mapa.cols; j++) { const c = d.getCell(SM + 1, 1 + j), y = `(MAX(0,${NN}-8)+${j})`; c.value = f(`IF(${y}>${NN},"",INDEX(${TO('B')},${y}))`); c.numFmt = 'dd/mm/yy'; c.font = { bold: true, color: { argb: CARBON } }; c.fill = relleno(AMARILLO); c.alignment = { horizontal: 'center' }; }
      encabezado(d, SM + 1, ['Última toma'], 10);
      for (let i = 1; i <= NM; i++) {
        const r = MA0 + i - 1, a = `$A${r}`;
        d.getCell(`A${r}`).value = f(`IFERROR(INDEX(${U('C')},MATCH(${i},${U('AH')},0)),"")`);
        for (let j = 1; j <= DASH.mapa.cols; j++) {
          const y = `(MAX(0,${NN}-8)+${j})`, id = `INDEX(${TO('A')},${y})`, crit = `${U('A')},${id},${U('C')},${a},${U('O')},""`, c = d.getCell(r, 1 + j);
          c.value = f(`IF(OR(${a}="",${y}>${NN}),"",IFERROR(1-SUMIFS(${U('AE')},${crit})/SUMIFS(${U('N')},${crit}),""))`); c.numFmt = plano; c.alignment = { horizontal: 'center' };
        }
        d.getCell(`J${r}`).value = f(`IF(${a}="","",IFERROR(INDEX($B${r}:$I${r},1,MIN(${NN},8)),""))`); d.getCell(`J${r}`).numFmt = plano; d.getCell(`J${r}`).alignment = { horizontal: 'center' }; d.getCell(`J${r}`).font = { color: { argb: GRIS_TXT } };
      }
      d.addConditionalFormatting(ESCALA(`B${MA0}:I${MA1}`));
    }

    // las 4 hojas con fórmulas, en este orden y ANTES de las hojas de datos
    function hojas(wb) { hojaDashboard(wb); hojaTiempos(wb); hojaLinea(wb); hojaTomas(wb); }

    // después de las hojas de datos (ws = { uni, pa, pc }): la lista de tomas, las columnas de ayuda a la derecha de lo que se pega y la lista de motivos de las pausas
    function completar(wb, ws) {
      wb.definedNames.add(`TOMAS!$A$2:$A$${NT + 1}`, 'ListaTomas');
      ayudas(ws.uni, 33, ['Lista toma', 'Lista unidad', 'Duplicada', 'N.º en toma mostrada'], NU + 1, (r) => [
        `IF($A${r}="","",IF(COUNTIF($A$1:$A${r - 1},$A${r})=0,MAX($AG$1:$AG${r - 1})+1,""))`,
        `IF($C${r}="","",IF(OR($O${r}<>"",$D${r}="puntual"),"",IF(COUNTIFS($C$1:$C${r - 1},$C${r},$O$1:$O${r - 1},"")=0,MAX($AH$1:$AH${r - 1})+1,"")))`,
        `IF($A${r}="","",IF(AND($O${r}="",COUNTIFS($A$1:$A${r - 1},$A${r},$C$1:$C${r - 1},$C${r},$O$1:$O${r - 1},"")>0),"DUPLICADA",""))`,
        `IF($A${r}="","",IF(AND($A${r}=TIEMPOS!$D$4,$O${r}=""),COUNTIFS($A$2:$A${r},TIEMPOS!$D$4,$O$2:$O${r},""),""))`]);
      ayudas(ws.pa, 17, ['N.º lista'], NPA + 1, (r) => [`IF($A${r}="","",IF($A${r}<>TIEMPOS!$D$4,"",IF($G${r}>TIEMPOS!$B$5,COUNTIFS($A$2:$A${r},TIEMPOS!$D$4,$G$2:$G${r},">"&TIEMPOS!$B$5),"")))`]);
      encabezado(ws.pa, 1, ['Motivo', 'Nota'], 14); ws.pa.getColumn(14).width = 24; ws.pa.getColumn(15).width = 34;
      for (const c of [14, 15]) ws.pa.getCell(1, c).fill = relleno('FFFFF2B3');
      ws.pa.dataValidations.add(`N2:N${NPA + 1}`, { type: 'list', allowBlank: true, formulae: ['"Descanso autorizado,Falla de equipo o señal,Buscando producto,Otro,Sin explicar"'], showErrorMessage: true, errorTitle: 'Motivo', error: 'Elige un motivo de la lista.' });
      ayudas(ws.pc, 8, ['Tomas con diferencia', 'N.º lista'], NPC + 1, (r) => [
        `IF($A${r}="","",IF($E${r}=0,"",COUNTIFS(${PC('B')},$B${r},${PC('E')},"<>0")))`,
        `IF($H${r}="","",IF($H${r}<2,"",IF(COUNTIFS($B$1:$B${r - 1},$B${r},$I$1:$I${r - 1},">0")=0,MAX($I$1:$I${r - 1})+1,"")))`]);
    }

    // el texto de la hoja LÉEME del MAESTRO (el primer renglón, «Qué es», lo pone quien llama)
    function leeme() {
      const n = (x) => x.toLocaleString('es-PE');
      return [
        ['Cómo se pega', 'Abre el TOMA, copia las filas de cada hoja (sin la fila de títulos) y pégalas en la primera fila libre de la hoja del MISMO nombre del MAESTRO, de la columna A en adelante: POSICIONES, UNIDADES, INCIDENCIAS, POR CÓDIGO, FRANJAS y PAUSAS. Pega las tomas en orden de fecha, de la más antigua a la más nueva. No pegues sobre las columnas grises ni sobre Motivo y Nota.'],
        ['Aviso de arriba', 'La fila 3 del Dashboard avisa si algo no está bien: que se llenaron las fórmulas de alguna hoja (lo que sobra no se cuenta: mira la tabla «Capacidad de las fórmulas» a la derecha de TOMAS), que alguna toma tiene un problema en la columna Carga, o —si abriste el archivo en «Vista protegida»— que pulses «Habilitar edición» para que Excel calcule. Vacía = todo en orden.'],
        ['Hojas', 'Dashboard: indicadores contra meta, exactitud por toma, qué mirar primero, códigos que repiten, incidencias y, al final, el mapa de calor por pasillo/torre. TIEMPOS y LÍNEA: tiempos por contador (confidenciales). TOMAS: una fila por toma cargada, con su estado de Carga, y la capacidad de las fórmulas. Las demás son los datos pegados.'],
        ['Carga', 'La columna Carga de TOMAS avisa si a una toma le falta una hoja, si pegaste sus unidades dos veces, si POSICIONES no cuadra con UNIDADES o si una toma quedó antes que otra más antigua.'],
        ['ERI y ERU', 'ERI posiciones = posiciones contadas que cuadran ÷ posiciones contadas (primer conteo). ERI unidades = 1 − (suma de |diferencias| ÷ stock del sistema de lo contado). ERU aprox. = 1 − (productos inesperados + posiciones donde el sistema tenía producto y se halló vacío) ÷ contadas: es APROXIMADO porque la app cuenta por posición y no sabe dónde terminó lo que sobra. Cobertura = contadas ÷ posiciones. Metas orientativas (98 %, 98,5 %, 95 %): cámbialas en la fila 9.'],
        ['Tiempos', 'Un guardado es una posición que el contador confirmó. Hueco = más de [umbral] minutos entre dos guardados seguidos de la misma unidad; el arranque (tomó la unidad → primer guardado) también cuenta. «Sin registros» suma huecos y arranque por encima del umbral; «% sin registros» = sin registros ÷ unidad abierta (tomó → último guardado). «Pos/h neto» = posiciones ÷ (abierta − sin registros). Cambia la toma y el umbral en las casillas amarillas de TIEMPOS. La línea de tiempo cubre de 06:00 a 22:00.'],
        ['Inventarios puntuales', 'Los inventarios por código no salen en el mapa de calor por pasillo/torre (no son un pasillo), pero sí en TIEMPOS, LÍNEA y en las cifras de cada toma.'],
        ['Motivo de cada pausa', 'En la hoja PAUSAS, columna Motivo, elige de la lista (Descanso autorizado, Falla de equipo o señal, Buscando producto, Otro, Sin explicar). Lo que dejes vacío figura como «Sin explicar».'],
        ['Límites', 'Un hueco prueba que no hubo guardados, no el porqué: ir al baño o tomar un descanso son necesidades normales, la bandeja Modula puede tardar y buscar un producto lleva tiempo. La app no ve lo que hace la persona fuera de ella, y se puede «burlar» guardando una posición cada pocos minutos: por eso se cruza con la exactitud (% con diferencia y reconteos). Que un código repita no prueba un robo: los errores del WMS dan falsos positivos.'],
        ['Datos personales', 'TIEMPOS, LÍNEA, UNIDADES, FRANJAS y PAUSAS llevan el código de usuario de cada contador (Ley 29733). Avisa a los contadores que se registra la hora de cada guardado y para qué, y acuerda con RR. HH. cómo se usa antes de aplicarlo a medidas disciplinarias.'],
        ['Compatibilidad', 'Solo funciones de Excel antiguo: CONTAR.SI.CONJUNTO, SUMAR.SI.CONJUNTO, SI.ERROR, INDICE, COINCIDIR, MAX, MIN, SI, Y, O, REDONDEAR. No usa MAXIFS, BUSCARX, TEXTO ni matrices dinámicas.'],
        ['Tamaño de las fórmulas', `Las fórmulas miran hasta ${n(NP)} filas de datos de POSICIONES, ${n(NU)} de UNIDADES, ${n(NPC)} de POR CÓDIGO, ${n(NF)} de FRANJAS, ${n(NPA)} de PAUSAS y ${n(NI)} de INCIDENCIAS (≈ ${NT} tomas). TIEMPOS y LÍNEA muestran ${NTI} unidades y ${NPL} pausas; el mapa de calor, ${NM} pasillos o torres. Si algo se llena, el aviso de arriba lo dice y hay que ampliar las fórmulas.`],
      ];
    }

    return { tam: { ...tam, NTI, NPL, NM }, DASH, TIEM, LINEA, CAP, hojas, completar, leeme };
  }

  return { crear, TAM_EJEMPLO, TAM_REAL, COLS_ESPERADAS, verificar };
}));
