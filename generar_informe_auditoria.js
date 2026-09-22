const {
  Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun,
  AlignmentType, WidthType, BorderStyle, ShadingType, VerticalAlign,
  TableLayoutType, Header, Footer, PageNumber, PageBreak, HeadingLevel
} = require('docx');
const fs = require('fs');

// ── Paleta institucional Banco Popular ────────────────────────────────────
const ORANGE       = 'F47920';
const ORANGE_DARK  = 'B85810';
const ORANGE_LIGHT = 'FDEEDE';
const PURPLE       = '7C51A1';
const PURPLE_DARK  = '5A3A7A';
const PURPLE_LIGHT = 'EDE8F4';
const WHITE        = 'FFFFFF';
const GRAY_LIGHT   = 'F4F4F4';
const GRAY_MID     = 'CCCCCC';
const TEXT_DARK    = '1A1A1A';
const GREEN_BG     = 'D4EDDA';
const GREEN_TEXT   = '155724';
const AMBER_BG     = 'FFF3CD';
const AMBER_TEXT   = '7B5800';
const RED_BG       = 'F8D7DA';
const RED_TEXT     = '721C24';

// ── Tipografía ─────────────────────────────────────────────────────────────
const FONT  = 'Arial';
const SZ    = 22;   // 11pt
const SZ_SM = 18;   // 9pt
const SZ_XS = 16;   // 8pt

// ── Helpers ────────────────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({
    text,
    font: FONT,
    size: opts.size || SZ,
    bold: opts.bold || false,
    color: opts.color || TEXT_DARK,
    italics: opts.italic || false,
  });
}

function para(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 100 },
    indent: opts.indent ? { left: opts.indent } : undefined,
    border: opts.border || undefined,
  });
}

function gap(n = 80) {
  return para([run('')], { after: n });
}

function pageBreakPara() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Línea separadora naranja
function divider(color = ORANGE) {
  return new Paragraph({
    children: [run('')],
    border: { bottom: { style: BorderStyle.SINGLE, size: 10, color, space: 1 } },
    spacing: { before: 40, after: 160 },
  });
}

// Celda de tabla
function cell(text, opts = {}) {
  return new TableCell({
    children: [new Paragraph({
      children: [run(text, {
        size: opts.size || SZ_SM,
        bold: opts.bold || false,
        color: opts.textColor || TEXT_DARK,
        italic: opts.italic || false,
      })],
      alignment: opts.align || AlignmentType.LEFT,
      spacing: { before: 40, after: 40 },
    })],
    width: { size: opts.w, type: WidthType.DXA },
    shading: opts.bg ? { type: ShadingType.CLEAR, fill: opts.bg, color: 'auto' } : undefined,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    columnSpan: opts.span || undefined,
    rowSpan: opts.rowSpan || undefined,
  });
}

function hCell(text, w, color = ORANGE) {
  return cell(text, { w, bold: true, bg: color, textColor: WHITE, size: SZ_SM, align: AlignmentType.CENTER });
}

function statusCell(status, w) {
  const MAP = {
    'CONFORME':  { bg: GREEN_BG,   color: GREEN_TEXT,  label: '✔ CONFORME'  },
    'PENDIENTE': { bg: AMBER_BG,   color: AMBER_TEXT,  label: '⏳ PENDIENTE' },
    'HALLAZGO':  { bg: RED_BG,     color: RED_TEXT,    label: '● HALLAZGO'  },
  };
  const s = MAP[status] || MAP['PENDIENTE'];
  return cell(s.label, { w, bg: s.bg, textColor: s.color, size: SZ_SM, align: AlignmentType.CENTER, bold: true });
}

const TABLE_BORDERS = {
  top:     { style: BorderStyle.SINGLE, size: 6, color: ORANGE_DARK },
  bottom:  { style: BorderStyle.SINGLE, size: 6, color: ORANGE_DARK },
  left:    { style: BorderStyle.SINGLE, size: 6, color: ORANGE_DARK },
  right:   { style: BorderStyle.SINGLE, size: 6, color: ORANGE_DARK },
  insideH: { style: BorderStyle.SINGLE, size: 1, color: GRAY_MID },
  insideV: { style: BorderStyle.SINGLE, size: 1, color: GRAY_MID },
};

function sectionTitle(text, color = ORANGE) {
  return para([run(text, { size: 28, bold: true, color })], { after: 100 });
}

function sectionDesc(label, content) {
  return para([
    run(label + ': ', { bold: true }),
    run(content),
  ], { after: 80 });
}

// Nota al pie de bloque con borde izquierdo morado
function nota(text) {
  return new Paragraph({
    children: [run(text, { size: SZ_SM })],
    border: { left: { style: BorderStyle.SINGLE, size: 20, color: PURPLE, space: 10 } },
    indent: { left: 280 },
    spacing: { before: 80, after: 160 },
  });
}

// Página separadora de PARTE
function parteSeparator(numRomano, titulo, descripcion) {
  return [
    pageBreakPara(),
    gap(360),
    new Paragraph({
      children: [run('PARTE ' + numRomano, { size: 60, bold: true, color: ORANGE })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
    }),
    new Paragraph({
      children: [run(titulo, { size: 32, bold: true, color: PURPLE_DARK })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 160 },
    }),
    new Paragraph({
      children: [run(descripcion, { size: SZ, color: '555555', italics: true })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 280 },
    }),
    divider(ORANGE),
    pageBreakPara(),
  ];
}

// Fila del índice
function indiceRow(numero, titulo, subtitulo) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [600, 8760],
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideH: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      insideV: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    },
    rows: [new TableRow({ children: [
      new TableCell({
        children: [para([run(numero, { bold: true, color: ORANGE, size: SZ })], { align: AlignmentType.RIGHT })],
        margins: { top: 40, bottom: 0, left: 0, right: 120 },
      }),
      new TableCell({
        children: [
          para([run(titulo, { bold: !!subtitulo, size: SZ })]),
          ...(subtitulo ? [para([run(subtitulo, { size: SZ_SM, color: '777777', italics: true })], { before: 0, after: 40 })] : []),
        ],
        margins: { top: 40, bottom: 0, left: 0, right: 0 },
      }),
    ]})],
  });
}


// BLOQUE CEL-03
const CEL03_COLS = [700, 2300, 2000, 1700, 900, 1760]; // sum = 9360
const cel03Rules = [
  { c:'JUI-001', estado:'Desistimiento total S.E.P.',       causa:'—',                       resultado:'Aceptable sin afectación',  e:'CONFORME', o:'Estado cerrado admisible. No activa causa CEL-02. Admisible en todos los segmentos.' },
  { c:'JUI-002', estado:'Satisfacción extraprocesal S.E.P.',causa:'—',                       resultado:'Aceptable sin afectación',  e:'CONFORME', o:'Estado cerrado admisible. No activa causa CEL-02. Admisible en todos los segmentos.' },
  { c:'JUI-003', estado:'Terminado',                        causa:'—',                       resultado:'Aceptable sin afectación',  e:'CONFORME', o:'Estado cerrado admisible. No activa causa CEL-02. Admisible en todos los segmentos.' },
  { c:'JUI-004', estado:'Trámite en etapa de ejecución',    causa:'JUICIO_ACTIVO_PRESENTE',  resultado:'No aprobación automática', e:'CONFORME', o:'Estado activo. Activa JUICIO_ACTIVO_PRESENTE. Alineado a ALCO-GEN-006 (no se admiten juicios activos).' },
  { c:'JUI-005', estado:'En trámite',                       causa:'JUICIO_ACTIVO_PRESENTE',  resultado:'No aprobación automática', e:'CONFORME', o:'Estado activo. Activa JUICIO_ACTIVO_PRESENTE. Alineado a ALCO-GEN-006.' },
  { c:'JUI-006', estado:'Solic. Parte actora',              causa:'JUICIO_ACTIVO_PRESENTE',  resultado:'No aprobación automática', e:'CONFORME', o:'Estado activo. Activa JUICIO_ACTIVO_PRESENTE. Alineado a ALCO-GEN-006.' },
  { c:'JUI-007', estado:'Inactivo',                         causa:'DATOS_JUICIOS_INCOMPLETOS',resultado:'Incidencia técnica',       e:'CONFORME', o:'Estado técnico ambiguo. Reprocesable cuando se clarifique la fuente. No se trata como cerrado ni como activo.' },
  { c:'JUI-008', estado:'Deshabilitar carpeta',             causa:'DATOS_JUICIOS_INCOMPLETOS',resultado:'Incidencia técnica',       e:'CONFORME', o:'Estado técnico ambiguo. Reprocesable cuando se clarifique la fuente.' },
  { c:'JUI-999', estado:'Cualquier otro estado',            causa:'DATOS_JUICIOS_INCOMPLETOS',resultado:'Incidencia técnica',       e:'CONFORME', o:'Guardia de catálogo: cualquier estado no enumerado queda como incidencia técnica. Protección ante estados futuros desconocidos.' },
];

function makeCEL03Table() {
  const cols = CEL03_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]),  hCell('Estado de juicio', cols[1]),
        hCell('Causa CEL-02', cols[2]), hCell('Resultado', cols[3]),
        hCell('Estado', cols[4]), hCell('Observación', cols[5]),
      ]}),
      ...cel03Rules.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        return new TableRow({ children: [
          cell(r.c,         { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.estado,    { w: cols[1], bg, size: SZ_SM }),
          cell(r.causa,     { w: cols[2], bg, size: SZ_SM, italic: r.causa === '—' }),
          cell(r.resultado, { w: cols[3], bg, size: SZ_SM }),
          statusCell(r.e,    cols[4]),
          cell(r.o,         { w: cols[5], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

function makeCEL03ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Entradas revisadas', cols[1], PURPLE),
        hCell('Conformes', cols[2], PURPLE), hCell('Hallazgos', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-03 — Catálogo de Juicios (versión ALCO-2026.08.17.1)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('9',  { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('9',  { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('0',  { w: cols[3], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

// ── ANCHO DE PÁGINA ────────────────────────────────────────────────────────
// Letter con márgenes 1": usable = 12240 - 2*1440 = 9360 DXA
const TW = 9360;

// ─────────────────────────────────────────────────────────────────────────
// DATOS
// ─────────────────────────────────────────────────────────────────────────

// BLOQUE 1
const B1_COLS = [980, 1020, 440, 3480, 1080, 2360]; // sum = 9360
const b1Rules = [
  { c:'REGLA-001', t:'No cliente BP', s:'A', d:'Asalariado · 19-51 años · ≤₡40M · Antigüedad ≥48m · Atraso 0-30d · CPH=1 · Score zona blanca · Tasa 12.18pp · Plazo ≤108m', e:'CONFORME', o:'P-001: corte score (181) pendiente respaldo normativo. P-002: SME ₡323,346 pendiente decreto.' },
  { c:'REGLA-002', t:'No cliente BP', s:'A', d:'Asalariado · 51-65 años · ≤₡30M · Resto de condiciones igual a REGLA-001', e:'CONFORME', o:'P-001 y P-002 aplican.' },
  { c:'REGLA-003', t:'No cliente BP', s:'B', d:'Asalariado · 19-65 años · ≤₡30M · Antigüedad 12-48m · Atraso 0d · Tasa 13.13pp', e:'CONFORME', o:'' },
  { c:'REGLA-004', t:'Cliente BP',    s:'A', d:'Asalariado deudor BP · 19-51 años · ≤₡40M · Antigüedad ≥48m · CPH=1 · Tasa 12.18pp · Plazo ≤108m', e:'CONFORME', o:'' },
  { c:'REGLA-005', t:'Cliente BP',    s:'A', d:'Ingresos propios deudor BP · 19-51 años · ≤₡40M · Antigüedad ≥48m · CPH=1 · Tasa 12.18pp · Plazo ≤108m', e:'CONFORME', o:'Ingresos propios aplica exclusivamente a deudores BP (ALCO).' },
  { c:'REGLA-006', t:'Cliente BP',    s:'B', d:'Asalariado deudor BP · 19-65 años · ≤₡30M · Antigüedad 12-48m · 0 días atraso · Tasa 13.13pp', e:'CONFORME', o:'' },
  { c:'REGLA-007', t:'Cliente BP',    s:'B', d:'Ingresos propios deudor BP · 19-65 años · ≤₡30M · Antigüedad 12-48m · 0 días atraso · Tasa 13.13pp', e:'CONFORME', o:'Ingresos propios aplica exclusivamente a deudores BP (ALCO).' },
  { c:'REGLA-008', t:'Cliente BP',    s:'A', d:'Pensionado · 19-65 años · ≤₡20M · Plazo ≤60m · Requiere BP Salario · Tasa 12.18pp', e:'CONFORME', o:'P-002 (SME) aplica. Pensionados exclusivamente para clientes de crédito BP; no clientes pensionados fuera del alcance del ALCO.' },
  { c:'REGLA-009', t:'Cliente BP',    s:'B', d:'Pensionado · 19-65 años · ≤₡15M · Plazo ≤60m · Requiere BP Salario · Tasa 13.13pp', e:'CONFORME', o:'Pensionados exclusivamente para clientes de crédito BP; no clientes pensionados quedan fuera del alcance del ALCO por diseño.' },
  { c:'REGLA-010', t:'Cliente BP',    s:'A', d:'Pensionado · 65-70 años · Requiere Póliza Obligatoria · Tasa 12.18pp', e:'CONFORME', o:'' },
  { c:'REGLA-011', t:'Cliente BP',    s:'B', d:'Pensionado · 65-70 años · Requiere Póliza Obligatoria · Tasa 13.13pp', e:'CONFORME', o:'' },
  { c:'REGLA-012', t:'Cliente BP',    s:'C', d:'Historial CPH ≤1.25 · ≤₡25M · Tasa 15.80pp · Garantía pagaré sin fiador', e:'CONFORME', o:'Monto ₡25M alineado al texto del acuerdo ALCO (prevalece sobre tabla pág. 2).' },
  { c:'REGLA-013', t:'Cliente BP',    s:'D', d:'Atraso 1-30 días · ≤₡25M · Tasa 17.50pp · Garantía pagaré sin fiador', e:'CONFORME', o:'Tasa 17.50pp confirmada físicamente contra ALCO. Monto ₡25M alineado al texto.' },
];

// BLOQUE 2
const B2_COLS = [1100, 3800, 1080, 3380]; // sum = 9360
const b2Rules = [
  { c:'ALCO-GEN-001', d:'BP Salario para pensionados — Todo pensionado aprobado debe autorizar BP Salario.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-002', d:'Autorización CIC — El CIC debe estar autorizado por el cliente antes del análisis.', e:'CONFORME', o:'CEL-27 operativo: WorkId por operación, EstadoCIC y AutorizacionUrl confirman el flujo de autorización. Operador documental SUGEF (Azure 1.6) activo con 69 consultas CIC preservadas; comparación automática de identidad aún deshabilitada; todas las decisiones requieren confirmación humana. Pendiente prueba de aceptación formal en expediente real. → P-005.' },
  { c:'ALCO-GEN-003', d:'Capacidad de pago — Sumatoria de cuotas con nueva solicitud ≤ sumatoria CIC al momento de aprobación.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-004', d:'Ingreso superior a cuotas — Ingresos del cliente > sumatoria de cuotas CIC.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-005', d:'Garantía — Pagaré sin fiador en todos los segmentos.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-006', d:'Juicios y embargos — No se admiten juicios activos ni embargos en Protectora.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-007', d:'Referencias comerciales — Se evalúan conforme a guía general de lineamientos.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-008', d:'Vigencia del acta — 60 días naturales desde emisión: 45 días para venta + 15 días para formalización.', e:'CONFORME', o:'La propuesta original contemplaba 3 meses; el Acuerdo ALCO estableció 60 días. Matriz correctamente actualizada.' },
  { c:'ALCO-GEN-009', d:'CIC más reciente — El cálculo se realiza con el CIC más reciente por DACD.', e:'CONFORME', o:'CEL-27 campos AprobadoUtc y EstadoActualizadoUtc confirman la trazabilidad temporal del CIC más reciente por operación. CONFORME reforzado con evidencia arquitectural CEA-02.' },
  { c:'ALCO-GEN-010', d:'No reevaluación — Durante vigencia del acta no se recalculan variables ya aprobadas.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-011', d:'Sin revisión posterior — Soporte al Crédito no revisa elementos del acta durante su vigencia; solo condiciones pendientes en aprobaciones condicionadas.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-012', d:'Modificación de operaciones — A solicitud del cliente se pueden modificar si nueva sumatoria de cuotas no supera la aprobada; si aumenta, vuelve a ventas.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-013', d:'Operaciones a cancelar — La orden de giro define operaciones a cancelar y atiende diferencias durante formalización.', e:'CONFORME', o:'' },
  { c:'ALCO-GEN-014', d:'Score interno — Aplica únicamente a no clientes BP; debe ubicarse en zona blanca. El corte numérico se administra en el gobierno de Score Interno.', e:'CONFORME', o:'P-001: corte 181 no referenciado explícitamente en ALCO; requiere documento de gobierno de Score Interno.' },
  { c:'ALCO-GEN-015', d:'Monto C y D — Máximo ₡25.000.000; el acuerdo sustituye el monto de la tabla visual de pág. 2.', e:'CONFORME', o:'Aclaratorio explícito incluido en la matriz. Correcto.' },
];

// BLOQUE 3
const B3_COLS = [1100, 4000, 1080, 3180]; // sum = 9360
const b3Rules = [
  { c:'ALCO-GOB-001', d:'Banca Digital coordina con DACD para el desarrollo y supervisión del mecanismo de aprobación automática.', e:'CONFORME', o:'' },
  { c:'ALCO-GOB-002', d:'DACD desarrolla y supervisa el mecanismo de aprobación masiva automática.', e:'CONFORME', o:'' },
  { c:'ALCO-GOB-003', d:'Banca Digital y Soporte al Negocio deben elaborar una guía conjunta de uso del mecanismo.', e:'HALLAZGO', o:'Guía conjunta NO implementada a la fecha. Obligación directa del ALCO.' },
  { c:'ALCO-GOB-004', d:'Subgerencia General de Negocios puede ajustar aspectos no contemplados en el ALCO sin requerir nueva sesión.', e:'CONFORME', o:'' },
  { c:'ALCO-GOB-005', d:'Soporte al Negocio define el mecanismo para identificar la cartera de no clientes BP elegibles.', e:'HALLAZGO', o:'Mecanismo de identificación de cartera NO definido por Soporte al Negocio.' },
  { c:'ALCO-GOB-006', d:'Banca Digital selecciona las bases de bajo riesgo para los no clientes BP.', e:'CONFORME', o:'' },
  { c:'ALCO-GOB-007', d:'Canal exclusivo del producto: Banca Digital.', e:'CONFORME', o:'' },
  { c:'ALCO-GOB-008', d:'Banca Digital debe implementar las instrucciones del DOCCORP-0719-2026 (actualización de información de clientes).', e:'HALLAZGO', o:'DOCCORP-0719-2026 NO procedimentado ni implementado en Banca Digital.' },
  { c:'ALCO-GOB-009', d:'Verificación de listas nacionales e internacionales antes de aprobación.', e:'CONFORME', o:'Verificación vigente en proceso. Nota de arquitectura: la consulta de listas no aparece como tabla explícita en CEA-02; se asume flujo externo integrado. Ver P-003.' },
  { c:'ALCO-GOB-010', d:'Verificación PCSC antes de la formalización de la operación.', e:'CONFORME', o:'PCSC no aparece como tabla explícita en CEA-02; la verificación opera fuera del modelo de datos documentado (posiblemente T24/SUGEF). Solicitar aclaración arquitectural a Jefatura. → P-003.' },
  { c:'ALCO-GOB-011', d:'Si el cliente tiene condición pendiente de PCSC, se verifica actualización antes de formalizar.', e:'CONFORME', o:'' },
];

// BLOQUE 4
const B4_COLS = [1100, 4200, 1080, 2980]; // sum = 9360
const b4Rules = [
  { c:'ALCO-ACT-001', d:'DACD emite el acta general que certifica el proceso de aprobación automática.', e:'CONFORME', o:'Proceso en fase de pruebas. Acta general generada en marco de validación del mecanismo.' },
  { c:'ALCO-ACT-002', d:'Cobertura del acta general: certifica análisis automático de todos los canales, condiciones y requisitos aprobados.', e:'CONFORME', o:'CEA-02 Capa 4 (CEL-15 Actas generales): confirma cobertura total del proceso. Relación CEL-10 → CEL-15 (1:1 por versión) garantiza trazabilidad operación-acta. CONFORME reforzado.' },
  { c:'ALCO-ACT-003', d:'Contenido mínimo del acta general: reglas de negocio, fuentes de consulta, versión del proceso, periodo certificado, controles ejecutados y ubicación de evidencia.', e:'CONFORME', o:'CEA-02: SHA-256 + URL como mecanismo de integridad del acta. CEL-27 campo HuellaOperador provee capa adicional de trazabilidad por operador documental. Cumplimiento fortalecido con doble cadena de integridad.' },
  { c:'ALCO-ACT-004', d:'Revisión trimestral — DACD revisa trimestralmente el proceso, cumplimiento de reglas y calidad de información.', e:'PENDIENTE', o:'Producto en pruebas desde agosto 2026; aún no ha transcurrido el primer trimestre de operación.' },
  { c:'ALCO-ACT-005', d:'DACD emite un acta individual para cada caso aprobado.', e:'CONFORME', o:'Actas individuales generadas durante fase de pruebas.' },
  { c:'ALCO-ACT-006', d:'Aprobación sin condiciones: cumple todos los requisitos sin condiciones pendientes.', e:'CONFORME', o:'' },
  { c:'ALCO-ACT-007', d:'Aprobación condicionada: cumple criterios crediticios pero debe subsanar condición expresa antes de formalizar (PCSC, T24, póliza u otra subsanable).', e:'CONFORME', o:'Riesgo derivado: gestión de condición PCSC pendiente de estandarizar (vinculado a hallazgo GOB-008).' },
  { c:'ALCO-ACT-008', d:'El acta individual contiene todas las operaciones consideradas para sumatoria de cuotas y el origen del cálculo.', e:'CONFORME', o:'' },
  { c:'ALCO-GIR-001', d:'La orden de giro detalla las operaciones finalmente acordadas con el cliente.', e:'CONFORME', o:'' },
  { c:'ALCO-GIR-002', d:'La orden de giro debe ser firmada por el ejecutivo a cargo.', e:'CONFORME', o:'Modalidad esperada: firma digital. Pendiente confirmar en expedientes de prueba.' },
  { c:'ALCO-GIR-003', d:'La orden de giro contiene el detalle de lo negociado y otras observaciones.', e:'CONFORME', o:'' },
];

// PENDIENTES GLOBALES
const pendientes = [
  { id:'P-001', blq:'B1, B2', desc:'Corte de Score Interno (181) — valor no referenciado en el ALCO. Evidencia CEL-02: causa SCORE_INTERNO_15V_NO_APROBADO (versión CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181, vigente 25/08/2026) documenta el corte de 181 pts y excluye zona gris 161-180. Evidencia CEL-05: instrumento formal identificado — JDN-5824-Acd-380-2021-Art-9 aprueba el corte 181 (Consumo 181-200; zona gris 161-180). El puntaje crudo se conserva en PIB para auditoría; no se expone en el modelo semántico por diseño. Pendiente: verificar que JDN-5824 sigue vigente y aplica específicamente al producto Cero Estrés en su versión 2026. La propuesta ALCO-GOB-004 queda como respaldo alternativo si el JDN no cubre el alcance requerido.', imp:'Medio' },
  { id:'P-002', blq:'B1',     desc:'SME ₡323,346 — monto no referenciado en el ALCO. Requiere decreto o normativa interna de respaldo. Propuesta a Jefatura: formalizar la referencia SME mediante ajuste de la Subgerencia General de Negocios (ALCO-GOB-004), sin requerir nueva sesión ALCO.', imp:'Medio' },
  { id:'P-003', blq:'B3',     desc:'La consulta PCSC se evalúa en DIRBDLEADS y su resultado se traslada al motor como parte de la evaluación previa. Si el cliente no cumple, CEL-02 registra la causa LISTAS_O_PCSC_NO_CONFORME y CEL-10 recibe el rechazo. Si cumple, no queda traza explícita en el modelo de datos — no existe tabla dedicada que persista el estado PCSC por cliente. Consulta pendiente a Jefatura: ¿se puede crear una tabla que referencie el estado PCSC por cliente, de manera que quede claro dónde vive esa capa de información dentro del modelo CEA-02?', imp:'Medio' },
  { id:'P-005', blq:'B2',     desc:'Según CEA-02, Azure 1.6 y el Operador documental SUGEF están publicados e instalados. Se verificaron: clasificación sugerida, encuadre y orientación de ambas caras de la cédula, firma completa, persistencia y orden por fecha de correo. Los flujos de reenvío/rechazo y CEL-27 están activos; se conservaron configuración y 69 trabajos CIC durante la instalación. Las pruebas realizadas usaron documentos sintéticos, sin avisos a clientes ni consultas CIC reales. Toda decisión final requiere confirmación humana; la comparación automática de identidad por firma continúa deshabilitada. Pendiente: prueba de aceptación del revisor con un expediente real antes de la operación plena.', imp:'Medio' },
  { id:'P-ACT-004', blq:'B4', desc:'Primer informe trimestral de DACD — no aplica aún. Verificar cuando transcurra el primer trimestre de operación.', imp:'Bajo' },
  { id:'P-GIR-002', blq:'B4', desc:'Firma digital en orden de giro — modalidad esperada confirmada informalmente. Pendiente verificar evidencia de firma digital en expedientes de prueba. Propuesta a Jefatura: formalizar la modalidad de firma mediante ajuste de la Subgerencia General de Negocios (ALCO-GOB-004), sin requerir nueva sesión ALCO.', imp:'Bajo' },
  { id:'P-CEL05-001', blq:'CEL-05', desc:'Dos puntos de mantenimiento en el catálogo de modelos: (1) la versión documentada en CEL-05 es V1.0-DRAFT, pero la versión de producción vigente desde 25/08/2026 es V1.1-SOLO-NO-CLIENTE-CORTE181 (referenciada en CEL-02) — el catálogo debe actualizarse para reflejar la versión activa; (2) el estado del modelo aparece como VALIDACION_TECNICA_MODO_SOMBRA, pero la propia entrada registra Decisión habilitada = Verdadero y las observaciones confirman que el modelo salió del modo sombra al registrar el corte aprobado — corresponde actualizar el estado a producción. Ambas son acciones de mantenimiento del Centro de Auditoría; consulta enviada a Jefatura.', imp:'Bajo' },
  { id:'P-CEL04-001', blq:'CEL-04', desc:'CE-RES-02 (Aprobado Condicionado) incluye "Operaciones en atraso superior al permitido" como condición subsanable. La segmentación B1 evalúa el historial SUGEF del cliente al momento del análisis (atraso máximo: 30 días para Seg.B; 0 días para Seg.A/C). Si un cliente pasa segmentación, su perfil SUGEF cumple ese criterio. La condición de CE-RES-02 opera en una capa distinta y posterior: dentro del período de vigencia del acta (60 días = 45 venta + 15 formalización), las operaciones específicas negociadas para cancelar con Cero Estrés pueden acumular atraso superior al umbral antes de formalizarse. En ese caso, la aprobación condicionada aplica porque el propio giro de Cero Estrés resuelve la condición — es la finalidad del producto. Esto es consistente con ALCO-GEN-013 (la orden de giro define las operaciones a cancelar y atiende diferencias durante la formalización). Interpretación verificada con Jefatura: correcta. Pendiente: confirmar formalmente en el expediente del producto para cierre de este ítem.', imp:'Bajo' },
  { id:'P-CEL06-001', blq:'CEL-06', desc:'Origen de fondos (N_ORI_FND_G_CM): cobertura del 17.34% en la muestra de evaluación (02/09/2026, n=406). El 82.66% de los casos no tiene dato y cae en el grupo MISSING del score (G08 = ADMON_PUBLICA/INMOBILIARIAS = 17 pts, rango medio de la variable). Confirmado en CEL-08: todos los leads del lote 17-18/09/2026 reciben MISSING = 17 pts para esta variable. Pendiente: confirmar con Jefatura si se contempla enriquecimiento de esta variable en próximas versiones del modelo o si la cobertura baja es esperada por diseño del segmento objetivo.', imp:'Medio' },
  { id:'P-CEL07-001', blq:'CEL-07', desc:'Dos observaciones de diseño documentadas para trazabilidad: (1) B_PAGOPLANILLA es el único grupo en el catálogo con puntuación negativa (0+MISSING = -4 pts) — el modelo penaliza activamente la ausencia de pago por planilla, lo que lo convierte en el predictor con mayor diferencial de puntuación; (2) PIB BCCR MISSING = 16 pts, superior a todos los grupos observados (11-12 pts) — el caso sin dato del PIB favorece al cliente. Ambos diseños son consistentes con la lógica WOE y el instrumento JDN-5824-Acd-380-2021-Art-9, pero deben quedar documentados para trazabilidad del modelo en auditorías futuras.', imp:'Bajo' },
  { id:'P-CEL08-001', blq:'CEL-08', desc:'CUOTAS_BP_SIN_CAMPO_CRUDO_EN_EXPEDIENTE: el valor bruto de la variable I_CUOTAS_BP (Cuotas BP) no está expuesto en el expediente DIRBDLEADS para ningún lead del lote evaluado (17-18/09/2026). El motor aplica MISSING = 11 pts (grupo mínimo, G01) en el 100% de los casos revisados. Pendiente: confirmar con Jefatura si esta ausencia es diseño (el campo se calcula internamente y no se almacena en el expediente) o si representa una brecha de datos que afecta la trazabilidad auditable del puntaje final.', imp:'Medio' },
  { id:'P-CEL09-001', blq:'CEL-09', desc:'CEL-09-EvidenciasDelScore no pudo revisarse directamente: la lista en SharePoint supera el umbral de vista de 5,000 elementos (Error de representación desconocido, Id. correlación: 21733ca2-e0e7-f000-60cf-399a9ec3c889, 18/09/2026). El volumen indica operación activa. La operatividad queda respaldada indirectamente por CEL-08, que confirma evaluaciones con Estado de conciliación = SCORE_OFICIAL_CEF_CEL. Ruta de resolución: crear vista filtrada e indexada por fecha o estado, o exportar subconjunto vía Power Automate para revisión auditable. No requiere aprobación adicional.', imp:'Bajo' },
  { id:'P-CEL21-001', blq:'CEL-21', desc:'Catálogo de equivalencias de ocupación CREDID (OCUP-CREDID-V1.0-20260907, 160 entradas, vigente 07/09/2026): el campo Estado de validación de las 160 entradas indica "Vigente para operación; pendiente de visto bueno formal de Riesgos". El catálogo opera en producción y afecta directamente el cálculo de la variable Ocupación laboral del score (N_OCU_LAB_G_CM, hasta 21 pts), confirmado en CEL-08. Pendiente: obtener el visto bueno formal de la Dirección de Riesgos para completar la cadena de gobernanza de la variable de ocupación en el modelo de score.', imp:'Medio' },
  { id:'P-CEL10-001', blq:'CEL-10', desc:'423 registros en CEL-10 sin versión de matriz de reglas, con nomenclatura de resultados distinta a CEL-04 ("Rechazado", "Aprobado condicionado", "Aprobado final", "Incidencia técnica"). Periodo de evaluación: 23/08/2026–29/08/2026 (entre la firma del ALCO 17/08/2026 y el despliegue de CEL-01-PRODUCCION-V1.0). Requiere confirmación de Jefatura: ¿corresponden a evaluaciones pre-producción / pruebas de fase que deben segregarse del expediente de auditoría ALCO, o deben integrarse bajo un registro propio con su versión de reglas?', imp:'Medio' },
  { id:'P-CEL10-002', blq:'CEL-10', desc:'439 registros evaluados bajo versiones CE-REGLAS-ALCO-2026-V2.x (V2.3: 429 · V2.0: 8 · V2.1: 2) no documentadas en CEL-05 ni referenciadas en el Acuerdo ALCO. Resultados: Aprobado 119 · Denegado 278 · Falta Información 28 · Aprobado Condicionado 4. Ruta de resolución: confirmar con Jefatura cuál es la relación entre CE-REGLAS-ALCO-2026-V2.x y CEL-01-PRODUCCION-V1.0 (¿versiones intermedias de desarrollo que quedaron registradas en producción?, ¿ramificación paralela del producto?), y si deben incorporarse en el catálogo CEL-05.', imp:'Medio' },
  { id:'P-CEL11-001', blq:'CEL-11', desc:'CE-REGLAS-ALCO-2026-V2.4 identificada en CEL-11 (375 registros) como versión de matriz no documentada en CEL-05. Todos los registros corresponden al flujo "CEF-2-RECUPERACION-CEL11-2026.09.01.1-MATRIZ-V2.4", fechas 24–29/08/2026 (lote de recuperación pre-ALCO, consistente con P-CEL10-001). Extiende P-CEL10-002: el universo de versiones CE-REGLAS-ALCO-2026-V2.x no documentadas comprende ahora V2.0 (8 registros en CEL-10), V2.1 (2), V2.3 (429 en CEL-10 · 414 en CEL-11) y V2.4 (375 en CEL-11), totalizando más de 800 registros evaluados bajo versiones sin respaldo en CEL-05. Ruta de resolución: idéntica a P-CEL10-002 — confirmar con Jefatura la relación de estas versiones con CEL-01-PRODUCCION-V1.0 y decidir si se incorporan en CEL-05.', imp:'Medio' },
  { id:'P-CEL11-002', blq:'CEL-11', desc:'Se identificaron 16 versiones distintas del flujo de evaluación en CEL-11 Vista 2 (2,379 registros). Entre ellas figura una versión de ambiente QA: CE-QA-REGLAS-V2-2026.08.29.10-HORIZONTAL (10 registros, 0.4% del total). La presencia de un flujo con prefijo CE-QA activo en la base de producción de leads requiere aclaración técnica. No compromete la validez de las decisiones de crédito de esos 10 registros, pero la separación de ambientes QA / Producción debe estar documentada y controlada. Ruta de resolución: confirmar con Jefatura o Arquitectura de Datos si el flujo QA quedó activo por diseño o por error de configuración; en caso de error, segregar o etiquetar los registros afectados en el expediente de auditoría.', imp:'Medio' },
  { id:'P-CEL11-003', blq:'CEL-11', desc:'105 registros en CEL-11 presentan valor numérico puro en el campo Tasa (ej. "16.7", "17.7") en lugar del texto completo de TRI esperado (formato "TRI6M_BCCR_..."). El campo Tasa es el registro auditable de la tasa aprobada por el motor. El formato inconsistente dificulta la verificación automatizada del spread por segmento (B1: 12.18pp, 13.13pp, 15.80pp, 17.50pp sobre TRI). Adicionalmente, 95 registros muestran "Pendiente de subsanar" como valor de tasa, indicando que el campo no se resolvió al momento de la evaluación. Ruta de resolución: confirmar con Arquitectura de Datos el campo fuente de la tasa en el expediente y estandarizar el formato de captura en CEL-11 para todos los registros futuros.', imp:'Bajo' },
  { id:'P-CEL12-001', blq:'CEL-12', desc:'113 registros en CEL-12 presentan Causa principal = Verdadero con campo Causa vacío: la razón principal de rechazo no puede identificarse porque CEL-02 no tiene entrada para el código que disparó el motor (Incidencia de catálogo). En total, 1,685 filas (29.4% de CEL-12) quedan con Causa vacía por esta razón — desglose de causas no catalogadas: SIN_OPERACIONES_CON_AHORRO_POSITIVO: 1,349 filas · CPH_FALTANTE: 141 · ANTIGUEDAD_CREDITO_FALTANTE: 128 · ORIGEN_INGRESOS_NO_ADMISIBLE: 52 · MONTO_SUPERA_MAXIMO_PERFIL: 4 · EDAD_MAYOR_65_NO_PENSIONADO: 1. Ruta de resolución: (1) Agregar las 6 causas faltantes al catálogo CEL-02 con su respaldo normativo correspondiente (ver H-CEL12-01). (2) Reprocesar los 113 registros con causa principal vacía una vez actualizado el catálogo. (3) Confirmar con Jefatura si SIN_OPERACIONES_CON_AHORRO_POSITIVO (79.9% de la incidencia) corresponde a un criterio vigente del producto o a una regla experimental.', imp:'Alto' },
];

// HALLAZGOS FORMALES
const hallazgos = [
  { id:'H-001', cod:'GOB-003', desc:'Guía conjunta de uso del mecanismo (Banca Digital + Soporte al Negocio) no ha sido elaborada ni implementada.', resp:'Banca Digital / Soporte al Negocio', accion:'Elaborar guía conjunta conforme al ALCO.' },
  { id:'H-002', cod:'GOB-005', desc:'Soporte al Negocio no ha definido el mecanismo para identificar la cartera de no clientes BP elegibles.', resp:'Soporte al Negocio', accion:'Definir y documentar el mecanismo de identificación de cartera.' },
  { id:'H-003', cod:'GOB-008', desc:'DOCCORP-0719-2026 no ha sido procedimentado ni implementado en Banca Digital para actualización de información de clientes.', resp:'Banca Digital', accion:'Implementar las instrucciones del DOCCORP-0719-2026. Riesgo derivado en aprobaciones condicionadas por PCSC.' },
];

// BLOQUE CEL-02
const CEL02_COLS = [2400, 700, 1100, 1260, 800, 3100]; // sum = 9360
const cel02Rules = [
  { c:'DATOS_PERFIL_FALTANTES',            tipo:'Técnica', fase:'Datos y perfil',   rel:'Integridad de entradas',    e:'CONFORME', o:'Reprocesable. Activa cuando faltan campos críticos para evaluar cualquier regla CEL-01.' },
  { c:'DATOS_PERFIL_INVALIDOS',            tipo:'Técnica', fase:'Datos y perfil',   rel:'Integridad de entradas',    e:'CONFORME', o:'Reprocesable. Datos fuera de rango, formato incorrecto o consistencia inválida.' },
  { c:'EDAD_MENOR_19',                     tipo:'Rechazo', fase:'Perfil',           rel:'Edad mínima (B1)',          e:'CONFORME', o:'No reprocesable. Alineado a REGLA-001…013: todos los segmentos exigen ≥19 años.' },
  { c:'PENSIONADO_MAYOR_70',               tipo:'Rechazo', fase:'Perfil',           rel:'Edad pensionado (B1)',       e:'CONFORME', o:'No reprocesable. Alineado a REGLA-010/011: pensionados hasta 70 años.' },
  { c:'POLIZA_MAYOR_65_NO_ACEPTADA',       tipo:'Rechazo', fase:'Perfil',           rel:'Póliza mayores de 65 (B1)', e:'CONFORME', o:'Reprocesable. Póliza obligatoria REGLA-010/011. Se puede resolver si el cliente acepta la póliza.' },
  { c:'CATEGORIA_INTERNA_NO_ADMISIBLE',    tipo:'Rechazo', fase:'Riesgo interno',   rel:'Categoría interna',         e:'HALLAZGO', o:'H-CEL02-02: Descripcion dice "no es 1" (solo cat.1 admisible), DescripcionReporte dice "no es 1 ni 2" (cat.1 y 2 admisibles). Criterios contradictorios. Determinar cuál aplica en el motor.' },
  { c:'EMBARGO_PRESENTE',                  tipo:'Rechazo', fase:'CREDID',           rel:'Embargos (ALCO-GEN-006)',   e:'CONFORME', o:'No reprocesable. Fuente: CREDID API. Alineado a ALCO-GEN-006.' },
  { c:'JUICIO_ACTIVO_PRESENTE',            tipo:'Rechazo', fase:'CREDID',           rel:'Juicios (ALCO-GEN-006)',    e:'CONFORME', o:'No reprocesable. Evalúa contra catálogo jurídico CEL-03. Alineado a ALCO-GEN-006.' },
  { c:'DATOS_JUICIOS_INCOMPLETOS',         tipo:'Técnica', fase:'CREDID',           rel:'Juicios (ALCO-GEN-006)',    e:'CONFORME', o:'Reprocesable. Gestiona juicios en estado no catalogado en CEL-03 o datos ambiguos.' },
  { c:'RELACION_CREDITICIA_NO_ADMISIBLE',  tipo:'Rechazo', fase:'Segmentación',     rel:'Relación crediticia (B1)',  e:'CONFORME', o:'No reprocesable. Cubre los tipos de ingreso admisibles por segmento: asalariado, ingresos propios, pensionado.' },
  { c:'SCORE_INTERNO_NO_APROBADO',         tipo:'Rechazo', fase:'Segmentación',     rel:'Score zona blanca (ALCO-GEN-014)', e:'CONFORME', o:'No reprocesable. El valor "1,0000" es la representación interna normalizada de score aprobado (solo no clientes BP).' },
  { c:'BASE_VENCIDA_60',                   tipo:'Rechazo', fase:'Vigencia',         rel:'Vigencia base (ALCO-GEN-008)', e:'CONFORME', o:'Reprocesable. Confirma exactamente los 60 días naturales del ALCO. Nuevo proceso reinicia conteo.' },
  { c:'CUOTA_PROYECTADA_SUPERA_CIC',       tipo:'Rechazo', fase:'Capacidad',        rel:'Cuota máxima CIC (ALCO-GEN-003)', e:'CONFORME', o:'Reprocesable. Cuota remanente + cuota nueva vs. tope CIC aprobado. Alineado a ALCO-GEN-003.' },
  { c:'INGRESO_NO_SUPERA_CUOTAS_CIC',      tipo:'Rechazo', fase:'Capacidad',        rel:'Ingreso vs cuotas (ALCO-GEN-004)', e:'CONFORME', o:'Reprocesable. Ingreso aplicable debe superar total cuotas CIC. Alineado a ALCO-GEN-004.' },
  { c:'INTANGIBLE_ULTIMO_MES_NO_CUMPLE',   tipo:'Rechazo', fase:'Capacidad',        rel:'Intangible dual — SIN REF.', e:'HALLAZGO', o:'H-CEL02-01: La regla "intangible versionado" (último mes) no aparece en CEL-01 ni en el texto del ALCO. El sistema puede rechazar por criterio sin respaldo normativo visible.' },
  { c:'INTANGIBLE_PROMEDIO_12M_NO_CUMPLE', tipo:'Rechazo', fase:'Capacidad',        rel:'Intangible dual — SIN REF.', e:'HALLAZGO', o:'H-CEL02-01: Ídem anterior; aplica promedio de los últimos 12 meses. Mismo hallazgo, segunda causa.' },
  { c:'PROMEDIO_12M_INCOMPLETO',           tipo:'Técnica', fase:'Capacidad',        rel:'Salario promedio 12M (B1)', e:'CONFORME', o:'Reprocesable. Maneja historial de salario incompleto en el cálculo del promedio mensual.' },
  { c:'SIN_SEGMENTO_CERO_ESTRES',          tipo:'Rechazo', fase:'Segmentación',     rel:'Segmentos A-D (B1)',        e:'CONFORME', o:'No reprocesable. Fallback: ninguna combinación de reglas A/B/C/D aplica al perfil.' },
  { c:'TRI6M_NO_DISPONIBLE',               tipo:'Técnica', fase:'Precio',           rel:'TRI 6M BCCR (B1)',         e:'CONFORME', o:'Reprocesable. Las tasas fijas del ALCO son spreads sobre la TRI 6M BCCR vigente. Sin la tasa base no se puede fijar precio.' },
  { c:'AUTORIZACION_SUGEF_NO_VIGENTE',     tipo:'Rechazo', fase:'Autorización',     rel:'Autorización CIC (ALCO-GEN-002)', e:'CONFORME', o:'Reprocesable. Refuerza ALCO-GEN-002 y el flujo CEL-27. La autorización del cliente para consulta CIC debe estar vigente.' },
  { c:'LISTAS_O_PCSC_NO_CONFORME',         tipo:'Rechazo', fase:'Cumplimiento',     rel:'PCSC y listas (ALCO-GOB-009/010/011)', e:'CONFORME', o:'Reprocesable. Confirma que el control PCSC/listas nacionales e internacionales existe en el motor de reglas (evidencia parcial P-003). Refuerza ALCO-GOB-009/010/011.' },
  { c:'ARCHIVO_ACTA_FALLIDO',              tipo:'Técnica', fase:'Emisión',          rel:'Acta individual (ALCO-ACT-005)', e:'CONFORME', o:'Reprocesable. Gestiona fallas en la creación y verificación del acta individual en el expediente. Alineado a ALCO-ACT-005.' },
  { c:'CORREO_CONTROL_FALLIDO',            tipo:'Técnica', fase:'Notificación',     rel:'Correo controlado',        e:'CONFORME', o:'Reprocesable. Aviso interno al correo de control. No bloquea el proceso con el cliente.' },
  { c:'CPH_NO_ADMISIBLE',                  tipo:'Rechazo', fase:'Segmentación',     rel:'CPH por segmento (B1)',     e:'CONFORME', o:'No reprocesable. Seg.A, B, D: CPH=1; Seg.C: CPH<1,25. Alineado a REGLA-004…013.' },
  { c:'ATRASO_MAXIMO_NO_ADMISIBLE',        tipo:'Rechazo', fase:'Segmentación',     rel:'Atraso por segmento (B1)', e:'CONFORME', o:'No reprocesable. Seg.A/C/D: 0 días; Seg.B: admite hasta 30d. Nota: el campo de la causa es más restrictivo que el ALCO para Seg.D (30d).' },
  { c:'ANTIGUEDAD_CREDITO_NO_ADMISIBLE',   tipo:'Rechazo', fase:'Segmentación',     rel:'Antigüedad por segmento (B1)', e:'CONFORME', o:'No reprocesable. Seg.A: ≥48m; Seg.B/C/D: 12-48m. Alineado a REGLA-001…013.' },
  { c:'SCORE_INTERNO_15V_NO_APROBADO',     tipo:'Rechazo', fase:'Segmentación',     rel:'Score auditable 15 vars',  e:'CONFORME', o:'Versión CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181 (vigente 25/08/2026). Documenta corte 181 pts y excluye zona gris 161-180 no autorizada por ALCO. Evidencia parcial de P-001.' },
];

const hallazgosCEL02 = [
  { id:'H-CEL02-01', cod:'CEL-02: INTANGIBLE_ULTIMO_MES_NO_CUMPLE · INTANGIBLE_PROMEDIO_12M_NO_CUMPLE', desc:'El catálogo aplica una lógica de "intangible versionado" en la fase de Capacidad (último mes y promedio 12 meses) que no tiene referencia en ninguna regla de CEL-01 ni en el texto del Acuerdo ALCO. El motor puede rechazar solicitudes por una condición sin respaldo normativo visible. Se desconoce si proviene de lineamientos internos de crédito BP o de normativa SUGEF.', resp:'Jefatura Dirección Banca Digital', accion:'1) Consultar a Jefatura si el intangible tiene respaldo en lineamientos internos de crédito BP (SUGEF u otro). 2) Si existe respaldo: documentarlo como regla en CEL-01 y referenciar la normativa. 3) Si no existe: evaluar si la causa debe retirarse o ajustarse al texto del ALCO.' },
  { id:'H-CEL02-02', cod:'CEL-02: CATEGORIA_INTERNA_NO_ADMISIBLE', desc:'La causa presenta una inconsistencia interna entre dos de sus campos: Descripcion dice "La categoría interna de riesgo no es 1" (criterio: solo cat.1 admisible), mientras que DescripcionReporte — el texto que se expone en reportes y expedientes — dice "La categoría interna no es 1 ni 2" (criterio: cat.1 y cat.2 admisibles). Los criterios son distintos; uno de ellos refleja incorrectamente la lógica real del motor.', resp:'Jefatura Dirección Banca Digital', accion:'1) Verificar en el código del motor cuál es el criterio efectivamente implementado. 2) Corregir el campo incorrecto (Descripcion o DescripcionReporte) para mantener consistencia. 3) Si la regla admite cat.2, actualizar también la regla correspondiente en CEL-01.' },
];

// HALLAZGOS CEL-12
const hallazgosCEL12 = [
  { id:'H-CEL12-01', cod:'CEL-12 / CEL-02: 6 causas del motor sin entrada en catálogo', desc:'El motor CEL-01 genera 6 códigos de causa que no tienen entrada en el catálogo CEL-02 vigente. Cuando cualquiera de estos códigos se activa, el campo Causa de CEL-12 queda vacío (Incidencia de catálogo). Detalle de las 1,685 filas afectadas (29.4% de CEL-12): SIN_OPERACIONES_CON_AHORRO_POSITIVO: 1,349 filas (79.9% de la incidencia) · CPH_FALTANTE: 141 · ANTIGUEDAD_CREDITO_FALTANTE: 128 · ORIGEN_INGRESOS_NO_ADMISIBLE: 52 · MONTO_SUPERA_MAXIMO_PERFIL: 4 · EDAD_MAYOR_65_NO_PENSIONADO: 1. De estas filas, 113 corresponden a registros donde Causa principal = Verdadero, es decir, la razón principal de rechazo queda sin identificar. Este hallazgo evidencia que el motor CEL-01 y el catálogo CEL-02 están desincronizados: el motor incorpora criterios de rechazo no reflejados en el catálogo de causas vigente.', resp:'Jefatura Dirección Banca Digital', accion:'1) Confirmar con Jefatura si las 6 causas corresponden a criterios vigentes del producto. 2) Agregar las entradas faltantes al catálogo CEL-02 con su respaldo normativo (ver CEL-01 para la regla origen de cada código). 3) Verificar si las causas están contempladas en alguna versión CE-REGLAS-ALCO-2026-V2.x (relacionado con P-CEL10-002). 4) Reprocesar los 113 registros con causa principal vacía una vez actualizado el catálogo.' },
];


// BLOQUE CEL-04
const CEL04_COLS = [700, 1800, 780, 760, 780, 900, 3640]; // sum = 9360
const cel04Rules = [
  {
    c:'CE-RES-01', nombre:'Aprobado',
    final:'Sí', califica:'Sí', analista:'No', reproc:'No',
    conds:'—',
    e:'CONFORME',
    o:'Decisión final positiva. El motor aprueba automáticamente sin condiciones pendientes. Alineado a ALCO-ACT-006 (aprobación sin condiciones).'
  },
  {
    c:'CE-RES-02', nombre:'Aprobado Condicionado',
    final:'No', califica:'Sí', analista:'Sí', reproc:'No',
    conds:'1) PCSC vencida · 2) Cédula de identidad vencida · 3) Operaciones en atraso superior al permitido · 4) Estado PCSC diferente de Activo / Activo CES / Expediente Simplificado · 5) "No cliente" admisible cuando la persona no posee PCSC',
    e:'PENDIENTE',
    o:'Alineado a ALCO-ACT-007 (aprobación condicionada). Enriquece P-003: define estados PCSC admisibles (Activo, Activo CES, Expediente Simplificado). Condición 3 (atraso): se aplica a las operaciones negociadas específicas que se cancelan con el giro, dentro del período de 60 días de vigencia del acta — no al historial SUGEF del cliente (ya evaluado en segmentación B1). La consolidación resuelve la condición por diseño del producto (ALCO-GEN-013). → P-CEL04-001: pendiente confirmación formal. Versión CE-RESULTADOS-2026.08.29.1 es la primera catalogación fechada 12 días post-ALCO; sugiere que las condiciones de aprobación condicionada se formalizaron después de la firma del acuerdo.'
  },
  {
    c:'CE-RES-03', nombre:'Denegado',
    final:'Sí', califica:'No', analista:'No', reproc:'No',
    conds:'—',
    e:'CONFORME',
    o:'Decisión final negativa. El motor rechaza automáticamente. No reprocesable bajo la versión de reglas vigente. Alineado a la lógica de rechazo definitivo de CEL-01 + CEL-02.'
  },
  {
    c:'CE-RES-04', nombre:'Falta Información',
    final:'No', califica:'Sin decisión', analista:'No', reproc:'Sí',
    conds:'—',
    e:'CONFORME',
    o:'Incidencia técnica. Tiene prioridad de reproceso sobre los demás resultados. Corresponde a casos donde datos faltantes impidieron evaluar las reglas. Alineado a la lógica de causas técnicas en CEL-02 (DATOS_PERFIL_FALTANTES, DATOS_PERFIL_INVALIDOS, etc.).'
  },
];

function makeCEL04Table() {
  const cols = CEL04_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]),
        hCell('Resultado', cols[1]),
        hCell('Dec. Final', cols[2]),
        hCell('Califica', cols[3]),
        hCell('Analista', cols[4]),
        hCell('Reproc.', cols[5]),
        hCell('Condiciones / Observación', cols[6]),
      ]}),
      ...cel04Rules.map((r, i) => {
        const bg = r.e === 'PENDIENTE' ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        const obsText = (r.conds !== '—' ? 'Condiciones: ' + r.conds + '\n\n' : '') + r.o;
        return new TableRow({ children: [
          cell(r.c,       { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.nombre,  { w: cols[1], bold: true, bg, size: SZ_SM }),
          cell(r.final,   { w: cols[2], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.califica,{ w: cols[3], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.analista,{ w: cols[4], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.reproc,  { w: cols[5], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(obsText,   { w: cols[6], bg, size: SZ_SM, italic: r.e !== 'PENDIENTE', bold: r.e === 'PENDIENTE', textColor: r.e === 'PENDIENTE' ? AMBER_TEXT : TEXT_DARK }),
        ]});
      }),
    ],
  });
}

function makeCEL04ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Resultados revisados', cols[1], PURPLE),
        hCell('Conformes', cols[2], PURPLE), hCell('Pendientes', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-04 — Catálogo de Resultados de Decisión (versión CE-RESULTADOS-2026.08.29.1)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('4',  { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('3',  { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('1',  { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}


// BLOQUE CEL-05
const CEL05_COLS = [1700, 1800, 1800, 660, 700, 1100, 1600]; // sum = 9360
const cel05Rules = [
  {
    titulo: 'CE-SCORE-CONSUMO-V1.0-DRAFT',
    estado: 'VALIDACION_TECNICA_MODO_SOMBRA',
    estadoCorte: 'APROBADO_JDN_5824_ACD_380_2021_ART_9',
    corte: '181',
    decHab: 'Verdadero',
    rango: '100 – 315 pts · 15 vars',
    alco: 'JDN-5824-Acd-380-2021-Art-9 · Consumo 181-200 · Zona gris 161-180',
    e: 'PENDIENTE',
    o: 'Corte 181 con respaldo normativo formal (JDN-5824). Puntaje crudo conservado en PIB para auditoría; no expuesto en modelo semántico por diseño. Pendientes de mantenimiento: (1) actualizar versión a V1.1-SOLO-NO-CLIENTE-CORTE181; (2) actualizar estado de MODO_SOMBRA a producción. → P-001 (verificar vigencia JDN-5824 para Cero Estrés 2026) · P-CEL05-001 (mantenimiento catálogo).',
  },
];

function makeCEL05Table() {
  const cols = CEL05_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Título / Versión', cols[0]),
        hCell('Estado operativo', cols[1]),
        hCell('Estado del corte', cols[2]),
        hCell('Corte', cols[3]),
        hCell('Dec. Hab.', cols[4]),
        hCell('Rango / Vars', cols[5]),
        hCell('Instrumento normativo', cols[6]),
      ]}),
      ...cel05Rules.map((r, i) => {
        const bg = r.e === 'PENDIENTE' ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.titulo,     { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.estado,     { w: cols[1], bg, size: SZ_XS }),
          cell(r.estadoCorte,{ w: cols[2], bg, size: SZ_XS }),
          cell(r.corte,      { w: cols[3], bg, size: SZ_SM, align: AlignmentType.CENTER, bold: true }),
          cell(r.decHab,     { w: cols[4], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.rango,      { w: cols[5], bg, size: SZ_XS }),
          cell(r.alco,       { w: cols[6], bg, size: SZ_XS, italic: true }),
        ]});
      }),
    ],
  });
}

function makeCEL05ObsTable() {
  const cols = [9360];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [ hCell('Observaciones y hallazgos de auditoría — CEL-05', cols[0], PURPLE) ]}),
      ...cel05Rules.map((r, i) => {
        const bg = r.e === 'PENDIENTE' ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.titulo + ': ' + r.o, { w: cols[0], bg, size: SZ_SM,
            italic: r.e !== 'PENDIENTE', bold: false,
            textColor: r.e === 'PENDIENTE' ? AMBER_TEXT : TEXT_DARK }),
        ]});
      }),
    ],
  });
}

function makeCEL05ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Modelos revisados', cols[1], PURPLE),
        hCell('Conformes', cols[2], PURPLE), hCell('Pendientes', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-05 — Modelos de Score (versión en catálogo: CE-SCORE-CONSUMO-V1.0-DRAFT · producción: V1.1)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('1', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('0', { w: cols[2], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('1', { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────
// DATOS Y TABLAS — BLOQUE SCORE (CEL-06, CEL-07, CEL-08, CEL-21)
// ─────────────────────────────────────────────────────────────────────────

// CEL-06: cobertura por variable
const cel06Coverage = [
  { v:'01. Edad',            cod:'I_ANIOS_NCMNT_CNSTTCN',   cob:'99.89', nivel:'ALTA',   miss:'G05 (32-38a)',     pts:'12' },
  { v:'02. Distrito',        cod:'N_DISTRITO',               cob:'99.23', nivel:'ALTA',   miss:'G05',              pts:'12' },
  { v:'03. Escolaridad',     cod:'N_NIVEL_ESCOLARIDAD',      cob:'49.65', nivel:'MEDIA',  miss:'G05 (máx.)',       pts:'19' },
  { v:'04. Oficina',         cod:'N_UNI_FIN_CM',             cob:'27.76', nivel:'BAJA',   miss:'G02',              pts:'8'  },
  { v:'05. Seg. ingreso',    cod:'N_CD_SEG_ING_CM',          cob:'98.27', nivel:'ALTA',   miss:'P+MISSING',        pts:'16' },
  { v:'06. Créd. consumo',   cod:'B_PCT_CONSUMO_L13TO24',    cob:'77.78', nivel:'MEDIA',  miss:'0+MISSING',        pts:'10' },
  { v:'07. Cuotas BP',       cod:'I_CUOTAS_BP',              cob:'27.91', nivel:'BAJA',   miss:'G01 (mín.)',       pts:'11' },
  { v:'08. Estado civil',    cod:'B_EDO_CIVIL',              cob:'96.62', nivel:'ALTA',   miss:'0+MISSING',        pts:'10' },
  { v:'09. Género',          cod:'N_CD_GEN_CM',              cob:'97.90', nivel:'ALTA',   miss:'1+MISSING',        pts:'10' },
  { v:'10. Ocupación lab.',  cod:'N_OCU_LAB_G_CM',           cob:'64.03', nivel:'MEDIA',  miss:'MISSING+UNKNOWN',  pts:'21' },
  { v:'11. Origen fondos',   cod:'N_ORI_FND_G_CM',           cob:'17.34', nivel:'BAJA ★', miss:'G08 (med.)',       pts:'17' },
  { v:'12. Pago planilla',   cod:'B_PAGOPLANILLA',           cob:'77.28', nivel:'MEDIA',  miss:'1+MISSING',        pts:'17' },
  { v:'13. PIB BCCR',        cod:'I_IM_MAX_PIB_YLD_L7_9M',  cob:'99.12', nivel:'ALTA',   miss:'MISSING (sup.)',   pts:'16' },
  { v:'14. Sector patronal', cod:'N_CD_TIP_SEC_CM',          cob:'90.32', nivel:'ALTA',   miss:'V,X+MISSING',      pts:'11' },
  { v:'15. Afiliado PAS',    cod:'B_AFILIADO_PAS',           cob:'77.28', nivel:'MEDIA',  miss:'0+MISSING',        pts:'10' },
];

// CEL-07: resumen por variable (grupos, rango, nota)
const cel07Summary = [
  { v:'01. Edad',            grp:'10', rango:'1–26',   miss:'G05 = 12',  nota:'' },
  { v:'02. Distrito',        grp:'10', rango:'-4–41',  miss:'G05 = 12',  nota:'Único negativo en distrito' },
  { v:'03. Escolaridad',     grp:'5',  rango:'11–19',  miss:'G05 = 19',  nota:'MISSING = grupo máximo' },
  { v:'04. Oficina',         grp:'10', rango:'5–29',   miss:'G02 = 8',   nota:'' },
  { v:'05. Seg. ingreso',    grp:'3',  rango:'7–26',   miss:'P+MISS = 16',nota:'' },
  { v:'06. Créd. consumo',   grp:'2',  rango:'10–16',  miss:'0+MISS = 10',nota:'' },
  { v:'07. Cuotas BP',       grp:'7',  rango:'11–17',  miss:'G01 = 11',  nota:'MISSING = grupo mínimo' },
  { v:'08. Estado civil',    grp:'2',  rango:'10–16',  miss:'0+MISS = 10',nota:'' },
  { v:'09. Género',          grp:'2',  rango:'10–16',  miss:'1+MISS = 10',nota:'' },
  { v:'10. Ocupación lab.',  grp:'6',  rango:'6–21',   miss:'MISS = 21', nota:'MISSING = grupo máximo (CEL-21)' },
  { v:'11. Origen fondos',   grp:'9',  rango:'8–18',   miss:'G08 = 17',  nota:'82.66% casos = MISSING' },
  { v:'12. Pago planilla',   grp:'2',  rango:'-4–17',  miss:'1+MISS = 17',nota:'★ Único grupo negativo del modelo' },
  { v:'13. PIB BCCR',        grp:'5',  rango:'11–12',  miss:'MISS = 16', nota:'MISSING supera todos los rangos' },
  { v:'14. Sector patronal', grp:'2',  rango:'11–16',  miss:'V,X+MISS = 11',nota:'' },
  { v:'15. Afiliado PAS',    grp:'2',  rango:'10–21',  miss:'0+MISS = 10',nota:'' },
];

// CEL-21: categorías normalizadas de ocupación
const cel21Cats = [
  { g:'1', cat:'ELEMENTALES_INSTALACION / ELEMENTALES_N_D', pts:'6',  ej:'Peón, Taxista, Empacador, Ama de casa, Estudiante, Misceláneo' },
  { g:'2', cat:'OFICIOS',                                   pts:'9',  ej:'Albañil, Carpintero, Mecánico automotriz, Ajustador de motores' },
  { g:'3', cat:'SERV_VENDEDORES',                           pts:'10', ej:'Cajero, Salonero, Agente de ventas, Cocinero, Empleado doméstico' },
  { g:'4', cat:'ADMINISTRATIVO',                            pts:'13', ej:'Oficinista, Secretaria, Telefonista, Digitador, Recepcionista' },
  { g:'5', cat:'AGRI_TECNICOS',                             pts:'17', ej:'Técnico TIC, Técnico de salud, Auxiliar enfermería, Mecánico instrumentos' },
  { g:'6', cat:'CIENTIFICOS_GERENCIALES',                   pts:'21', ej:'Abogado, Médico, Gerente, Programador, Decano, Contador, Arquitecto' },
];

function makeCEL06ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Variables revisadas', cols[1], PURPLE),
        hCell('Cobertura ≥90%', cols[2], PURPLE), hCell('Pendientes', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-06 — Variables del Score · muestra 02/09/2026, n=406 · métricas 17/09/2026 · base 2,723 casos', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('15', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('6',  { w: cols[2], bg: GREEN_BG,   textColor: GREEN_TEXT,  size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('1',  { w: cols[3], bg: AMBER_BG,   textColor: AMBER_TEXT,  size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

function makeCEL06Table() {
  const cols = [1900, 2600, 840, 840, 1680, 1500];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Variable', cols[0], PURPLE),     hCell('Campo técnico', cols[1], PURPLE),
        hCell('Cobertura', cols[2], PURPLE),    hCell('Nivel', cols[3], PURPLE),
        hCell('Grupo MISSING asignado', cols[4], PURPLE), hCell('Pts MISSING', cols[5], PURPLE),
      ]}),
      ...cel06Coverage.map((r, i) => {
        const isBaja = r.nivel.startsWith('BAJA');
        const isAlta = r.nivel === 'ALTA';
        const bg     = isBaja ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        const cobBg  = isBaja ? AMBER_BG : (isAlta ? GREEN_BG : bg);
        const cobCol = isBaja ? AMBER_TEXT : (isAlta ? GREEN_TEXT : TEXT_DARK);
        return new TableRow({ children: [
          cell(r.v,          { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.cod,        { w: cols[1], bg, size: SZ_XS, italic: true }),
          cell(r.cob + '%',  { w: cols[2], bg: cobBg, textColor: cobCol, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.nivel,      { w: cols[3], bg: cobBg, textColor: cobCol, size: SZ_XS, align: AlignmentType.CENTER }),
          cell(r.miss,       { w: cols[4], bg, size: SZ_XS }),
          cell(r.pts,        { w: cols[5], bg, size: SZ_SM, align: AlignmentType.CENTER }),
        ]});
      }),
    ],
  });
}

function makeCEL07ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Reglas WOE', cols[1], PURPLE),
        hCell('Variables', cols[2], PURPLE), hCell('Instrumento normativo', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-07 — Reglas del Score · todas gobernadas por JDN-5824-Acd-380-2021-Art-9 · rango 100–315 pts', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('78', { w: cols[1], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('15', { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('JDN-5824', { w: cols[3], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

function makeCEL07Table() {
  const cols = [1900, 700, 1200, 1800, 3760];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Variable', cols[0], PURPLE),   hCell('Grupos', cols[1], PURPLE),
        hCell('Rango pts', cols[2], PURPLE),  hCell('MISSING → pts', cols[3], PURPLE),
        hCell('Nota de diseño', cols[4], PURPLE),
      ]}),
      ...cel07Summary.map((r, i) => {
        const isFlag = r.nota.startsWith('★') || r.nota.includes('MISSING supera') || r.nota.includes('82.66%');
        const bg = isFlag ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.v,     { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.grp,   { w: cols[1], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.rango, { w: cols[2], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.miss,  { w: cols[3], bg, size: SZ_SM }),
          cell(r.nota,  { w: cols[4], bg: r.nota ? AMBER_BG : bg, textColor: r.nota ? AMBER_TEXT : TEXT_DARK, size: SZ_XS, italic: !!r.nota }),
        ]});
      }),
    ],
  });
}

function makeCEL21ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Entradas', cols[1], PURPLE),
        hCell('Vigentes', cols[2], PURPLE), hCell('Pendientes', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-21 — Equivalencias Ocupación CREDID · versión OCUP-CREDID-V1.0-20260907 · vigente 07/09/2026', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('160', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('160', { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('1',   { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

function makeCEL21Table() {
  const cols = [600, 3000, 900, 4860];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Grp.', cols[0], PURPLE), hCell('Categoría normalizada', cols[1], PURPLE),
        hCell('Pts', cols[2], PURPLE),  hCell('Ejemplos de ocupaciones (muestra)', cols[3], PURPLE),
      ]}),
      ...cel21Cats.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        return new TableRow({ children: [
          cell(r.g,   { w: cols[0], bold: true, bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.cat, { w: cols[1], bold: true, bg, size: SZ_SM }),
          cell(r.pts, { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.ej,  { w: cols[3], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

// ─── CEL-10 DATA & TABLES ────────────────────────────────────────────────
const cel10VersionData = [
  { ver:'CEL-01-PRODUCCION-V1.0',    n:2108, apro:456,  cond:18,  fi:193, den:1441, otros:0,   pct:'71.0%' },
  { ver:'CE-REGLAS-ALCO-2026-V2.3',  n:429,  apro:119,  cond:4,   fi:28,  den:278,  otros:0,   pct:'14.4%' },
  { ver:'Sin versión (pre-ALCO)',     n:423,  apro:0,    cond:132, fi:0,   den:0,    otros:291, pct:'14.2%' },
  { ver:'CE-REGLAS-ALCO-2026-V2.0',  n:8,    apro:0,    cond:4,   fi:4,   den:0,    otros:0,   pct:'0.3%'  },
  { ver:'CE-REGLAS-ALCO-2026-V2.1',  n:2,    apro:0,    cond:0,   fi:2,   den:0,    otros:0,   pct:'0.1%'  },
];

const cel10TopCausas = [
  { causa:'No se identificaron operaciones con ahorro y monto positivo para la oferta', n:1116, pct:'37.6%', relaciona:'CEL-01 Seg. B/C/D' },
  { causa:'Sin causa adversa', n:575, pct:'19.4%', relaciona:'Aprobados' },
  { causa:'El salario promedio de 12 meses no supera el intangible después de cuotas', n:518, pct:'17.4%', relaciona:'REGLA-001/002' },
  { causa:'Score interno auditable inferior a 181 puntos para no cliente de crédito', n:479, pct:'16.1%', relaciona:'ALCO-GEN-014 / JDN-5824' },
  { causa:'El salario del último mes no supera el intangible después de cuotas', n:471, pct:'15.9%', relaciona:'REGLA-001/002' },
  { causa:'Relación crediticia no admisible para el perfil', n:459, pct:'15.5%', relaciona:'ALCO-GEN-001/CEL-01' },
  { causa:'No cumple las condiciones de los segmentos Cero Estrés', n:451, pct:'15.2%', relaciona:'CEL-01 B1-B4' },
  { causa:'Falta el estado de la PCSC', n:394, pct:'13.3%', relaciona:'H-003 · ALCO-GOB-008' },
  { causa:'Datos obligatorios del perfil incompletos', n:180, pct:'6.1%', relaciona:'CEL-04 CE-RES-04' },
  { causa:'Antigüedad del crédito fuera del límite permitido', n:159, pct:'5.4%', relaciona:'REGLA-001' },
];

function makeCEL10VersionTable() {
  const cols = [2900, 600, 700, 600, 700, 700, 700, 960];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Versión de matriz', cols[0], PURPLE),
        hCell('Total', cols[1], PURPLE),
        hCell('%', cols[2], PURPLE),
        hCell('Aprobado', cols[3], PURPLE),
        hCell('Aprobado Cond.', cols[4], PURPLE),
        hCell('Falta Info.', cols[5], PURPLE),
        hCell('Denegado', cols[6], PURPLE),
        hCell('Otros*', cols[7], PURPLE),
      ]}),
      ...cel10VersionData.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        const isAudit = r.ver === 'CEL-01-PRODUCCION-V1.0';
        return new TableRow({ children: [
          cell(r.ver,          { w: cols[0], bold: isAudit, bg: isAudit ? PURPLE_LIGHT : bg, size: SZ_SM }),
          cell(String(r.n),    { w: cols[1], bold: isAudit, bg: isAudit ? PURPLE_LIGHT : bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.pct,          { w: cols[2], bg: isAudit ? PURPLE_LIGHT : bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.apro > 0  ? String(r.apro)  : '—', { w: cols[3], bg: r.apro  > 0 ? GREEN_BG : (isAudit ? PURPLE_LIGHT : bg), textColor: r.apro  > 0 ? GREEN_TEXT : undefined, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.cond > 0  ? String(r.cond)  : '—', { w: cols[4], bg: r.cond  > 0 ? AMBER_BG : (isAudit ? PURPLE_LIGHT : bg), size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.fi   > 0  ? String(r.fi)    : '—', { w: cols[5], bg: isAudit ? PURPLE_LIGHT : bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.den  > 0  ? String(r.den)   : '—', { w: cols[6], bg: r.den   > 0 ? RED_BG   : (isAudit ? PURPLE_LIGHT : bg), textColor: r.den   > 0 ? RED_TEXT   : undefined, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.otros > 0 ? String(r.otros) : '—', { w: cols[7], bg: isAudit ? PURPLE_LIGHT : bg, size: SZ_SM, align: AlignmentType.CENTER }),
        ]});
      }),
      new TableRow({ children: [
        cell('TOTAL', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('2,970', { w: cols[1], bold: true, bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('100%',  { w: cols[2], bold: true, bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('575',   { w: cols[3], bold: true, bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('158',   { w: cols[4], bold: true, bg: AMBER_BG, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('227',   { w: cols[5], bold: true, bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('1,719', { w: cols[6], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('291',   { w: cols[7], bold: true, bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
      ]},),
    ],
  });
}

function makeCEL10CausasTable() {
  const cols = [3760, 600, 700, 4300];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Causa de la decisión', cols[0], PURPLE),
        hCell('N', cols[1], PURPLE),
        hCell('%', cols[2], PURPLE),
        hCell('Norma / catálogo relacionado', cols[3], PURPLE),
      ]}),
      ...cel10TopCausas.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        const isPCSC = r.causa.includes('PCSC');
        return new TableRow({ children: [
          cell(r.causa,   { w: cols[0], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM }),
          cell(r.n.toLocaleString('es-CR'), { w: cols[1], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.pct,     { w: cols[2], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.relaciona,{ w: cols[3], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM }),
        ]});
      }),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────
// CONSTRUCTORES DE TABLAS
// ─────────────────────────────────────────────────────────────────────────

function makeB1Table() {
  const cols = B1_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]), hCell('Tipo cliente', cols[1]), hCell('Seg.', cols[2]),
        hCell('Descripción', cols[3]), hCell('Estado', cols[4]), hCell('Observaciones / Pendientes', cols[5]),
      ]}),
      ...b1Rules.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        return new TableRow({ children: [
          cell(r.c, { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.t, { w: cols[1], bg, size: SZ_SM }),
          cell(r.s, { w: cols[2], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.d, { w: cols[3], bg, size: SZ_SM }),
          statusCell(r.e, cols[4]),
          cell(r.o, { w: cols[5], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

function makeB2Table() {
  const cols = B2_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]), hCell('Regla de negocio', cols[1]),
        hCell('Estado', cols[2]), hCell('Observaciones', cols[3]),
      ]}),
      ...b2Rules.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        return new TableRow({ children: [
          cell(r.c, { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.d, { w: cols[1], bg, size: SZ_SM }),
          statusCell(r.e, cols[2]),
          cell(r.o, { w: cols[3], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

function makeB3Table() {
  const cols = B3_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]), hCell('Regla de gobernanza', cols[1]),
        hCell('Estado', cols[2]), hCell('Observaciones / Hallazgo', cols[3]),
      ]}),
      ...b3Rules.map((r, i) => {
        const bg = r.e === 'HALLAZGO' ? RED_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.c, { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.d, { w: cols[1], bg, size: SZ_SM }),
          statusCell(r.e, cols[2]),
          cell(r.o, { w: cols[3], bg, size: SZ_SM, italic: r.e !== 'HALLAZGO', bold: r.e === 'HALLAZGO', textColor: r.e === 'HALLAZGO' ? RED_TEXT : TEXT_DARK }),
        ]});
      }),
    ],
  });
}

function makeB4Table() {
  const cols = B4_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código', cols[0]), hCell('Regla de negocio', cols[1]),
        hCell('Estado', cols[2]), hCell('Observaciones', cols[3]),
      ]}),
      ...b4Rules.map((r, i) => {
        const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
        return new TableRow({ children: [
          cell(r.c, { w: cols[0], bold: true, bg, size: SZ_SM }),
          cell(r.d, { w: cols[1], bg, size: SZ_SM }),
          statusCell(r.e, cols[2]),
          cell(r.o, { w: cols[3], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}


function makeHallazgosTable() {
  const cols = [600, 900, 3360, 2000, 2500];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('N°', cols[0], PURPLE), hCell('Regla', cols[1], PURPLE),
        hCell('Descripción del hallazgo', cols[2], PURPLE),
        hCell('Responsable', cols[3], PURPLE),
        hCell('Acción requerida', cols[4], PURPLE),
      ]}),
      ...hallazgos.map((h, i) => {
        const bg = i % 2 === 0 ? RED_BG : '#FCEAEA';
        return new TableRow({ children: [
          cell(h.id,     { w: cols[0], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(h.cod,    { w: cols[1], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM }),
          cell(h.desc,   { w: cols[2], bg, size: SZ_SM }),
          cell(h.resp,   { w: cols[3], bg, size: SZ_SM }),
          cell(h.accion, { w: cols[4], bg, size: SZ_SM }),
        ]});
      }),
    ],
  });
}

function makePendientesTable() {
  // cols: ID / Origen / Descripción / Impacto
  const cols = [700, 800, 5060, 800];

  // Grupos por tipo — orden de prioridad de gestión
  const grupos = [
    {
      tipo: 'NORMATIVO',
      desc: 'Instrumento o decreto de respaldo faltante o pendiente de verificar vigencia',
      ids: ['P-001', 'P-002', 'P-GIR-002'],
    },
    {
      tipo: 'GOBERNANZA',
      desc: 'Aprobación formal o confirmación en expediente pendiente',
      ids: ['P-CEL21-001', 'P-CEL04-001'],
    },
    {
      tipo: 'ARQUITECTURAL / DATOS',
      desc: 'Campo, capa o exposición de datos en el modelo CEA-02',
      ids: ['P-003', 'P-005', 'P-CEL06-001', 'P-CEL08-001'],
    },
    {
      tipo: 'MANTENIMIENTO DE CATÁLOGO',
      desc: 'Actualización administrativa de versión, estado o sincronización entre catálogos del Centro de Auditoría',
      ids: ['P-CEL05-001', 'P-CEL07-001', 'P-CEL12-001'],
    },
    {
      tipo: 'TÉCNICO / OPERATIVO',
      desc: 'Limitación técnica documentada, separación de ambientes o ítem condicionado a temporalidad',
      ids: ['P-ACT-004', 'P-CEL09-001', 'P-CEL11-002', 'P-CEL11-003'],
    },
    {
      tipo: 'CLASIFICACIÓN / DOCUMENTACIÓN',
      desc: 'Aclaración sobre registros históricos y versiones no documentadas en CEL-05',
      ids: ['P-CEL10-001', 'P-CEL10-002', 'P-CEL11-001'],
    },
  ];

  const pMap = {};
  pendientes.forEach(p => { pMap[p.id] = p; });

  const rows = [
    new TableRow({ tableHeader: true, children: [
      hCell('ID', cols[0], PURPLE), hCell('Origen', cols[1], PURPLE),
      hCell('Descripción y ruta de resolución', cols[2], PURPLE), hCell('Impacto', cols[3], PURPLE),
    ]}),
  ];

  grupos.forEach(g => {
    // Fila separadora de grupo
    rows.push(new TableRow({ children: [
      new TableCell({
        children: [new Paragraph({
          children: [
            run(g.tipo, { bold: true, color: WHITE, size: SZ_SM }),
            run('  —  ' + g.desc, { color: 'DDDDDD', size: SZ_XS, italics: true }),
          ],
          spacing: { before: 60, after: 60 },
        })],
        columnSpan: 4,
        shading: { type: ShadingType.CLEAR, fill: PURPLE_DARK, color: 'auto' },
        margins: { top: 60, bottom: 60, left: 120, right: 80 },
      }),
    ]}));

    // Filas de ítems de este grupo
    g.ids.forEach((id, i) => {
      const p = pMap[id];
      if (!p) return;
      const bg = i % 2 === 0 ? WHITE : PURPLE_LIGHT;
      const impColor = p.imp === 'Alto' ? RED_TEXT : (p.imp === 'Medio' ? AMBER_TEXT : GREEN_TEXT);
      const impBg    = p.imp === 'Alto' ? RED_BG   : (p.imp === 'Medio' ? AMBER_BG   : GREEN_BG);
      rows.push(new TableRow({ children: [
        cell(p.id,  { w: cols[0], bold: true, bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell(p.blq, { w: cols[1], bg, size: SZ_SM, align: AlignmentType.CENTER }),
        cell(p.desc,{ w: cols[2], bg, size: SZ_SM }),
        cell(p.imp, { w: cols[3], bg: impBg, textColor: impColor, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}));
    });
  });

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows,
  });
}

function makeResumenTable() {
  const cols = [3200, 1240, 1480, 1720, 1720];
  // Filas: sección / unidades revisadas / conformes / hallazgos / pendientes
  const rows = [
    // PARTE I — Acuerdo ALCO
    { n:'PARTE I · Revisión Normativa — Acuerdo ALCO', sep: true },
    { n:'B1 — Reglas de Crédito',           u:'13 reglas',         ok:13, h:0, p:2 },
    { n:'B2 — Condiciones Generales',        u:'15 reglas',         ok:15, h:0, p:2 },
    { n:'B3 — Gobernanza',                   u:'11 reglas',         ok:8,  h:3, p:1 },
    { n:'B4 — Actas y Orden de Giro',        u:'11 reglas',         ok:9,  h:0, p:2 },
    // PARTE II — Catálogos de Decisión
    { n:'PARTE II · Catálogos de Decisión', sep: true },
    { n:'CEL-02 — Catálogo de Causas',       u:'27 causas',         ok:25, h:2, p:0 },
    { n:'CEL-03 — Catálogo de Juicios',      u:'9 estados',         ok:9,  h:0, p:0 },
    { n:'CEL-04 — Resultados de Decisión',   u:'4 resultados',      ok:4,  h:0, p:1 },
    // PARTE III — Modelo de Score
    { n:'PARTE III · Modelo de Score', sep: true },
    { n:'CEL-05 — Modelos de Score',         u:'1 modelo',          ok:0,  h:0, p:1 },
    { n:'CEL-06 — Variables del Score',      u:'15 variables',      ok:14, h:0, p:1 },
    { n:'CEL-07 — Reglas WOE del Score',     u:'78 reglas WOE',     ok:78, h:0, p:1 },
    { n:'CEL-08 — Evaluaciones del Score',   u:'Lote 17–18/09',     ok:0,  h:0, p:1 },
    { n:'CEL-09 — Evidencias del Score',     u:'>5,000 elementos',  ok:0,  h:0, p:1 },
    { n:'CEL-21 — Equivalencias Ocupación',  u:'160 entradas',      ok:160,h:0, p:1 },
    // PARTE IV — Registro Maestro
    { n:'PARTE IV · Registro Maestro de Decisiones', sep: true },
    { n:'CEL-10 — Registro Maestro',         u:'2,970 registros',   ok:0,  h:0, p:2 },
    // PARTE V — CEL-11 Análisis de Leads
    { n:'PARTE V · Análisis de Leads — CEL-11', sep: true },
    { n:'CEL-11 — Análisis de Leads',        u:'2,783 leads (V1) · 2,379 (V2)', ok:0, h:0, p:3 },
    // PARTE VI — CEL-12 Causas de Rechazo
    { n:'PARTE VI · Causas de Rechazo de Crédito — CEL-12', sep: true },
    { n:'CEL-12 — Causas de Rechazo de Crédito', u:'5,751 filas · 2,129 leads', ok:0, h:1, p:1 },
  ];

  const dataRows = rows.map((r, i) => {
    if (r.sep) {
      return new TableRow({ children: [
        cell(r.n, { w: cols[0]+cols[1]+cols[2]+cols[3]+cols[4], bold: true,
                    bg: PURPLE_DARK, textColor: WHITE, size: SZ_SM, span: 5 }),
      ]});
    }
    const bg = i % 2 === 0 ? WHITE : GRAY_LIGHT;
    const okStr = r.ok > 0 ? String(r.ok) : (r.n.includes('CEL-0') ? '—' : '—');
    return new TableRow({ cantSplit: true, children: [
      cell(r.n, { w: cols[0], bold: false, bg, size: SZ_SM }),
      cell(r.u, { w: cols[1], bg, size: SZ_XS, align: AlignmentType.CENTER }),
      cell(r.ok > 0 ? String(r.ok) : '—', { w: cols[2], bg: r.ok > 0 ? GREEN_BG : bg, textColor: r.ok > 0 ? GREEN_TEXT : TEXT_DARK, size: SZ_SM, bold: r.ok > 0, align: AlignmentType.CENTER }),
      cell(r.h > 0 ? String(r.h) : '—', { w: cols[3], bg: r.h > 0 ? RED_BG : bg, textColor: r.h > 0 ? RED_TEXT : TEXT_DARK, size: SZ_SM, bold: r.h > 0, align: AlignmentType.CENTER }),
      cell(r.p > 0 ? String(r.p) : '—', { w: cols[4], bg: r.p > 0 ? AMBER_BG : bg, textColor: r.p > 0 ? AMBER_TEXT : TEXT_DARK, size: SZ_SM, align: AlignmentType.CENTER }),
    ]});
  });

  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Sección', cols[0]), hCell('Unidades', cols[1]),
        hCell('Conformes', cols[2]), hCell('Hallazgos', cols[3]),
        hCell('Pendientes', cols[4]),
      ]}),
      ...dataRows,
      new TableRow({ children: [
        cell('TOTALES (50 reglas ALCO + catálogos CEL)', { w: cols[0]+cols[1], bold: true, bg: PURPLE_LIGHT, textColor: PURPLE_DARK, size: SZ_SM, span: 2 }),
        cell('—',  { w: cols[2], bold: true, bg: GREEN_BG,  textColor: GREEN_TEXT,  size: SZ_SM, align: AlignmentType.CENTER }),
        cell('6',  { w: cols[3], bold: true, bg: RED_BG,    textColor: RED_TEXT,    size: SZ_SM, align: AlignmentType.CENTER }),
        cell('19', { w: cols[4], bold: true, bg: AMBER_BG,  textColor: AMBER_TEXT,  size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

function makeCEL02Table() {
  const cols = CEL02_COLS;
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Código de causa', cols[0]), hCell('Tipo', cols[1]),
        hCell('Fase', cols[2]),           hCell('Regla relacionada', cols[3]),
        hCell('Estado', cols[4]),         hCell('Observaciones / Hallazgo', cols[5]),
      ]}),
      ...cel02Rules.map((r, i) => {
        const bg = r.e === 'HALLAZGO' ? RED_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.c,    { w: cols[0], bold: true, bg, size: SZ_XS }),
          cell(r.tipo, { w: cols[1], bg, size: SZ_XS, align: AlignmentType.CENTER }),
          cell(r.fase, { w: cols[2], bg, size: SZ_XS }),
          cell(r.rel,  { w: cols[3], bg, size: SZ_XS }),
          statusCell(r.e, cols[4]),
          cell(r.o,    { w: cols[5], bg, size: SZ_XS,
            italic: r.e !== 'HALLAZGO',
            bold: r.e === 'HALLAZGO',
            textColor: r.e === 'HALLAZGO' ? RED_TEXT : TEXT_DARK }),
        ]});
      }),
    ],
  });
}

function makeHallazgosCEL02Table() {
  const cols = [700, 2400, 3060, 1600, 1600];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('N°', cols[0], PURPLE), hCell('Causa(s) afectada(s)', cols[1], PURPLE),
        hCell('Descripción del hallazgo', cols[2], PURPLE),
        hCell('Responsable', cols[3], PURPLE),
        hCell('Acción requerida', cols[4], PURPLE),
      ]}),
      ...hallazgosCEL02.map((h, i) => {
        const bg = i % 2 === 0 ? RED_BG : '#FCEAEA';
        return new TableRow({ children: [
          cell(h.id,    { w: cols[0], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(h.cod,   { w: cols[1], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM }),
          cell(h.desc,  { w: cols[2], bg, size: SZ_SM }),
          cell(h.resp,  { w: cols[3], bg, size: SZ_SM }),
          cell(h.accion,{ w: cols[4], bg, size: SZ_SM }),
        ]});
      }),
    ],
  });
}

function makeHallazgosCEL12Table() {
  const cols = [700, 2400, 3060, 1600, 1600];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('N°', cols[0], PURPLE), hCell('Código(s) afectado(s)', cols[1], PURPLE),
        hCell('Descripción del hallazgo', cols[2], PURPLE),
        hCell('Responsable', cols[3], PURPLE),
        hCell('Acción requerida', cols[4], PURPLE),
      ]}),
      ...hallazgosCEL12.map((h, i) => {
        const bg = i % 2 === 0 ? RED_BG : '#FCEAEA';
        return new TableRow({ children: [
          cell(h.id,    { w: cols[0], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(h.cod,   { w: cols[1], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM }),
          cell(h.desc,  { w: cols[2], bg, size: SZ_SM }),
          cell(h.resp,  { w: cols[3], bg, size: SZ_SM }),
          cell(h.accion,{ w: cols[4], bg, size: SZ_SM }),
        ]});
      }),
    ],
  });
}

function makeCEL12ResumenTable() {
  const cols = [3120, 1240, 1240, 1240, 1240, 1280];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Archivo / Estado', cols[0], PURPLE),
        hCell('Filas', cols[1], PURPLE),
        hCell('Leads únicos', cols[2], PURPLE),
        hCell('Causa vacía', cols[3], PURPLE),
        hCell('Causa ppal vacía', cols[4], PURPLE),
        hCell('PCSC', cols[5], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('v3 — Denegado (4,307) + Rechazado (450)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('4,757', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('1,694', { w: cols[2], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('1,456', { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('111',   { w: cols[4], bg: RED_BG,   textColor: RED_TEXT,   size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('0', { w: cols[5], bg: WHITE, size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
      new TableRow({ children: [
        cell('v4 — Falta Información (985)', { w: cols[0], bold: true, bg: WHITE, size: SZ_SM }),
        cell('985',   { w: cols[1], bg: WHITE, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('435',   { w: cols[2], bg: WHITE, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('233',   { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('2',     { w: cols[4], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('0', { w: cols[5], bg: WHITE, size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
      new TableRow({ children: [
        cell('v5 — Incidencia técnica (9)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('9',   { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('9',   { w: cols[2], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('0',   { w: cols[3], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('0',   { w: cols[4], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('0', { w: cols[5], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
      new TableRow({ children: [
        cell('TOTAL', { w: cols[0], bold: true, bg: PURPLE_LIGHT, textColor: PURPLE_DARK, size: SZ_SM }),
        cell('5,751', { w: cols[1], bold: true, bg: PURPLE_LIGHT, textColor: PURPLE_DARK, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('2,129', { w: cols[2], bold: true, bg: PURPLE_LIGHT, textColor: PURPLE_DARK, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('1,689', { w: cols[3], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('113',   { w: cols[4], bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('0', { w: cols[5], bold: true, bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

function makeCEL12CausasTable() {
  const cols = [3600, 1400, 1400, 2960];
  const causas = [
    { c:'RELACION_CREDITICIA_NO_ADMISIBLE', n:502,  pct:'22.1%', tipo:'Rechazo' },
    { c:'SIN_SEGMENTO_CERO_ESTRES',         n:353,  pct:'15.5%', tipo:'Rechazo' },
    { c:'INTANGIBLE_ULTIMO_MES_NO_CUMPLE',  n:205,  pct:'9.0%',  tipo:'Rechazo' },
    { c:'CATEGORIA_INTERNA_NO_ADMISIBLE',   n:195,  pct:'8.6%',  tipo:'Rechazo' },
    { c:'DATOS_PERFIL_FALTANTES',           n:182,  pct:'8.0%',  tipo:'Técnica' },
    { c:'(vacía — Incidencia de catálogo)', n:113,  pct:'5.0%',  tipo:'Sin código' },
    { c:'SCORE_INTERNO_NO_APROBADO',        n:97,   pct:'4.3%',  tipo:'Rechazo' },
    { c:'BASE_VENCIDA_60',                  n:96,   pct:'4.2%',  tipo:'Rechazo' },
    { c:'CUOTA_PROYECTADA_SUPERA_CIC',      n:83,   pct:'3.6%',  tipo:'Rechazo' },
    { c:'INGRESO_NO_SUPERA_CUOTAS_CIC',     n:77,   pct:'3.4%',  tipo:'Rechazo' },
  ];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Causa principal (Verdadero)', cols[0], PURPLE),
        hCell('Registros', cols[1], PURPLE),
        hCell('% del total (2,275)', cols[2], PURPLE),
        hCell('Tipo CEL-02', cols[3], PURPLE),
      ]}),
      ...causas.map((r, i) => {
        const isVacia = r.c.includes('vacía');
        const bg = isVacia ? RED_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        const tc = isVacia ? RED_TEXT : TEXT_DARK;
        return new TableRow({ children: [
          cell(r.c,   { w: cols[0], bold: isVacia, bg, textColor: tc, size: SZ_SM }),
          cell(String(r.n),   { w: cols[1], bg, textColor: tc, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.pct, { w: cols[2], bg, textColor: tc, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.tipo,{ w: cols[3], bg, textColor: tc, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

function makeCEL02ResumenTable() {
  const cols = [4680, 1560, 1560, 1560];
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: cols,
    layout: TableLayoutType.FIXED,
    borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Catálogo', cols[0], PURPLE), hCell('Causas revisadas', cols[1], PURPLE),
        hCell('Conformes', cols[2], PURPLE), hCell('Hallazgos', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('CEL-02 — Catálogo de Causas (versión ALCO-2026.08.17.1)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('27', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('24', { w: cols[2], bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('3',  { w: cols[3], bg: RED_BG,   textColor: RED_TEXT,   size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
      ]}),
    ],
  });
}

// ─── CEL-11 DATA & TABLES ────────────────────────────────────────────────

// Vista 2 (2,379 registros) — distribución por versión de matriz
const cel11VersionData = [
  { ver:'CEL-01-PRODUCCION-V1.0',        n:1590, pct:'66.8%', tipo:'Producción',   obs:'Versión auditada — nomenclatura CEL-04' },
  { ver:'CE-REGLAS-ALCO-2026-V2.3',      n:414,  pct:'17.4%', tipo:'Pre-prod.',    obs:'No documentada en CEL-05 — ver P-CEL10-002' },
  { ver:'CE-REGLAS-ALCO-2026-V2.4',      n:375,  pct:'15.8%', tipo:'Recuperación', obs:'NUEVO · flujo CEF-2-RECUPERACION · 24–29/08/2026 · P-CEL11-001' },
];

// Vista 1 (2,783 registros) — distribución de resultados por nomenclatura
const cel11ResultadosData = [
  { res:'Denegado',             n:1731, pct:'62.2%', nota:'Nomenclatura CEL-04 (producción)' },
  { res:'Aprobado',             n:565,  pct:'20.3%', nota:'Nomenclatura CEL-04' },
  { res:'Rechazado',            n:259,  pct:'9.3%',  nota:'Nomenclatura pre-ALCO · lote V2.4 · 24–29/08/2026' },
  { res:'Aprobado condicionado',n:127,  pct:'4.6%',  nota:'Nomenclatura CEL-04' },
  { res:'Falta Información',    n:95,   pct:'3.4%',  nota:'Nomenclatura CEL-04 (reprocesable)' },
  { res:'Incidencia técnica',   n:4,    pct:'0.1%',  nota:'Nomenclatura CEL-04' },
  { res:'Aprobado final',       n:2,    pct:'0.1%',  nota:'Nomenclatura pre-ALCO · lote V2.4' },
];

// Vista 1 — causas y eventos principales
const cel11CausasData = [
  { causa:'SIN_REGLA — lead descartado en pre-screening, sin evaluación formal del motor', n:1618, pct:'58.1%', rel:'Pre-evaluación — sin causa CEL-02 asignada' },
  { causa:'No se identificaron operaciones que generen ahorro y monto positivo',           n:1132, pct:'40.7%', rel:'CEL-01 Seg. B/C/D · operaciones a consolidar' },
  { causa:'Falta el estado de la PCSC',                                                   n:308,  pct:'11.1%', rel:'H-003 · ALCO-GOB-008 · DOCCORP-0719-2026' },
  { causa:'Vigencia de base de datos "No cumple" (datos >60 días)',                       n:66,   pct:'2.4%',  rel:'ALCO-GEN-008 · BASE_VENCIDA_60' },
];

function makeCEL11ResumenTable() {
  const cols = [5400, 1080, 960, 1920];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Vista', cols[0], PURPLE), hCell('Registros', cols[1], PURPLE),
        hCell('Columnas', cols[2], PURPLE), hCell('Observación', cols[3], PURPLE),
      ]}),
      new TableRow({ children: [
        cell('Vista 1 — Simplificada (todos los leads, incluye pre-screening SIN_REGLA)', { w: cols[0], bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
        cell('2,783', { w: cols[1], bg: GRAY_LIGHT, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('17',    { w: cols[2], bg: GRAY_LIGHT, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('1,618 SIN_REGLA · 308 causa PCSC · H-003 evidenciado', { w: cols[3], bg: GRAY_LIGHT, size: SZ_SM, italic: true }),
      ]}),
      new TableRow({ children: [
        cell('Vista 2 — Detallada (leads con evaluación formal; 404 más recientes solo en V1)', { w: cols[0], bold: true, bg: WHITE, size: SZ_SM }),
        cell('2,379', { w: cols[1], bg: WHITE, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
        cell('79',    { w: cols[2], bg: WHITE, size: SZ_SM, align: AlignmentType.CENTER }),
        cell('V2.4 nueva · 16 flujos · media score 180.0 · 3 nuevos pendientes', { w: cols[3], bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ_SM, italic: true }),
      ]}),
    ],
  });
}

function makeCEL11VersionTable() {
  const cols = [2960, 600, 840, 1200, 3760];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Versión de matriz (Vista 2)', cols[0], PURPLE),
        hCell('Total', cols[1], PURPLE), hCell('%', cols[2], PURPLE),
        hCell('Tipo', cols[3], PURPLE), hCell('Observación', cols[4], PURPLE),
      ]}),
      ...cel11VersionData.map((r, i) => {
        const isNew  = r.ver.includes('V2.4');
        const isProd = r.ver === 'CEL-01-PRODUCCION-V1.0';
        const bg = isNew ? AMBER_BG : (isProd ? PURPLE_LIGHT : (i % 2 === 0 ? WHITE : GRAY_LIGHT));
        return new TableRow({ children: [
          cell(r.ver,   { w: cols[0], bold: isProd || isNew, bg, size: SZ_SM }),
          cell(String(r.n), { w: cols[1], bg, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.pct,   { w: cols[2], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.tipo,  { w: cols[3], bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.obs,   { w: cols[4], bg: isNew ? AMBER_BG : bg, textColor: isNew ? AMBER_TEXT : TEXT_DARK, size: SZ_SM, italic: !isNew }),
        ]});
      }),
    ],
  });
}

function makeCEL11ResultadosTable() {
  const cols = [2400, 640, 840, 5480];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Resultado (Vista 1 · 2,783 leads)', cols[0], PURPLE),
        hCell('N', cols[1], PURPLE), hCell('%', cols[2], PURPLE),
        hCell('Nota', cols[3], PURPLE),
      ]}),
      ...cel11ResultadosData.map((r, i) => {
        const isPreAlco = r.nota.includes('pre-ALCO');
        const isApro    = r.res === 'Aprobado';
        const bg = isPreAlco ? AMBER_BG : (isApro ? GREEN_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT));
        const tc = isPreAlco ? AMBER_TEXT : (isApro ? GREEN_TEXT : TEXT_DARK);
        return new TableRow({ children: [
          cell(r.res,        { w: cols[0], bold: true, bg, textColor: tc, size: SZ_SM }),
          cell(String(r.n),  { w: cols[1], bg, textColor: tc, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.pct,        { w: cols[2], bg, textColor: tc, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.nota,       { w: cols[3], bg, size: SZ_SM, italic: true }),
        ]});
      }),
    ],
  });
}

function makeCEL11CausasTable() {
  const cols = [4480, 600, 840, 3440];
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: cols,
    layout: TableLayoutType.FIXED, borders: TABLE_BORDERS,
    rows: [
      new TableRow({ tableHeader: true, children: [
        hCell('Causa / Evento (Vista 1 · 2,783 leads)', cols[0], PURPLE),
        hCell('N', cols[1], PURPLE), hCell('%', cols[2], PURPLE),
        hCell('Relación normativa', cols[3], PURPLE),
      ]}),
      ...cel11CausasData.map((r, i) => {
        const isPCSC     = r.causa.includes('PCSC');
        const isSinRegla = r.causa.includes('SIN_REGLA');
        const bg = isPCSC ? AMBER_BG : (i % 2 === 0 ? WHITE : GRAY_LIGHT);
        return new TableRow({ children: [
          cell(r.causa,      { w: cols[0], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM, bold: isPCSC }),
          cell(String(r.n),  { w: cols[1], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM, bold: true, align: AlignmentType.CENTER }),
          cell(r.pct,        { w: cols[2], bg: isPCSC ? AMBER_BG : bg, size: SZ_SM, align: AlignmentType.CENTER }),
          cell(r.rel,        { w: cols[3], bg: isPCSC ? AMBER_BG : bg, textColor: isPCSC ? AMBER_TEXT : TEXT_DARK, size: SZ_SM }),
        ]});
      }),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────
// ENCABEZADO Y PIE DE PÁGINA
// ─────────────────────────────────────────────────────────────────────────
const docHeader = new Header({
  children: [
    new Table({
      width: { size: TW, type: WidthType.DXA },
      columnWidths: [6800, 2560],
      layout: TableLayoutType.FIXED,
      borders: { top:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE}, insideH:{style:BorderStyle.NONE}, insideV:{style:BorderStyle.NONE},
        bottom:{ style:BorderStyle.SINGLE, size: 8, color: ORANGE } },
      rows: [new TableRow({ children: [
        new TableCell({
          children: [
            para([run('BANCO POPULAR Y DE DESARROLLO COMUNAL', { size: 18, bold: true, color: ORANGE })], { after: 20 }),
            para([run('Dirección Banca Digital', { size: 16, color: '777777' })], { after: 0 }),
          ],
          width: { size: 6800, type: WidthType.DXA },
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        }),
        new TableCell({
          children: [para([run('Setiembre 2026', { size: 16, color: '777777' })], { align: AlignmentType.RIGHT, after: 0 })],
          width: { size: 2560, type: WidthType.DXA },
          verticalAlign: VerticalAlign.BOTTOM,
          borders: { top:{style:BorderStyle.NONE}, bottom:{style:BorderStyle.NONE}, left:{style:BorderStyle.NONE}, right:{style:BorderStyle.NONE} },
        }),
      ]})],
    }),
    gap(40),
  ],
});

const docFooter = new Footer({
  children: [new Paragraph({
    children: [
      run('CONFIDENCIAL — USO INTERNO  ·  Informe de Auditoría Cero Estrés CEL-01-PRODUCCION-V1.0  ·  Página ', { size: 16, color: '999999' }),
      new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '999999' }),
      run(' de ', { size: 16, color: '999999' }),
      new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 16, color: '999999' }),
    ],
    alignment: AlignmentType.CENTER,
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: GRAY_MID } },
    spacing: { before: 80, after: 0 },
  })],
});

// ─────────────────────────────────────────────────────────────────────────
// DOCUMENTO
// ─────────────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: { default: { document: { run: { font: FONT, size: SZ, color: TEXT_DARK } } } },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1080, bottom: 1080, left: 1440, right: 1440 } },
    },
    headers:  { default: docHeader },
    footers:  { default: docFooter },
    children: [

      // ─────────────── PORTADA ───────────────────────────────────────────
      gap(240),
      new Paragraph({
        children: [run('INFORME DE AUDITORÍA', { size: 40, bold: true, color: ORANGE })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      }),
      new Paragraph({
        children: [run('Producto Cero Estrés — Línea de Crédito Consolidación', { size: 26, bold: true })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 80 },
      }),
      new Paragraph({
        children: [run('Revisión Matriz CEL-01-PRODUCCION-V1.0 vs. Acuerdo ALCO No. 22-art.02 Acd.02 ALCO-2026', { size: SZ, color: '555555', italics: true })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 },
      }),
      new Paragraph({
        children: [run('Firmado: 17 de agosto de 2026  ·  Fernando Gutiérrez Marín', { size: SZ_SM, color: '999999' })],
        alignment: AlignmentType.CENTER, spacing: { before: 0, after: 280 },
      }),
      divider(),

      // ─────────────── FICHA TÉCNICA ─────────────────────────────────────
      sectionTitle('FICHA TÉCNICA DEL DOCUMENTO'),
      new Table({
        width: { size: TW, type: WidthType.DXA },
        columnWidths: [2600, 6760],
        layout: TableLayoutType.FIXED,
        borders: TABLE_BORDERS,
        rows: [
          ['Versión del documento', 'V2.2 — Pivote editorial cualitativo · hilo conductor CEL-10/11/12 · narrativa entradas-salidas-continuidad · correcciones de consistencia · 19 pendientes · 6 hallazgos'],
          ['Fecha de elaboración', 'Setiembre 2026'],
          ['Elaborado por', 'Dirección Banca Digital'],
          ['Destinatario', 'Jefatura Dirección Banca Digital'],
          ['Matriz auditada', 'CEL-01-PRODUCCION-V1.0'],
          ['Fuente normativa', 'Acuerdo ALCO No. 22-art.02 Acd.02 ALCO-2026 (17/08/2026)'],
          ['Estado del producto', 'En fase de pruebas desde la firma del ALCO'],
          ['Clasificación', 'CONFIDENCIAL — USO INTERNO'],
        ].map(([k, v], i) => new TableRow({ children: [
          cell(k, { w: 2600, bold: true, bg: GRAY_LIGHT, size: SZ_SM }),
          k === 'Clasificación'
            ? cell(v, { w: 6760, bold: true, textColor: ORANGE, size: SZ_SM })
            : cell(v, { w: 6760, size: SZ_SM }),
        ]})),
      }),
      gap(200),
      divider(),

      // ─────────────── RESUMEN EJECUTIVO ────────────────────────────────
      sectionTitle('RESUMEN EJECUTIVO'),
      sectionDesc('Objetivo', 'Documentar de forma proactiva y preventiva la alineación de la Matriz de Reglas de Negocio CEL-01-PRODUCCION-V1.0 del producto Cero Estrés (Línea de Crédito Consolidación) contra el Acuerdo ALCO No. 22-art.02 Acd.02 ALCO-2026, firmado el 17 de agosto de 2026. La finalidad es garantizar que no existan cabos sueltos ni errores antes de la entrada en operación formal del producto, y habilitar el trabajo anticipado sobre los puntos accionables identificados.'),
      sectionDesc('Alcance y metodología', 'Revisión interna de la Dirección Banca Digital, bloque por bloque, de las 50 reglas de la matriz. Se confronta cada regla, condición, monto, tasa y obligación de gobernanza contra el texto del acuerdo ALCO. En esta versión se integra como contexto adicional la Arquitectura de Datos CEA-02 (modelo de 4 capas: DIRBDLEADS → gobierno de reglas/score → ejecución → actas/expediente) y el Modelo de control documental y consulta CIC CEL-27 (Operador documental SUGEF, Azure 1.6), verificando que la infraestructura de datos soporta las reglas auditadas. Los hallazgos identificados son oportunidades de mejora que la Dirección puede gestionar de forma anticipada; los pendientes son ítems de respaldo normativo o aclaración arquitectural aún pendientes, sin que ninguno de ellos constituya un error en la definición del producto.'),
      sectionDesc('Periodo de referencia', 'El producto se encuentra en fase de pruebas desde la firma del ALCO (17/08/2026). Las actas generales e individuales se generan con el fin de validar que el mecanismo de aprobación masiva es idóneo antes de la operación plena.'),
      gap(120),
      makeResumenTable(),
      gap(160),

      // Cuadros de estado
      new Table({
        width: { size: TW, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        layout: TableLayoutType.FIXED,
        borders: TABLE_BORDERS,
        rows: [new TableRow({ children: [
          cell('✔ 50 reglas CONFORMES', { w: 3120, bold: true, bg: GREEN_BG, textColor: GREEN_TEXT, size: SZ, align: AlignmentType.CENTER }),
          cell('● 6 HALLAZGOS (3 ALCO · 3 técnicos)', { w: 3120, bold: true, bg: RED_BG, textColor: RED_TEXT, size: SZ, align: AlignmentType.CENTER }),
          cell('⏳ 19 PENDIENTES operativos', { w: 3120, bold: true, bg: AMBER_BG, textColor: AMBER_TEXT, size: SZ, align: AlignmentType.CENTER }),
        ]})],
      }),
      gap(160),

      nota('Los seis hallazgos se dividen en dos categorías: H-001, H-002 y H-003 son incumplimientos directos del Acuerdo ALCO (gobernanza no implementada); H-CEL02-01, H-CEL02-02 y H-CEL12-01 son hallazgos técnicos de sincronización entre catálogos (H-CEL02-01: 2 causas sin respaldo normativo visible; H-CEL02-02: 1 causa observada pendiente; H-CEL12-01: 6 causas del motor sin entrada en CEL-02, 1,685 filas afectadas). En todos los casos, el motor CEL-01-PRODUCCION-V1.0 opera correctamente: los hallazgos técnicos reflejan brechas de documentación, no errores de ejecución. Los 19 pendientes operativos no son incumplimientos; requieren confirmación de respaldo normativo, evidencia durante la fase de pruebas, aclaración arquitectural o gestión de gobernanza (P-001: vigencia JDN-5824; P-003: PCSC en CEA-02; P-005: Operador SUGEF; P-CEL04-001: atraso CE-RES-02; P-CEL05-001: mantenimiento catálogo; P-CEL06-001: cobertura N_ORI_FND_G_CM; P-CEL07-001: WOE diseño; P-CEL08-001: campo I_CUOTAS_BP; P-CEL09-001: umbral SharePoint; P-CEL21-001: visto bueno Riesgos; P-CEL10-001: 423 registros pre-ALCO sin versión; P-CEL10-002: 439 registros bajo CE-REGLAS-ALCO-2026-V2.x; P-CEL11-001: V2.4 no documentada en CEL-05; P-CEL11-002: flujo QA activo en base de producción; P-CEL11-003: formato inconsistente de tasa; P-CEL12-001: 113 registros con causa principal vacía por desincronización CEL-01/CEL-02).'),

      divider(),
      pageBreakPara(),

      // ═══════════════════════════════════════════════════════════════════
      //  ÍNDICE
      // ═══════════════════════════════════════════════════════════════════
      sectionTitle('ÍNDICE', ORANGE),
      divider(PURPLE),
      gap(80),
      indiceRow('I',   'PARTE I — Revisión Normativa: Acuerdo ALCO', '1.1 Bloque 1 — Reglas de Crédito  ·  1.2 Bloque 2 — Condiciones Generales  ·  1.3 Bloque 3 — Gobernanza  ·  1.4 Bloque 4 — Actas y Orden de Giro  ·  1.5 Hallazgos Formales de Incumplimiento'),
      gap(80),
      indiceRow('II',  'PARTE II — Catálogos de Decisión', '2.1 CEL-02 — Catálogo de Causas  ·  2.2 CEL-03 — Catálogo de Juicios  ·  2.3 CEL-04 — Catálogo de Resultados de Decisión'),
      gap(80),
      indiceRow('III', 'PARTE III — Modelo de Score', '3.1 CEL-05 — Modelos de Score  ·  3.2 CEL-06 — Variables del Score  ·  3.3 CEL-07 — Reglas WOE  ·  3.4 CEL-08 — Evaluaciones  ·  3.5 CEL-09 — Evidencias  ·  3.6 CEL-21 — Equivalencias Ocupación CREDID'),
      gap(80),
      indiceRow('IV',  'PARTE IV — Registro Maestro de Decisiones', '4.1 CEL-10 — Registro Maestro de Decisiones (2,970 registros)'),
      gap(80),
      indiceRow('V',   'PARTE V — Análisis de Leads — CEL-11', '5.1 CEL-11 — Vista 1 (2,783 leads) · Vista 2 (2,379 leads evaluados) · 3 nuevos pendientes: P-CEL11-001/002/003'),
      gap(80),
      indiceRow('VI',  'PARTE VI — Causas de Rechazo de Crédito — CEL-12', '6.1 CEL-12 — 5,751 filas · 2,129 leads · 3 archivos: v3 Denegado+Rechazado · v4 Falta Info · v5 Incidencia · 1 hallazgo H-CEL12-01 · 1 pendiente P-CEL12-001'),
      gap(80),
      indiceRow('VII', 'PARTE VII — Ítems Pendientes Consolidados', '19 pendientes de respaldo normativo, arquitecturales y de gobernanza'),
      gap(200),
      divider(),
      pageBreakPara(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE I — REVISIÓN NORMATIVA
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('I', 'Revisión Normativa — Acuerdo ALCO',
        'Bloques B1–B4 · 50 reglas auditadas · 3 hallazgos formales · Hallazgos formales al cierre de la parte'),

      // ─────────────── BLOQUE 1 ─────────────────────────────────────────
      sectionTitle('1.1 · BLOQUE 1 — REGLAS DE CRÉDITO'),
      sectionDesc('Alcance', '13 reglas (REGLA-001 a REGLA-013) · Segmentos A, B, C y D · No clientes BP y clientes BP (asalariados, ingresos propios, pensionados)'),
      sectionDesc('Resultado', '13 conformes · 0 hallazgos · 2 pendientes de respaldo normativo (P-001, P-002)'),
      gap(80),
      makeB1Table(),
      gap(120),
      para([
        run('Leyenda: ', { bold: true, size: SZ_SM }),
        run('✔ CONFORME ', { bold: true, color: GREEN_TEXT, size: SZ_SM }),
        run('= alineado al ALCO.  ', { size: SZ_SM }),
        run('⏳ PENDIENTE ', { bold: true, color: AMBER_TEXT, size: SZ_SM }),
        run('= requiere respaldo normativo adicional.', { size: SZ_SM }),
      ], { after: 160 }),
      divider(),
      pageBreakPara(),

      // ─────────────── BLOQUE 2 ─────────────────────────────────────────
      sectionTitle('1.2 · BLOQUE 2 — CONDICIONES GENERALES'),
      sectionDesc('Alcance', '15 reglas (ALCO-GEN-001 a ALCO-GEN-015) · Condiciones transversales: autorización CIC, capacidad de pago, garantías, vigencia, no reevaluación y Score Interno'),
      sectionDesc('Resultado', '15 conformes · 0 hallazgos · 2 pendientes (P-001: corte Score Interno · P-005: aceptación formal Operador CIC)'),
      gap(80),
      makeB2Table(),
      gap(160),
      divider(),
      pageBreakPara(),

      // ─────────────── BLOQUE 3 ─────────────────────────────────────────
      sectionTitle('1.3 · BLOQUE 3 — GOBERNANZA', PURPLE),
      sectionDesc('Alcance', '11 reglas (ALCO-GOB-001 a ALCO-GOB-011) · Asignación de responsabilidades institucionales para la operación del mecanismo de aprobación masiva'),
      sectionDesc('Resultado', '8 conformes · 3 HALLAZGOS formales de incumplimiento · 1 pendiente arquitectural (P-003: PCSC en CEA-02)'),
      gap(80),
      makeB3Table(),
      gap(120),
      nota('GOB-003, GOB-005 y GOB-008 representan incumplimientos directos del ALCO. Las demás reglas con verificación de evidencia quedan sustentadas en el propio texto del acuerdo. Con la integración de CEA-02: la verificación PCSC y de listas (GOB-009/010/011) no aparece como capa explícita en el modelo de datos documentado → se solicita aclaración a Jefatura (P-003).'),
      divider(),
      pageBreakPara(),

      // ─────────────── BLOQUE 4 ─────────────────────────────────────────
      sectionTitle('1.4 · BLOQUE 4 — ACTAS Y ORDEN DE GIRO'),
      sectionDesc('Alcance', '11 reglas (ALCO-ACT-001 a ALCO-GIR-003) · Contenido y emisión de actas generales e individuales, tipos de aprobación y requisitos de la orden de giro'),
      sectionDesc('Resultado', '9 conformes · 0 hallazgos · 2 pendientes operativos (fase de pruebas)'),
      gap(80),
      makeB4Table(),
      gap(120),
      nota('El producto se encuentra en fase de pruebas desde agosto 2026. Los pendientes operativos (ACT-004, GIR-002) no son hallazgos; se resolverán en la medida en que avance la operación y se estandaricen los procedimientos.'),
      divider(),
      pageBreakPara(),

      // ─────────────── HALLAZGOS FORMALES (cierre PARTE I) ──────────────
      sectionTitle('1.5 · HALLAZGOS FORMALES DE INCUMPLIMIENTO', PURPLE),
      para([run('Los siguientes hallazgos representan obligaciones establecidas en el Acuerdo ALCO que no han sido cumplidas a la fecha de este informe:', { size: SZ })], { after: 120 }),
      makeHallazgosTable(),
      gap(160),
      divider(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE II — CATÁLOGOS DE DECISIÓN
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('II', 'Catálogos de Decisión',
        'CEL-02 · CEL-03 · CEL-04 — Causas, juicios y resultados del motor de aprobación automática'),

      // ─────────────── CEL-02 — CATÁLOGO DE CAUSAS ─────────────────────
      sectionTitle('2.1 · REVISIÓN CEL-02 — CATÁLOGO DE CAUSAS', PURPLE),
      sectionDesc('Alcance', '27 causas · Versión ALCO-2026.08.17.1 (base) + CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181 (25/08/2026) · Tipos: Rechazo (decisión final) e Incidencia técnica (evaluación fallida, generalmente reprocesable)'),
      sectionDesc('Resultado', '24 conformes · 2 hallazgos nuevos (H-CEL02-01: 2 causas sin respaldo normativo visible · H-CEL02-02: 1 causa observada pendiente de clasificación) · Evidencias parciales para P-001 y P-003'),
      para([run('El catálogo CEL-02 registra el motivo de cada resultado negativo del motor de aprobación automática. Esta revisión valida que cada causa tiene respaldo en las reglas CEL-01 y el Acuerdo ALCO, e identifica causas que aplican criterios sin referencia normativa visible.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL02ResumenTable(),
      gap(160),
      makeCEL02Table(),
      gap(120),
      nota('Las causas de tipo "Incidencia técnica" son reprocesables: cuando se corrige la condición que las originó (datos faltantes, autorización CIC vencida, etc.) el motor puede reevaluar el caso. Las causas de tipo "Rechazo" son definitivas para la versión de reglas vigente, salvo que el producto cambie de segmento o condiciones.'),
      divider(),
      pageBreakPara(),

      // ─────────────── HALLAZGOS CEL-02 ─────────────────────────────────
      sectionTitle('2.1.1 · HALLAZGOS CEL-02 — CAUSAS SIN RESPALDO NORMATIVO VISIBLE', PURPLE),
      para([run('Los siguientes hallazgos identifican causas en el catálogo CEL-02 que aplican criterios no referenciados en CEL-01 ni en el Acuerdo ALCO, o con inconsistencias internas que deben resolverse:', { size: SZ })], { after: 120 }),
      makeHallazgosCEL02Table(),
      gap(120),
      nota('H-CEL02-01 (intangible dual) podría resolverse si Jefatura confirma que existe una normativa SUGEF o lineamiento interno de BP que establezca el cálculo de intangible. De ser así, el próximo paso es incluirlo como regla en CEL-01 y crear la referencia cruzada. H-CEL02-02 requiere únicamente consistencia de texto en la misma causa — no implica cambio de criterio si la lógica del motor ya es correcta.'),
      divider(),
      pageBreakPara(),


      // ─────────────── CEL-03 — CATÁLOGO DE JUICIOS ────────────────────
      sectionTitle('2.2 · REVISIÓN CEL-03 — CATÁLOGO DE JUICIOS', PURPLE),
      sectionDesc('Alcance', '9 entradas (JUI-001 a JUI-999) · Versión ALCO-2026.08.17.1 · Vigencia desde 16/08/2026 · Aplica a segmentos A, B, C y D · Evalúa únicamente casos donde el cliente figura como demandado'),
      sectionDesc('Resultado', '9 conformes · 0 hallazgos · Respaldo normativo formal en fuentes de criterio'),
      para([run('El catálogo CEL-03 define cuáles estados de juicio en el sistema CREDID se consideran activos (bloquean la aprobación automática), cuáles son cerrados (admisibles) y cuáles requieren reproceso por datos incompletos. Operacionaliza directamente ALCO-GEN-006 (no se admiten juicios activos ni embargos en Protectora).', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL03ResumenTable(),
      gap(160),
      makeCEL03Table(),
      gap(120),
      nota('Todas las entradas tienen respaldo normativo formal documentado en la FuenteCriterio del catálogo. TipoParte = "Demandado" únicamente: el catálogo no evalúa cuando el cliente figura como actor (demandante), lo cual es correcto por diseño — el riesgo crediticio relevante es ser demandado, no demandar. La entrada JUI-999 actúa como guardia de catálogo y es una buena práctica: cualquier estado no enumerado queda como incidencia técnica, impidiendo aprobaciones automáticas ante estados desconocidos.'),
      divider(),
      pageBreakPara(),


      // ─────────────── CEL-04 — CATÁLOGO DE RESULTADOS DE DECISIÓN ────────
      sectionTitle('2.3 · REVISIÓN CEL-04 — CATÁLOGO DE RESULTADOS DE DECISIÓN', PURPLE),
      sectionDesc('Alcance', '4 resultados (CE-RES-01 a CE-RES-04) · Versión CE-RESULTADOS-2026.08.29.1 · Nota: versión fechada 12 días después del Acuerdo ALCO (17/08/2026), siendo el primer catálogo del Centro de Auditoría posterior a la firma · Vigencia aplicable a todos los segmentos'),
      sectionDesc('Resultado', '3 conformes · 0 hallazgos formales · 1 pendiente (P-CEL04-001: condición de atraso en CE-RES-02 vs. segmentación B1)'),
      para([run('El catálogo CEL-04 define los cuatro posibles resultados de la decisión del motor de aprobación automática: Aprobado, Aprobado Condicionado, Denegado y Falta Información. CE-RES-04 tiene prioridad de reproceso sobre los demás. CE-RES-02 enriquece el entendimiento de P-003 al especificar los estados PCSC admisibles para aprobación condicionada (Activo, Activo CES, Expediente Simplificado).', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL04ResumenTable(),
      gap(160),
      makeCEL04Table(),
      gap(120),
      nota('La versión CE-RESULTADOS-2026.08.29.1, fechada 12 días después del ALCO, sugiere que la formalización detallada de las condiciones de aprobación condicionada (CE-RES-02) se produjo en el período posterior a la firma del acuerdo. Esto es admisible y consistente con ALCO-ACT-007. La condición de atraso (ítem 3 de CE-RES-02) opera en una capa distinta a la segmentación B1: aplica a las operaciones específicas a cancelar con Cero Estrés dentro del período de 60 días, no al historial SUGEF del cliente (ya evaluado antes). La consolidación resuelve la condición por diseño del producto (ALCO-GEN-013). Interpretación verificada con Jefatura. Pendiente confirmación formal para cierre de P-CEL04-001. CE-RES-02 también documenta que el estado "No cliente" es admisible cuando la persona no posee PCSC, complementando el entendimiento de P-003.'),
      divider(),
      pageBreakPara(),


      // ═══════════════════════════════════════════════════════════════════
      //  PARTE III — MODELO DE SCORE
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('III', 'Modelo de Score',
        'CEL-05 · CEL-06 · CEL-07 · CEL-08 · CEL-09 · CEL-21 — Gobernanza, variables, reglas WOE, evaluaciones y equivalencias de ocupación'),

      // ─────────────── CEL-05 — MODELOS DE SCORE ───────────────────────────
      sectionTitle('3.1 · REVISIÓN CEL-05 — CATÁLOGO DE MODELOS DE SCORE', PURPLE),
      sectionDesc('Alcance', '1 modelo revisado · Versión en catálogo: CE-SCORE-CONSUMO-V1.0-DRAFT · Versión de producción referenciada en CEL-02: V1.1-SOLO-NO-CLIENTE-CORTE181 (vigente 25/08/2026) · Rango teórico: 100–315 pts · 15 variables'),
      sectionDesc('Resultado', '0 conformes plenos · 1 pendiente (P-CEL05-001: mantenimiento del catálogo) · Evidencia normativa aportada para P-001'),
      para([run('El catálogo CEL-05 documenta los modelos de score utilizados en la segmentación del producto. Esta revisión valida el corte aprobado, el instrumento normativo de respaldo y el estado operativo del modelo.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL05ResumenTable(),
      gap(160),
      makeCEL05Table(),
      gap(120),
      makeCEL05ObsTable(),
      gap(120),
      nota('CEL-05 aporta evidencia normativa clave para P-001: el instrumento JDN-5824-Acd-380-2021-Art-9 es el respaldo formal del corte 181 (Consumo 181-200; zona gris 161-180 no autorizada). El puntaje crudo se conserva en PIB para auditoría interna aunque no está expuesto en el modelo semántico — decisión de diseño que protege la integridad del proceso. El ALCO-2026 delega correctamente la gobernanza del corte numérico a este instrumento (ALCO-GEN-014). Queda pendiente verificar vigencia del JDN para el alcance específico de Cero Estrés 2026. Los dos puntos de mantenimiento (versión y estado operativo) fueron consultados a Jefatura y están en proceso de actualización.'),
      divider(),
      pageBreakPara(),

      // ─────────────── CEL-06 — VARIABLES DEL SCORE ────────────────────
      sectionTitle('3.2 · REVISIÓN CEL-06 — VARIABLES DEL SCORE', PURPLE),
      sectionDesc('Alcance', '15 variables · Corte de datos: 17/09/2026 · Muestra: 406 casos de base 2,723 · Versión referenciada: CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181'),
      sectionDesc('Resultado', '14 variables con cobertura ≥ 80% · 1 pendiente de cobertura (P-CEL06-001: N_ORI_FND_G_CM 17.34%)'),
      para([run('El catálogo CEL-06 registra las métricas de cobertura de cada variable predictora del Score Interno sobre la muestra de leads activos. Esta revisión verifica que las coberturas operativas son consistentes con el diseño del modelo y que las variables con baja cobertura tienen una causa identificada.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL06ResumenTable(),
      gap(160),
      makeCEL06Table(),
      gap(120),
      nota('La variable N_ORI_FND_G_CM (Origen de fondos) presenta 17.34% de cobertura en la muestra de 406 casos: 335 casos MISSING y solo 71 con valor. El campo se origina en un módulo de incorporación no habilitado para el flujo Cero Estrés (no clientes BP), por lo que su ausencia es sistemática y no aleatoria. El modelo asigna el grupo WOE del segmento MISSING cuando el campo es nulo (14 pts, ver CEL-07). Las demás 14 variables presentan cobertura ≥ 80%.'),
      divider(),
      pageBreakPara(),

      // ─────────────── CEL-07 — REGLAS DEL SCORE (WOE) ─────────────────
      sectionTitle('3.3 · REVISIÓN CEL-07 — REGLAS DEL SCORE (WOE)', PURPLE),
      sectionDesc('Alcance', '78 reglas WOE · 15 variables · Instrumento normativo: JDN-5824-Acd-380-2021-Art-9 · Versión: CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181'),
      sectionDesc('Resultado', '78 reglas conformes con JDN-5824 · 1 pendiente de diseño de bajo impacto (P-CEL07-001)'),
      para([run('El catálogo CEL-07 contiene las 78 reglas WOE (Weight of Evidence) distribuidas en 15 variables. Cada regla asigna un puntaje parcial según el grupo en que cae el valor de la variable; el puntaje total es la suma de los 15 parciales. Esta revisión verifica la consistencia interna de las reglas (monotonía WOE esperada, ausencia de grupos huérfanos) y su alineación con JDN-5824.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL07ResumenTable(),
      gap(160),
      makeCEL07Table(),
      gap(120),
      nota('P-CEL07-001 (Bajo impacto): (1) B_PAGOPLANILLA es la única variable con puntaje WOE negativo (–4 pts para valor = 0), lo cual es matemáticamente válido pero conviene documentar en el diccionario de variables para auditorías futuras. (2) PIB BCCR MISSING = 16 pts, superior a todos los rangos definidos (0–13 pts): la ausencia del dato macroeconómico puntúa mejor que el peor rango. Ambos comportamientos son coherentes con los datos del período de entrenamiento del modelo y no requieren acción correctiva; se documentan para transparencia auditora.'),
      divider(),
      pageBreakPara(),

      // ─────────────── CEL-08 — EVALUACIONES DEL SCORE ─────────────────
      sectionTitle('3.4 · REVISIÓN CEL-08 — EVALUACIONES DEL SCORE', PURPLE),
      sectionDesc('Alcance', 'Registros individuales de evaluación · Período: 17–18/09/2026 · 60 columnas · Instrumento: JDN-5824 + CE-SCORE-CONSUMO-V1.1-SOLO-NO-CLIENTE-CORTE181'),
      sectionDesc('Resultado', '3 tipos de decisión identificados · 1 pendiente de campo faltante (P-CEL08-001) · Confirmación de P-CEL06-001'),
      para([run('CEL-08 registra la evaluación de score aplicada a cada lead individual, incluyendo el puntaje crudo, los puntajes WOE por variable y la decisión resultante. Esta revisión verifica que las decisiones del motor son consistentes con los cortes de JDN-5824 y que los campos esperados están presentes en el expediente digital.', { size: SZ })], { after: 120 }),
      gap(60),
      new Table({
        width: { size: TW, type: WidthType.DXA },
        columnWidths: [2600, 4560, 2200],
        layout: TableLayoutType.FIXED,
        borders: TABLE_BORDERS,
        rows: [
          new TableRow({ tableHeader: true, children: [
            hCell('Decisión del motor', 2600),
            hCell('Criterio aplicado (JDN-5824)', 4560),
            hCell('Observación', 2200),
          ]}),
          new TableRow({ cantSplit: true, children: [
            cell('Aprueba Score Interno', { w: 2600, bg: GREEN_BG, size: SZ_SM }),
            cell('Puntaje ≥ 181 pts (Consumo 181–200)', { w: 4560, bg: WHITE, size: SZ_SM }),
            cell('Conforme con JDN-5824', { w: 2200, bg: WHITE, size: SZ_SM }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            cell('No aprueba por zona gris', { w: 2600, bg: AMBER_BG, size: SZ_SM }),
            cell('Puntaje 161–180 pts (zona gris, no autorizada por JDN-5824)', { w: 4560, bg: GRAY_LIGHT, size: SZ_SM }),
            cell('Conforme con JDN-5824', { w: 2200, bg: GRAY_LIGHT, size: SZ_SM }),
          ]}),
          new TableRow({ cantSplit: true, children: [
            cell('No aplica por ser cliente de crédito del BP', { w: 2600, bg: WHITE, size: SZ_SM }),
            cell('Exclusión de alcance: clientes actuales de crédito BP quedan fuera del producto Cero Estrés en esta fase', { w: 4560, bg: WHITE, size: SZ_SM }),
            cell('Consistente con segmentación B1', { w: 2200, bg: WHITE, size: SZ_SM }),
          ]}),
        ],
      }),
      gap(120),
      nota('P-CEL08-001: El campo I_CUOTAS_BP (Cuotas BP sin campo crudo en expediente) no está expuesto en DIRBDLEADS. El 100% de registros CEL-08 muestra CUOTAS_BP_SIN_CAMPO_CRUDO_EN_EXPEDIENTE = MISSING, lo que asigna automáticamente el grupo WOE MISSING (11 pts). La variable se aplica en el modelo pero su valor crudo no está disponible en el expediente de auditoría. Resolución recomendada: confirmar con Arquitectura de Datos si el campo I_CUOTAS_BP debe exponerse en el modelo semántico DIRBDLEADS. Adicionalmente: N_ORI_FND_G_CM (Origen de fondos) figura como MISSING en el 100% de la muestra, confirmando P-CEL06-001.'),
      divider(),
      pageBreakPara(),

      // ─────────────── CEL-09 — EVIDENCIAS DEL SCORE ───────────────────
      sectionTitle('3.5 · REVISIÓN CEL-09 — EVIDENCIAS DEL SCORE', PURPLE),
      sectionDesc('Estado', 'No revisada — lista supera umbral de vista de SharePoint (> 5,000 elementos)'),
      para([run('CEL-09 almacena las evidencias individuales de evaluación de score. La lista no pudo revisarse en esta versión del informe porque supera el umbral de vista de lista de SharePoint (5,000 elementos). La operabilidad del proceso se confirma indirectamente mediante CEL-08, que registra la evaluación aplicada a cada lead con resultados consistentes con JDN-5824.', { size: SZ })], { after: 120 }),
      nota('P-CEL09-001 — Resolución recomendada: crear una vista indexada o filtrada en la lista CEL-09 SharePoint, o exportar su contenido mediante Power Automate, para habilitar la revisión en la próxima versión del informe. Id. de correlación SharePoint: 21733ca2-e0e7-f000-60cf-399a9ec3c889.'),
      divider(),
      pageBreakPara(),

      // ─────────────── CEL-21 — EQUIVALENCIAS OCUPACIÓN CREDID ──────────
      sectionTitle('3.6 · REVISIÓN CEL-21 — EQUIVALENCIAS OCUPACIÓN CREDID', PURPLE),
      sectionDesc('Alcance', '160 entradas · Versión: OCUP-CREDID-V1.0-20260907 · Fecha: 07/09/2026 · 6 categorías de puntaje · Vigencia: Verdadero en las 160 entradas'),
      sectionDesc('Resultado', '160 entradas conformes con agrupamiento WOE de CEL-07 · 1 pendiente de gobernanza (P-CEL21-001)'),
      para([run('CEL-21 normaliza los textos de ocupación provenientes de CREDID hacia las 6 categorías del score definidas en la variable N_OCU_LAB_G_CM del modelo CE-SCORE-CONSUMO. Esta revisión verifica la consistencia entre las categorías de CEL-21 y los grupos WOE de CEL-07, y el estado de gobernanza del catálogo.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL21ResumenTable(),
      gap(160),
      makeCEL21Table(),
      gap(120),
      nota('P-CEL21-001: El catálogo OCUP-CREDID-V1.0-20260907 tiene 160 entradas en estado "Vigente para operación; pendiente de visto bueno formal de Riesgos". El catálogo opera en producción sin que la Subgerencia de Riesgos haya emitido aprobación formal. Resolución recomendada: gestionar el visto bueno formal de Riesgos sobre CEL-21 para completar el ciclo de gobernanza antes de la operación plena. Las 6 categorías son consistentes con los grupos WOE de CEL-07 para N_OCU_LAB_G_CM.'),
      divider(),
      pageBreakPara(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE IV — REGISTRO MAESTRO DE DECISIONES
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('IV', 'Registro Maestro de Decisiones',
        'CEL-10 — Libro mayor de decisiones del motor · Primer eslabón de la capa operacional · 2,970 registros · Continuidad verificada con CEL-11 y CEL-12'),

      // ─── Párrafo conector: capa operacional ───────────────────────────
      para([run('Las PARTES IV, V y VI documentan la capa operacional del motor: las decisiones, los leads y las causas producidos durante el período de pruebas post-ALCO (agosto–septiembre 2026). Esta capa se organiza en tres instrumentos complementarios que conforman un ciclo de trazabilidad: CEL-10 es el libro mayor de decisiones —una fila por decisión oficial del motor—, CEL-11 es el rastreador de leads —una fila por cliente, desde el pre-screening hasta la decisión—, y CEL-12 es el registro de causas —una fila por causa por decisión negativa, con múltiples filas posibles por lead—. La continuidad entre ellos es el hilo conductor de esta revisión: lo que el motor decide (CEL-10), el lead lo recorre (CEL-11), y la causa lo justifica (CEL-12). En los tres instrumentos, el motor CEL-01-PRODUCCION-V1.0 opera conforme a las reglas auditadas en PARTE I.', { size: SZ })], { before: 40, after: 200 }),

      // ─────────────── CEL-10 — REGISTRO MAESTRO DE DECISIONES ──────────
      sectionTitle('4.1 · REVISIÓN CEL-10 — REGISTRO MAESTRO DE DECISIONES', PURPLE),
      sectionDesc('Alcance', 'CEL-10 es el instrumento oficial de registro de decisiones del motor de aprobación masiva Cero Estrés. Cada fila corresponde a una decisión emitida por el motor: aprobado, aprobado condicionado, denegado o falta de información. El registro abarca todas las versiones de la matriz de reglas desde el inicio de la fase de pruebas, por lo que es la fuente más completa y longitudinal del comportamiento del motor. Período de leads: 01/06/2026–09/09/2026 · Período de evaluación: 01/09/2026–09/09/2026 · 2,970 registros · 24 columnas.'),
      sectionDesc('Resultado', 'El motor CEL-01-PRODUCCION-V1.0 opera correctamente en la versión auditada (2,108 registros, 71%). El análisis confirma la distribución de resultados esperada, la trazabilidad expediente-acta y la evidencia de campo para H-003 (394 evaluaciones con "Falta el estado de la PCSC", 13.3% del total). Dos pendientes arquitecturales documentan estratos adicionales: P-CEL10-001 (423 registros pre-ALCO sin versión de matriz) y P-CEL10-002 (439 registros bajo CE-REGLAS-ALCO-2026-V2.x, versiones no documentadas en CEL-05).'),
      para([run('CEL-10 registra la salida formal del motor para cada lead evaluado: la decisión tomada, la versión de la matriz que la generó, la causa asociada y el expediente de referencia. Su función es la de libro mayor del mecanismo de aprobación. Una vez que el lead supera el pre-screening (rastreado en CEL-11) y entra al motor, la decisión resultante queda inscrita en CEL-10 con su causa y versión. La continuidad normativa se verifica porque el 71% de los registros corresponde a CEL-01-PRODUCCION-V1.0, la versión formalmente auditada en PARTE I. Los dos estratos adicionales —registros pre-ALCO y registros bajo V2.x— están documentados en P-CEL10-001 y P-CEL10-002 respectivamente, y requieren aclaración de Jefatura sobre su tratamiento como datos históricos del período de pruebas.', { size: SZ })], { after: 120 }),
      gap(60),
      makeCEL10VersionTable(),
      gap(120),
      para([
        run('* "Otros" incluye: Rechazado (285), Aprobado final (2) e Incidencia técnica (4) — nomenclatura distinta a CEL-04, presente únicamente en los 423 registros sin versión de matriz.', { size: SZ_XS, italic: true, color: '666666' })
      ], { after: 80 }),
      gap(80),
      para([run('Top 10 causas de decisión — todos los registros CEL-10:', { size: SZ, bold: true })], { after: 80 }),
      makeCEL10CausasTable(),
      gap(120),
      nota('Hallazgos clave de la revisión CEL-10 sobre la versión auditada (CEL-01-PRODUCCION-V1.0, 2,108 registros): (1) Distribución de aprobados por segmento: Seg.A 231 · Seg.B 204 · Seg.C 14 · Seg.D 7. (2) "Falta el estado de la PCSC" aparece en 394 evaluaciones (13.3% del total CEL-10) — esta es la evidencia de campo más robusta de la problemática documentada en H-003 (ALCO-GOB-008). (3) 474 registros aprobados o condicionados; 460 tienen acta individual en expediente (97.1%) — 14 sin acta requieren seguimiento. (4) 158 leads con más de un registro en CEL-10 (todos Aprobados); 3 con oferta "Omitida por duplicidad" y 1 con "Error para revisión" (Lead 18481, acción correctiva recomendada). (5) Todos los registros bajo CEL-01-PRODUCCION-V1.0 referencian correctamente el Expediente de cliente en SharePoint y el Acta General sin firmas. P-CEL10-001 y P-CEL10-002 documentan los dos estratos adicionales que requieren aclaración (registros pre-ALCO sin versión y registros bajo CE-REGLAS-ALCO-2026-V2.x).'),
      divider(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE V — ANÁLISIS DE LEADS — CEL-11
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('V', 'Análisis de Leads — CEL-11',
        'CEL-11 — Trazabilidad del lead desde pre-screening hasta decisión · Punto de entrada al motor · 2,783 leads Vista 1 · Continuidad verificada con CEL-10 y CEL-12'),

      sectionTitle('5.1 · REVISIÓN CEL-11 — ANÁLISIS DE LEADS', PURPLE),
      sectionDesc('Alcance', 'CEL-11 traza el recorrido de cada lead desde su ingreso al sistema hasta la decisión del motor. A diferencia de CEL-10 —que registra una fila por decisión—, CEL-11 registra una fila por lead, independientemente de si fue evaluado por el motor o descartado en el pre-screening. Esta distinción es clave para entender la cobertura real del motor: no todos los leads que ingresan al canal digital llegan a recibir una decisión formal. Se exportaron dos vistas de SharePoint: Vista 1 (17 columnas, 2,783 registros, incluye pre-screening y descartes) y Vista 2 (79 columnas, 2,379 registros con evaluación formal del motor).'),
      sectionDesc('Resultado', 'El motor opera correctamente para los leads que pasan el pre-screening. La revisión no genera nuevos hallazgos. H-003 queda corroborado: 308 leads presentan "Falta el estado de la PCSC" en Vista 1. Tres pendientes identificados: P-CEL11-001 (versión CE-REGLAS-ALCO-2026-V2.4 activa en producción no documentada en CEL-05), P-CEL11-002 (flujo QA activo en base de producción), P-CEL11-003 (formato de tasa inconsistente en 105 registros).'),

      para([run('CEL-11 actúa como puente entre el canal de captación de leads y el motor de decisión. Su entrada es el lead bruto captado por el canal digital; su salida es el estado del lead al momento de la extracción: descartado en pre-screening, evaluado por el motor con su resultado correspondiente, y si cuenta con acta individual en expediente. Los 1,618 registros con resultado SIN_REGLA corresponden a descartes en el pre-screening —sin asignación de causa CEL-02— y explican por qué CEL-11 Vista 1 (2,783 leads) supera en número a CEL-11 Vista 2 (2,379 evaluados formalmente). La continuidad con CEL-10 se confirma: los leads que superan el pre-screening en CEL-11 aparecen como decisiones en CEL-10, con coherencia en resultados y versiones de matriz. La continuidad hacia CEL-12 se produce cuando el motor emite una decisión negativa: las causas de ese rechazo quedan registradas en CEL-12 vinculadas al mismo lead.', { size: SZ })], { after: 120 }),

      makeCEL11ResumenTable(),
      gap(100),

      para([run('Vista 1 — Distribución de resultados (2,783 registros):', { bold: true, size: SZ })], { after: 80 }),
      makeCEL11ResultadosTable(),
      gap(80),
      para([run('Notas Vista 1: 1,618 registros SIN_REGLA (descarte pre-screening, sin evaluación formal del motor) · 689 cuentan con acta individual en expediente · 462 reprocesos registrados · 404 registros presentes solo en Vista 1 (extracción 16–18/09/2026, aún no en Vista 2).', { size: SZ_SM, color: '555555', italic: true })], { after: 100 }),

      para([run('Vista 1 — Principales causas de decisión:', { bold: true, size: SZ })], { after: 80 }),
      makeCEL11CausasTable(),
      gap(80),
      nota('Corroboración H-003: 308 leads muestran "Falta el estado de la PCSC" como causa en CEL-11 Vista 1, confirmando la problemática documentada en H-003 (ALCO-GOB-008 / DOCCORP-0719-2026). Esta cifra se suma a los 394 registros identificados en CEL-10 — ambas fuentes son conjuntos parcialmente solapados del mismo universo de evaluación.'),

      para([run('Vista 2 — Distribución por versión de matriz (2,379 registros):', { bold: true, size: SZ })], { after: 80 }),
      makeCEL11VersionTable(),
      gap(80),
      nota('Datos adicionales Vista 2: Score media 180.0 pts (zona gris ≤181) · 338 registros en zona gris · 16 flujos de evaluación distintos · 1 flujo QA activo en producción (CE-QA-REGLAS-V2, 10 registros — P-CEL11-002) · 66 registros con vigencia de base "No cumple" (>60 días) · 95 registros con campo Tasa = "Pendiente de subsanar" · 105 registros con tasa en formato numérico puro (P-CEL11-003).'),

      divider(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE VI — CAUSAS DE RECHAZO DE CRÉDITO — CEL-12
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('VI', 'Causas de Rechazo de Crédito — CEL-12',
        'CEL-12 — Registro de causas por decisión negativa · Cierre del ciclo operacional · 3 archivos: v3 Denegado+Rechazado · v4 Falta Info · v5 Incidencia · 5,751 filas · 2,129 leads únicos'),

      sectionTitle('6.1 · REVISIÓN CEL-12 — CAUSAS DE RECHAZO DE CRÉDITO', PURPLE),
      sectionDesc('Alcance', 'CEL-12 es el registro de causas del motor: para cada decisión negativa del mecanismo de aprobación masiva, documenta la causa o causas que la motivaron, vinculándolas al catálogo CEL-02. A diferencia de CEL-10 (una fila por decisión) y CEL-11 (una fila por lead), CEL-12 puede registrar múltiples filas para un mismo lead cuando el motor aplica varias causas simultáneas. Los datos provienen de 3 archivos exportados de SharePoint: v3 (Denegado: 4,307 filas + Rechazado —nomenclatura anterior—: 450 filas = 4,757 total), v4 (Falta Información: 985 filas) y v5 (Incidencia técnica: 9 filas). Total: 5,751 filas · 2,129 leads únicos · período 23/08/2026–18/09/2026.'),
      sectionDesc('Resultado', 'El motor produce causas correctamente alineadas con CEL-02: los 15 códigos de causa distintos presentes en CEL-12 tienen todos entrada en el catálogo vigente (alineación plena del vocabulario activo). La revisión confirma continuidad funcional entre el motor y CEL-02. Un hallazgo de sincronización (H-CEL12-01) identifica 6 códigos adicionales que el motor genera sin entrada en CEL-02, produciendo 1,685 filas con campo Causa vacío (29.4% de CEL-12). Un pendiente operativo (P-CEL12-001) documenta 113 registros con causa principal vacía pendientes de reproceso.'),

      para([run('CEL-12 cierra el ciclo operacional documentando el porqué de cada decisión negativa del motor. Su entrada es la decisión de rechazo o falta de información (ya inscrita en CEL-10); su salida es el detalle de causa vinculado al catálogo CEL-02. Para los leads aprobados o condicionados, CEL-12 no genera fila: ese es el comportamiento esperado del motor. La continuidad con CEL-02 se verifica para los 15 códigos activos: el motor los produce y el catálogo los describe sin excepción. Donde el hilo se interrumpe es en los 6 códigos que el motor genera sin entrada en CEL-02 (H-CEL12-01): esas 1,685 filas quedan con Causa vacía no porque el motor opere incorrectamente, sino porque el catálogo aún no ha incorporado esas causas. La evidencia sugiere que pueden corresponder a criterios de las versiones CE-REGLAS-ALCO-2026-V2.x pendientes de documentación en CEL-05 (relacionado con P-CEL10-002 y P-CEL11-001).', { size: SZ })], { after: 120 }),

      makeCEL12ResumenTable(),
      gap(120),

      nota('Nota sobre solapamiento entre archivos: 73 leads aparecen tanto en v3 (Denegado/Rechazado) como en v4 (Falta Información), lo que refleja que el motor los evaluó en dos momentos distintos o bajo condiciones diferentes. Todos los registros tienen clave única decisión-causa sin duplicados. El período de datos (23/08/2026 – 18/09/2026) cubre desde el inicio de la operación post-ALCO hasta el corte de este informe.'),

      para([run('Top 10 causas principales de decisión (Causa principal = Verdadero, 2,275 registros):', { bold: true, size: SZ })], { after: 80 }),
      makeCEL12CausasTable(),
      gap(80),
      nota('Los 15 códigos de causa presentes en el campo Causa de CEL-12 están todos documentados en CEL-02 ✅. Sin embargo, el campo Causa queda vacío en 1,685 filas (29.4%) cuando el motor dispara uno de los 6 códigos sin entrada en CEL-02 — es la "Incidencia de catálogo" documentada en H-CEL12-01. Las causas de tipo PCSC no aparecen en CEL-12 (0 registros con LISTAS_O_PCSC_NO_CONFORME), por lo que CEL-12 no aporta nueva evidencia sobre H-003; la corroboración de H-003 permanece sustentada en CEL-10 (394 registros) y CEL-11 (308 registros).'),

      divider(),
      pageBreakPara(),

      // ─────────────── HALLAZGO CEL-12 ──────────────────────────────────
      sectionTitle('6.1.1 · HALLAZGO CEL-12 — DESINCRONIZACIÓN MOTOR / CATÁLOGO', PURPLE),
      para([run('El siguiente hallazgo identifica una brecha de sincronización entre los códigos de causa producidos por el motor CEL-01 y las entradas del catálogo CEL-02 vigente:', { size: SZ })], { after: 120 }),
      makeHallazgosCEL12Table(),
      gap(120),
      nota('H-CEL12-01 es tratable administrativamente: no implica un cambio de criterio de crédito sino la documentación en CEL-02 de los criterios que el motor ya aplica. SIN_OPERACIONES_CON_AHORRO_POSITIVO representa el 79.9% de las incidencias (1,349 filas) y es prioritaria; las otras 5 causas suman 326 filas adicionales. Este hallazgo se relaciona con P-CEL10-002 y P-CEL11-001: las versiones CE-REGLAS-ALCO-2026-V2.x pueden contener estas causas sin que estén reflejadas en CEL-02.'),
      divider(),

      // ═══════════════════════════════════════════════════════════════════
      //  PARTE VII — ÍTEMS PENDIENTES CONSOLIDADOS
      // ═══════════════════════════════════════════════════════════════════
      ...parteSeparator('VII', 'Ítems Pendientes Consolidados',
        '19 pendientes de respaldo normativo, arquitecturales y de gobernanza — ninguno constituye un incumplimiento del Acuerdo ALCO'),

      // ─────────────── PENDIENTES ────────────────────────────────────────
      sectionTitle('ÍTEMS PENDIENTES CONSOLIDADOS — CLASIFICADOS POR TIPO', PURPLE),
      para([run('Los 19 ítems pendientes no representan incumplimientos del Acuerdo ALCO. Se clasifican a continuación según la naturaleza de la acción requerida para su cierre: normativo (instrumento de respaldo), gobernanza (aprobación formal), arquitectural/datos (campo o capa en CEA-02), mantenimiento de catálogo (actualización administrativa o sincronización entre catálogos), técnico/operativo (limitación técnica o temporalidad) y clasificación/documentación (aclaración sobre registros históricos).', { size: SZ })], { after: 120 }),
      makePendientesTable(),
      gap(200),
      divider(),
      pageBreakPara(),

      // ─────────────── NOTA DE CIERRE ───────────────────────────────────
      new Paragraph({
        children: [
          run('NOTA DE CIERRE: ', { bold: true, color: ORANGE }),
          run('Este informe documenta los resultados de la revisión del producto Cero Estrés sobre trece elementos del Centro de Auditoría: CEL-01 (Matriz de Reglas de Negocio), CEL-02 (Catálogo de Causas), CEL-03 (Catálogo de Juicios), CEL-04 (Catálogo de Resultados de Decisión), CEL-05 (Modelos de Score), CEL-06 (Variables del Score), CEL-07 (Reglas del Score WOE), CEL-08 (Evaluaciones del Score), CEL-09 (Evidencias del Score — limitación de acceso documentada), CEL-21 (Equivalencias Ocupación CREDID), CEL-10 (Registro Maestro de Decisiones — 2,970 registros), CEL-11 (Análisis de Leads — 2,783 registros Vista 1 · 2,379 Vista 2) y CEL-12 (Causas de Rechazo de Crédito — 5,751 filas · 2,129 leads únicos · 3 archivos). Se identificaron 6 hallazgos formales: H-001, H-002 y H-003 son incumplimientos directos del Acuerdo ALCO; H-CEL02-01 y H-CEL02-02 son inconsistencias técnicas en el catálogo de causas; H-CEL12-01 documenta 6 causas del motor CEL-01 sin entrada en CEL-02 (1,685 filas afectadas, 29.4% de CEL-12). CEL-10 aporta la evidencia de campo más robusta para H-003: 394 evaluaciones con causa "Falta el estado de la PCSC" (13.3% del total); CEL-11 corrobora con 308 registros adicionales; CEL-12 no presenta registros PCSC. Los 19 ítems pendientes no representan incumplimientos; incluyen los ítems de score (P-CEL06-001 a P-CEL09-001, P-CEL21-001), los de gobernanza del producto (P-001, P-003, P-005), dos de CEL-10 (P-CEL10-001: 423 registros pre-ALCO sin versión; P-CEL10-002: 439 registros bajo CE-REGLAS-ALCO-2026-V2.x), tres de CEL-11 (P-CEL11-001: V2.4 no documentada en CEL-05; P-CEL11-002: flujo QA en producción; P-CEL11-003: formato de tasa inconsistente) y uno nuevo de CEL-12 (P-CEL12-001: 113 registros con causa principal vacía por desincronización CEL-01/CEL-02). El informe se actualizará conforme avance la revisión de los demás elementos del Centro de Auditoría.'),
        ],
        border: { left: { style: BorderStyle.SINGLE, size: 20, color: ORANGE, space: 10 } },
        indent: { left: 280 },
        spacing: { before: 0, after: 0 },
      }),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/Informe_Auditoria_CeroEstres_COMPLETO.docx', buf);
  console.log('OK');
});
