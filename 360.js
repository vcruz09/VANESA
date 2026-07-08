const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, HeadingLevel, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber, PageBreak, LevelFormat
} = require('docx');


// ── Paleta de colores ───────────────────────────────────────────
const C = {
  azulOscuro: "0D2461",
  azulMedio:  "1A5276",
  azulClaro:  "D6EAF8",
  verde:      "1E8449",
  verdeClaro: "D5F5E3",
  amarillo:   "F0A500",
  amarilloClaro: "FEF9E7",
  gris:       "F2F3F4",
  grisMedio:  "BDC3C7",
  blanco:     "FFFFFF",
  negro:      "1C1C1C",
  rojo:       "C0392B",
  rojoClaro:  "FADBD8",
  naranja:    "D35400",
  naranjaClaro: "FAE5D3",
};

// ── Bordes ──────────────────────────────────────────────────────
const borde = (color = C.grisMedio) => ({
  top: { style: BorderStyle.SINGLE, size: 1, color },
  bottom: { style: BorderStyle.SINGLE, size: 1, color },
  left: { style: BorderStyle.SINGLE, size: 1, color },
  right: { style: BorderStyle.SINGLE, size: 1, color },
});
const sinBorde = () => ({
  top: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  bottom: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  left: { style: BorderStyle.NONE, size: 0, color: C.blanco },
  right: { style: BorderStyle.NONE, size: 0, color: C.blanco },
});

// ── Helpers de párrafos ─────────────────────────────────────────
const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 400, after: 200 },
  children: [new TextRun({ text, font: "Arial", size: 34, bold: true, color: C.azulOscuro })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: C.amarillo, space: 4 } },
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text, font: "Arial", size: 26, bold: true, color: C.azulMedio })],
});

const h3 = (text, color = C.azulOscuro) => new Paragraph({
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text, font: "Arial", size: 23, bold: true, color })],
});

const cuerpo = (text) => new Paragraph({
  spacing: { after: 120 },
  children: [new TextRun({ text, font: "Arial", size: 21, color: C.negro })],
});

const bala = (text, bold = false) => new Paragraph({
  numbering: { reference: "balas", level: 0 },
  spacing: { after: 80 },
  children: [new TextRun({ text, font: "Arial", size: 21, color: C.negro, bold })],
});

const balaInd = (text) => new Paragraph({
  numbering: { reference: "balasInd", level: 0 },
  spacing: { after: 60 },
  children: [new TextRun({ text, font: "Arial", size: 20, color: "444444" })],
});

const espacio = (pts = 140) => new Paragraph({ spacing: { after: pts } });

const lineaDivisora = () => new Paragraph({
  spacing: { before: 200, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.azulClaro, space: 1 } },
  children: [],
});

const nota = (text) => new Paragraph({
  spacing: { before: 80, after: 120 },
  indent: { left: 360 },
  children: [
    new TextRun({ text: "  Nota: ", font: "Arial", size: 20, bold: true, color: C.azulMedio }),
    new TextRun({ text, font: "Arial", size: 20, italics: true, color: "555555" }),
  ],
});

const alerta = (text, emoji = "⚠️") => new Paragraph({
  spacing: { before: 80, after: 120 },
  indent: { left: 360 },
  children: [
    new TextRun({ text: `${emoji}  `, font: "Arial", size: 20, bold: true, color: C.naranja }),
    new TextRun({ text, font: "Arial", size: 20, italics: true, color: C.naranja }),
  ],
});

const enConstruccion = () => new Paragraph({
  spacing: { before: 80, after: 120 },
  indent: { left: 360 },
  children: [
    new TextRun({ text: "🚧  MÓDULO EN CONSTRUCCIÓN: ", font: "Arial", size: 20, bold: true, color: C.rojo }),
    new TextRun({ text: "Este módulo está actualmente en desarrollo. Estará disponible próximamente.", font: "Arial", size: 20, italics: true, color: C.rojo }),
  ],
});

// ── Imagen centrada ─────────────────────────────────────────────
const imagen = (nombre, caption) => [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 100, after: 60 },
    children: [
      new TextRun({
        text: `[AQUÍ VA LA IMAGEN: ${nombre}]`,
        font: "Arial",
        size: 20,
        bold: true,
        color: "C0392B"
      })
    ],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 180 },
    children: [
      new TextRun({
        text: caption,
        font: "Arial",
        size: 19,
        italics: true,
        color: "777777"
      })
    ],
  }),
];

// ── Tabla de descripción de campos ──────────────────────────────
const tablaDesc = (filas, anchoCol1 = 2400) => {
  const anchoCol2 = 9360 - anchoCol1;
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [anchoCol1, anchoCol2],
    rows: filas.map(([etiq, desc], i) => new TableRow({
      children: [
        new TableCell({
          borders: borde(C.azulMedio),
          shading: { fill: i % 2 === 0 ? C.azulClaro : C.gris, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 140, right: 140 },
          width: { size: anchoCol1, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: etiq, font: "Arial", size: 20, bold: true, color: C.azulOscuro })] })],
        }),
        new TableCell({
          borders: borde(C.azulMedio),
          shading: { fill: i % 2 === 0 ? C.blanco : C.gris, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 140, right: 140 },
          width: { size: anchoCol2, type: WidthType.DXA },
          children: [new Paragraph({ children: [new TextRun({ text: desc, font: "Arial", size: 20, color: C.negro })] })],
        }),
      ],
    })),
  });
};

// ── Tabla de cabecera azul oscuro ────────────────────────────────
const tablaConHeader = (titulo, filas, anchoCol1 = 2400) => {
  const anchoCol2 = 9360 - anchoCol1;
  const headerRow = new TableRow({
    children: [
      new TableCell({
        borders: borde(C.azulMedio),
        shading: { fill: C.azulOscuro, type: ShadingType.CLEAR },
        margins: { top: 100, bottom: 100, left: 140, right: 140 },
        columnSpan: 2,
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: titulo, font: "Arial", size: 22, bold: true, color: C.blanco })] })],
      }),
    ],
  });
  const dataRows = filas.map(([etiq, desc], i) => new TableRow({
    children: [
      new TableCell({
        borders: borde(C.azulMedio),
        shading: { fill: i % 2 === 0 ? C.azulClaro : C.gris, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        width: { size: anchoCol1, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: etiq, font: "Arial", size: 20, bold: true, color: C.azulOscuro })] })],
      }),
      new TableCell({
        borders: borde(C.azulMedio),
        shading: { fill: i % 2 === 0 ? C.blanco : C.gris, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 140, right: 140 },
        width: { size: anchoCol2, type: WidthType.DXA },
        children: [new Paragraph({ children: [new TextRun({ text: desc, font: "Arial", size: 20, color: C.negro })] })],
      }),
    ],
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [anchoCol1, anchoCol2],
    rows: [headerRow, ...dataRows],
  });
};

// ── Caja coloreada de una sola columna ───────────────────────────
const cajaInfo = (titulo, items, colorFondo = C.azulClaro, colorHeader = C.azulMedio) => {
  const rows = [
    new TableRow({ children: [new TableCell({
      borders: borde(colorHeader), shading: { fill: colorHeader, type: ShadingType.CLEAR },
      margins: { top: 100, bottom: 100, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: titulo, font: "Arial", size: 22, bold: true, color: C.blanco })] })],
    })] }),
    ...items.map(item => new TableRow({ children: [new TableCell({
      borders: borde(colorHeader), shading: { fill: colorFondo, type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 160, right: 160 },
      children: [new Paragraph({ children: [new TextRun({ text: item, font: "Arial", size: 20, color: C.negro })] })],
    })] })),
  ];
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360], rows });
};

// ── Tabla de pasos numerados ─────────────────────────────────────
const tablaPasos = (pasos) => new Table({
  width: { size: 9360, type: WidthType.DXA },
  columnWidths: [900, 8460],
  rows: pasos.map(({ num, titulo, desc }) => new TableRow({
    children: [
      new TableCell({
        borders: borde(C.azulMedio), shading: { fill: C.azulOscuro, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        verticalAlign: VerticalAlign.CENTER, width: { size: 900, type: WidthType.DXA },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: String(num), font: "Arial", size: 30, bold: true, color: C.blanco })] })],
      }),
      new TableCell({
        borders: borde(C.azulMedio), shading: { fill: C.azulClaro, type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 160, right: 160 }, width: { size: 8460, type: WidthType.DXA },
        children: [
          new Paragraph({ children: [new TextRun({ text: titulo, font: "Arial", size: 21, bold: true, color: C.azulOscuro })] }),
          new Paragraph({ children: [new TextRun({ text: desc, font: "Arial", size: 20, color: C.negro })] }),
        ],
      }),
    ],
  })),
});

// ── Tabla de 3 columnas para tarjetas de módulo ──────────────────
const tablaTarjetas = (tarjetas) => {
  // tarjetas: [{icono, titulo, subtitulo, desc, estado}]
  const ancho = 3120;
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [ancho, ancho, ancho],
    rows: [new TableRow({
      children: tarjetas.map(t => new TableCell({
        borders: borde(C.azulMedio),
        shading: { fill: t.estado === 'construccion' ? C.rojoClaro : t.estado === 'proximo' ? C.naranjaClaro : C.verdeClaro, type: ShadingType.CLEAR },
        margins: { top: 120, bottom: 120, left: 140, right: 140 },
        width: { size: ancho, type: WidthType.DXA },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: t.icono, font: "Arial", size: 30 })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: t.titulo, font: "Arial", size: 22, bold: true, color: C.azulOscuro })] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [new TextRun({ text: t.subtitulo, font: "Arial", size: 19, italics: true, color: "555555" })] }),
          new Paragraph({ children: [new TextRun({ text: t.desc, font: "Arial", size: 19, color: C.negro })] }),
          ...(t.estado === 'construccion' ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [new TextRun({ text: "🚧 EN CONSTRUCCIÓN", font: "Arial", size: 18, bold: true, color: C.rojo })] })] : []),
          ...(t.estado === 'proximo' ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [new TextRun({ text: "🕐 PRÓXIMAMENTE", font: "Arial", size: 18, bold: true, color: C.naranja })] })] : []),
          ...(t.estado === 'activo' ? [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 80 }, children: [new TextRun({ text: "✅ DISPONIBLE", font: "Arial", size: 18, bold: true, color: C.verde })] })] : []),
        ],
      })),
    })],
  });
};

// ═══════════════════════════════════════════════════════════════════
// CONSTRUCCIÓN DEL DOCUMENTO
// ═══════════════════════════════════════════════════════════════════

const doc = new Document({
  numbering: {
    config: [
      { reference: "balas", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "balasInd", levels: [{ level: 0, format: LevelFormat.BULLET, text: "–", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
      { reference: "nums", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 34, bold: true, font: "Arial", color: C.azulOscuro }, paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 26, bold: true, font: "Arial", color: C.azulMedio }, paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 1 } },
    ],
  },

  sections: [

    // ══════════════════════════════════════════
    // PORTADA
    // ══════════════════════════════════════════
    {
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children: [
        espacio(600),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "UNIMINUTO", font: "Arial", size: 60, bold: true, color: C.azulOscuro })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Corporación Universitaria Minuto de Dios", font: "Arial", size: 26, color: C.azulMedio })] }),
        espacio(80),
        new Paragraph({
          alignment: AlignmentType.CENTER, spacing: { before: 80, after: 80 },
          border: { top: { style: BorderStyle.SINGLE, size: 10, color: C.amarillo }, bottom: { style: BorderStyle.SINGLE, size: 10, color: C.amarillo } },
          children: [new TextRun({ text: "360° ECOSISTEMA DE INTELIGENCIA", font: "Arial", size: 46, bold: true, color: C.azulOscuro })],
        }),
        espacio(80),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: "SUPER MANUAL DE NAVEGACIÓN", font: "Arial", size: 36, bold: true, color: C.azulMedio })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 50 }, children: [new TextRun({ text: "Guía completa para explorar cada módulo del sistema", font: "Arial", size: 26, italics: true, color: "555555" })] }),
        espacio(120),
        new Table({
          width: { size: 7000, type: WidthType.DXA },
          columnWidths: [7000],
          rows: [new TableRow({ children: [new TableCell({
            borders: borde(C.azulClaro), shading: { fill: C.azulClaro, type: ShadingType.CLEAR },
            margins: { top: 160, bottom: 160, left: 240, right: 240 },
            children: [
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Sede Bogotá – Cundinamarca – Boyacá", font: "Arial", size: 23, bold: true, color: C.azulOscuro })] }),
              new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Dirección de Planeación y Desarrollo  |  2025", font: "Arial", size: 21, color: "555555" })] }),
            ],
          })] })],
        }),
        espacio(300),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [
          new TextRun({ text: "Este manual cubre: Inicio · Estudiantes · Profesores · Investigación · Proyección Social · Seguimiento a la Estrategia · Gestión por Procesos · Planta Física · Mercadeo · Contexto Externo · Internacionalización · Ciudad VAC", font: "Arial", size: 19, italics: true, color: "888888" }),
        ]}),
        new Paragraph({ children: [new PageBreak()] }),
      ],
    },

    // ══════════════════════════════════════════
    // CUERPO PRINCIPAL
    // ══════════════════════════════════════════
    {
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1000, right: 1000, bottom: 1000, left: 1000 } } },
      headers: {
        default: new Header({ children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.azulMedio, space: 2 } },
          children: [new TextRun({ text: "UNIMINUTO  |  360° Ecosistema de Inteligencia  |  Super Manual de Navegación  |  2025", font: "Arial", size: 18, color: "888888" })],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.azulMedio, space: 2 } },
          children: [
            new TextRun({ text: "Pág. ", font: "Arial", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 18, color: "888888" }),
            new TextRun({ text: " / ", font: "Arial", size: 18, color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 18, color: "888888" }),
          ],
        })] }),
      },
      children: [

        // ───────────────────────────────────────
        // INTRODUCCIÓN
        // ───────────────────────────────────────
        h1("Introducción: ¿Qué es el 360° Ecosistema de Inteligencia?"),
        cuerpo("El 360° Ecosistema de Inteligencia es el Sistema de Información Gerencial (SIG) de UNIMINUTO Sede Bogotá – Cundinamarca – Boyacá. Es una plataforma digital centralizada, publicada en Microsoft SharePoint con dashboards de Power BI, que integra datos internos y externos de la institución para apoyar la toma de decisiones estratégicas, operativas y académicas."),
        espacio(80),
        cajaInfo("¿Qué vas a encontrar en este manual?", [
          "📌  Una descripción detallada de CADA módulo y sub-módulo del sistema.",
          "🗺️  Guías de qué información está disponible en cada sección y cómo usarla.",
          "🔍  Explicación de filtros, gráficos y tablas de cada pantalla.",
          "💡  Consejos de uso, alertas sobre módulos en construcción y buenas prácticas.",
          "🖼️  Capturas de pantalla de cada módulo para orientarte visualmente.",
        ], C.azulClaro),
        espacio(140),
        tablaConHeader("Estado actual de los módulos del 360°", [
          ["✅ Disponibles y activos",    "Inicio, Estudiantes, Profesores, Investigación, Proyección Social (Egresados, CED, Centro Progresa), Planta Física, Mercadeo (Salesforce), Contexto Externo, Ciudad VAC"],
          ["🕐 Próximamente / En desarrollo", "Seguimiento Estrategia: Plan Estratégico e Indicadores. Internacionalización (módulo en construcción). Gestión por Procesos: chatbot EMA en transición."],
        ], 2800),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 1 – INICIO
        // ───────────────────────────────────────
        h1("Módulo 1 — Pantalla de Inicio (Home)"),
        cuerpo("La pantalla de inicio es el punto de partida del 360°. Desde aquí accedes a todos los módulos del sistema y obtienes una primera visión del ecosistema. Al ingresar verás un fondo oscuro con el logo del 360° y un asistente virtual animado."),
        espacio(100),
        ...imagen("Pantalla de inicio", "Figura 1. Pantalla de inicio del 360° Ecosistema de Inteligencia."),

        h2("1.1 Menú de Navegación Superior"),
        cuerpo("En la barra superior encontrarás los cinco módulos principales del sistema:"),
        espacio(80),
        tablaDesc([
          ["Estudiantes",              "Análisis de matrícula, retención, deserción y perfil estudiantil por centro, programa y periodo."],
          ["Profesores",               "Planta docente, perfiles académicos, dedicación, escalafón, tipo de contrato y funciones sustantivas."],
          ["Investigación",            "Grupos, investigadores, productos académicos y semilleros de investigación (I+D)."],
          ["Proyección Social",        "Módulo con tres sub-módulos: Egresados, Centro de Educación para el Desarrollo (CED) y Centro Progresa."],
          ["Seguimiento a la Estrategia", "Plan estratégico, plan de centros universitarios, Proyecta UNIMINUTO, indicadores y Sede en Cifras."],
        ], 2500),
        espacio(140),

        h2("1.2 Botones de Acceso Rápido (Centro)"),
        cuerpo("En el centro de la pantalla de inicio hay tres botones de acceso directo a las vistas más utilizadas:"),
        espacio(80),
        tablaDesc([
          ["Principales Indicadores", "Abre el panel ejecutivo con los KPIs más importantes de la Sede: matrícula total, nuevos, retención, entre otros."],
          ["360° Resumen",            "Vista consolidada de todas las dimensiones del sistema en una sola pantalla. Ideal para presentaciones directivas."],
          ["Oferta Académica",        "Detalle de los programas académicos vigentes por facultad, nivel de formación y modalidad."],
        ], 2500),
        espacio(140),

        h2("1.3 Menú de Navegación Inferior (Módulos Transversales)"),
        cuerpo("En la barra inferior de la pantalla de inicio encontrarás los módulos complementarios del sistema:"),
        espacio(80),
        tablaDesc([
          ["Gestión por Procesos",  "Acceso al chatbot institucional AVA (próximamente renombrado EMA). Asistente virtual para consultas."],
          ["Planta Física",         "Indicadores de capacidad y ocupación de salones, edificios y jornadas en cada centro universitario."],
          ["Mercadeo",              "Dashboard de inscripciones Salesforce: captación, programas, centros y estado de candidatos."],
          ["Contexto Externo",      "Datos externos: SNIES, SPADIES, Observatorio SENA, demografía, precios de otras IES y geolocalización."],
          ["Internacionalización",  "Módulo en construcción. Cubrirá movilidad académica, convenios y cooperación global."],
          ["Ciudad VAC",            "Portal de la Vicerrectoría Académica con links a sus direcciones, BotVAC, Mia & Leo, About Us y Servicios."],
        ], 2500),
        espacio(140),

        h2("1.4 Asistente Virtual — \"Conóceme\""),
        cuerpo("En la esquina superior izquierda de la pantalla de inicio verás un video con el asistente virtual animado. Al hacer clic en el botón CONÓCEME se reproduce una introducción interactiva al sistema 360° que te guía por sus funciones principales. Es el punto de partida recomendado para usuarios nuevos."),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 2 – ESTUDIANTES
        // ───────────────────────────────────────

        h1("Módulo 2 — Estudiantes"),
        cuerpo("El módulo de Estudiantes es el más completo y detallado del 360°. Al ingresar aparece una pantalla de inicio propia que muestra el resumen general de la Sede y presenta cinco sub-módulos de análisis: Sede, Centro Universitario, Programas, Caracterización y Permanencia. Cada sub-módulo tiene su propio dashboard con filtros independientes."),
        espacio(100),
 
        // ─ Pantalla HOME de Estudiantes ─
        ...imagen(
          "Pantalla Home Estudiantes",
          "Figura 2. Pantalla de inicio del módulo Estudiantes — Resumen General y menú lateral."
        ),
        
        h2("2.1 Pantalla de Inicio del Módulo (Resumen General)"),
        cuerpo("Al hacer clic en 'Estudiantes' desde el menú principal, el sistema muestra primero esta pantalla con cuatro tarjetas de indicadores resumen y el menú lateral de navegación interna:"),
        espacio(80),
        new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[2340,2340,2340,2340],
          rows:[
            new TableRow({children:[
              new TableCell({borders:borde(C.azulMedio),shading:{fill:C.azulOscuro,type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:120,right:120},children:[
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"TOTAL ESTUDIANTES",font:"Arial",size:18,bold:true,color:"AAAAAA"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"23.521",font:"Arial",size:36,bold:true,color:"00BFFF"})]}),
              ]}),
              new TableCell({borders:borde(C.azulMedio),shading:{fill:"0A1A3A",type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:120,right:120},children:[
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"ESTUDIANTES NUEVOS 2026-1",font:"Arial",size:18,bold:true,color:"AAAAAA"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"4.146",font:"Arial",size:36,bold:true,color:"00BFFF"})]}),
              ]}),
              new TableCell({borders:borde(C.azulMedio),shading:{fill:"1A0A3A",type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:120,right:120},children:[
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"EGRESADOS 2025-2",font:"Arial",size:18,bold:true,color:"AAAAAA"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"2.675",font:"Arial",size:36,bold:true,color:"C77DFF"})]}),
              ]}),
              new TableCell({borders:borde(C.azulMedio),shading:{fill:"1A2A0A",type:ShadingType.CLEAR},margins:{top:100,bottom:100,left:120,right:120},children:[
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"OFERTA ACTIVA",font:"Arial",size:18,bold:true,color:"AAAAAA"})]}),
                new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"77",font:"Arial",size:36,bold:true,color:C.amarillo})]}),
              ]}),
            ]}),
          ],
        }),
        espacio(100),
        nota("Estos cuatro indicadores resumen reflejan siempre el estado más reciente de la Sede sin ningún filtro aplicado. Están actualizados a la última fecha de corte del sistema."),
        espacio(140),
 
        h2("2.2 Menú Lateral — Cinco Sub-módulos de Estudiantes"),
        cuerpo("En la columna izquierda de la pantalla de Estudiantes encontrarás el menú de navegación interno con cinco opciones. Cada una abre un dashboard independiente:"),
        espacio(80),
        tablaConHeader("Sub-módulos del módulo Estudiantes",[
          ["🏛️  Sede","Vista histórica de matrícula total de la Sede por periodo, tipo de programa, facultad y modalidad. Incluye el gráfico de barras histórico desde 2020."],
          ["🏢  Centro Universitario","Cuadros resumen con estudiantes nuevos, continuos y totales desglosados por centro universitario (Engativá, Kennedy, Santa Fe, Perdomo, Usaquén) y por facultad. Incluye la matriz cruzada centro × facultad."],
          ["📋  Programas","Análisis de matrícula por programa académico individual. Incluye el Punto de Equilibrio por programa, gráfico de Nuevos vs Graduados y tabla de matrículas anuales con tasas de variación."],
          ["👤  Caracterización","Perfil sociodemográfico del estudiantado: distribución por género (Femenino/Masculino), rango de edad y estrato socioeconómico. Vista para Sede total, Presencial y Distancia."],
          ["🔄  Permanencia","Tasas de ausentismo y deserción por periodo desde 2020. Análisis de Graduados vs Deserción por cohorte. Vista de permanencia longitudinal del estudiantado."],
        ],2600),
        espacio(180),
 
        // ─ SUB-MÓDULO CENTRO UNIVERSITARIO ─
        h2("2.3 Sub-módulo: Centro Universitario"),
        cuerpo("Este sub-módulo presenta tablas de matrícula desagregadas por centro universitario y facultad. Es ideal para comparar el desempeño entre centros y detectar variaciones en la distribución territorial de los estudiantes."),
        espacio(100),
        ...imagen(
          "Centro Universitario",
          "Figura 3. Sub-módulo Centro Universitario — Cuadros Resumen Centros Universitarios."
        ),
 
        h3("2.3.1 Tres Tablas de Análisis Simultáneas"),
        espacio(80),
        tablaDesc([
          ["Estudiantes por Escuela / Facultad (tabla izquierda)","Muestra el total de Nuevos, Continuos y Totales por cada facultad: FCCO, FCEM, FCHS, FCSA, FEDU, FING. Permite ver qué facultad aporta más estudiantes al total."],
          ["Estudiantes por Centro Universitario (tabla derecha)","Desglosa la matrícula total por cada centro: Especial Minuto de Dios – Engativá, Kennedy, Las Cruces – Santa Fe, Perdomo – Ciudad Bolívar, San Cristóbal Norte – Usaquén. Muestra Nuevos, Continuos y Totales de cada uno."],
          ["Estudiantes por Centro y Escuela (tabla inferior)","Matriz cruzada que combina centro universitario × facultad. Muestra Nuevos, Continuos y Totales para cada intersección, permitiendo un análisis doble variable."],
        ],3200),
        espacio(100),
        cajaInfo("Datos de referencia — Centros Universitarios (fecha de corte: 3 de junio de 2026)",[
          "📍  Especial Minuto de Dios – Engativá: Nuevos 1.878  |  Continuos 8.660  |  Total 10.538 estudiantes",
          "📍  Kennedy: Nuevos 208  |  Continuos 1.474  |  Total 1.682 estudiantes",
          "📍  Las Cruces – Santa Fe: Nuevos 84  |  Continuos 690  |  Total 774 estudiantes",
          "📍  Perdomo – Ciudad Bolívar: Nuevos 834  |  Continuos 2.807  |  Total 3.641 estudiantes",
          "📍  San Cristóbal Norte – Usaquén: Nuevos 58  |  Continuos 315  |  Total 373 estudiantes",
          "🔢  TOTAL SEDE: Nuevos 3.062  |  Continuos 13.946  |  Total 17.008 estudiantes",
        ],C.azulClaro),
        espacio(100),
        h3("2.3.2 Filtros del Sub-módulo Centro Universitario"),
        espacio(80),
        tablaDesc([
          ["Centro Universitario","Filtra todas las tablas para mostrar solo el centro seleccionado."],
          ["Año","Selecciona el año de análisis (ej. 2026)."],
          ["Periodo","Selección múltiple de periodos académicos (Q1, Q2, Q3, S1, S2)."],
          ["Modalidad","Presencial, Distancia u otras modalidades."],
          ["Programa Académico","Restringe la vista a un programa específico dentro del centro."],
          ["Facultad","Filtra por facultad académica."],
          ["Nivel de Formación","Técnico, Tecnológico, Universitario, Especialización o Maestría."],
          ["Nivel Académico","Semestre o cuatrimestre activo."],
        ],2400),
        espacio(180),
 
        // ─ SUB-MÓDULO PROGRAMAS ─
        h2("2.4 Sub-módulo: Programas Académicos"),
        cuerpo("Este sub-módulo permite el análisis profundo de la matrícula por programa académico individual. Incluye el concepto de Punto de Equilibrio, la comparación entre estudiantes nuevos y graduados, y la evolución anual con tasas de variación. Es clave para la toma de decisiones sobre la oferta académica."),
        espacio(100),
        ...imagen(
          "Programas Académicos",
          "Figura 4. Sub-módulo Programas — Estudiantes por Programa Académico."
        ),

 
        h3("2.4.1 Indicador Clave: Punto de Equilibrio"),
        espacio(80),
        cajaInfo("¿Qué es el Punto de Equilibrio (P. Equilibrio)?",[
          "📊  El Punto de Equilibrio es el número mínimo de estudiantes matriculados que un programa necesita para cubrir sus costos operativos.",
          "🔢  El valor mostrado (ej. 14) indica cuántos estudiantes nuevos mínimo debe tener el programa en cada periodo para ser sostenible.",
          "⚠️  Si los estudiantes nuevos de un programa están por debajo del Punto de Equilibrio, el programa puede estar en riesgo financiero.",
          "🎯  Este indicador es fundamental para decisiones de apertura, cierre o fortalecimiento de programas académicos.",
        ],C.amarilloClaro,C.amarillo),
        espacio(140),
 
        h3("2.4.2 Filtros y Búsqueda del Sub-módulo"),
        espacio(80),
        tablaDesc([
          ["Buscar por Programa Académico (desplegable)","Selecciona un programa específico del catálogo completo. Por defecto muestra 'Todas' con el consolidado."],
          ["Buscar por SNIES (campo de búsqueda)","Permite buscar un programa directamente por su código SNIES registrado ante el Ministerio de Educación."],
          ["Selector de SNIES (botones)","Accesos directos a programas por código SNIES: 2049, 2051, 2052, 2579, 7687, 7815, 8903, 9363, 9421 (entre otros)."],
          ["Nuevos / Continuos / Totales","Cambia la vista del gráfico entre los tres tipos de matrícula."],
        ],3000),
        espacio(140),
 
        h3("2.4.3 Gráfico: Nuevos vs Graduados vs Punto de Equilibrio"),
        cuerpo("El gráfico principal del sub-módulo de Programas muestra tres series superpuestas para analizar la dinámica de cada programa:"),
        espacio(80),
        tablaDesc([
          ["Nuevos Q2 (azul)","Estudiantes nuevos que ingresan en el segundo cuatrimestre del año."],
          ["Nuevos S1 (amarillo)","Estudiantes nuevos que ingresan en el primer semestre. El periodo más importante para la mayoría de programas."],
          ["Graduados (verde)","Número de estudiantes que se gradúan en cada periodo. Permite comparar la entrada (nuevos) vs la salida (graduados) del programa."],
          ["Punto de Equilibrio (morado – línea horizontal)","Línea de referencia que indica el mínimo de estudiantes nuevos requerido para la sostenibilidad del programa."],
        ],2800),
        espacio(100),
        nota("Cuando la barra de Nuevos esté por debajo de la línea de Punto de Equilibrio en múltiples periodos consecutivos, es una señal crítica para revisar la estrategia del programa."),
        espacio(140),
 
        h3("2.4.4 Tabla: Matrículas Anuales por Programa"),
        cuerpo("A la derecha del gráfico, la tabla muestra el histórico anual desde 2020 hasta 2026 con los siguientes campos:"),
        espacio(80),
        tablaDesc([
          ["Año","Año de registro de la matrícula."],
          ["Nuevos","Total de estudiantes nuevos matriculados en el año."],
          ["Continuos","Total de estudiantes continuos en el año."],
          ["Totales","Suma de nuevos y continuos en el año."],
          ["Tasa de variación Nuevos","Variación porcentual de nuevos respecto al año anterior. 🟢 positivo (crecimiento) / 🔴 negativo (caída)."],
          ["Tasa de variación Continuos","Variación porcentual de continuos respecto al año anterior. Indica tendencia de retención."],
          ["Tasa de variación Total","Variación porcentual de la matrícula total respecto al año anterior."],
        ],2800),
        espacio(100),
        cajaInfo("Datos históricos de referencia — Toda la Sede (2020–2026)",[
          "2026: Nuevos 4.146 | Continuos 19.375 | Total 23.521  (Var. Nuevos: -3,96%)",
          "2025: Nuevos 4.317 | Continuos 19.350 | Total 23.667  (Var. Nuevos: +9,46%)",
          "2024: Nuevos 3.944 | Continuos 19.895 | Total 23.839  (Var. Nuevos: -16,83%)",
          "2023: Nuevos 4.742 | Continuos 19.916 | Total 24.658  (Var. Nuevos: +14,62%)",
          "2022: Nuevos 4.137 | Continuos 20.816 | Total 24.953  (Var. Nuevos: -10,88%)",
          "2021: Nuevos 4.642 | Continuos 22.384 | Total 27.026  (Var. Nuevos: +5,48%)",
          "2020: Nuevos 4.401 | Continuos 25.113 | Total 29.514  (línea base histórica)",
        ],C.gris),
        espacio(180),
 
        // ─ SUB-MÓDULO CARACTERIZACIÓN ─
        h2("2.5 Sub-módulo: Caracterización Histórica de Estudiantes"),
        cuerpo("Este sub-módulo muestra el perfil sociodemográfico de los estudiantes de la Sede. Permite entender quiénes son los estudiantes: su distribución por género, rango de edad y estrato socioeconómico. La vista se puede cambiar entre Nuevos, Continuos y Totales, y entre dos variables: Edad y Estrato."),
        espacio(100),
        ...imagen(
          "Caracterización",
          "Figura 5. Sub-módulo Caracterización — Distribución por género y edad de estudiantes."
        ),
 
        h3("2.5.1 Pestañas de Variable y Tipo de Matrícula"),
        espacio(80),
        tablaDesc([
          ["Edad (pestaña activa – amarillo)","Muestra la distribución de estudiantes por rangos de edad: 16–26 años, 27–36, 37–46 y Mayores de 46."],
          ["Estrato","Cambia la visualización para mostrar la distribución por estrato socioeconómico (1 al 6). Fundamental para análisis de equidad e inclusión."],
          ["Nuevos","Aplica la vista de caracterización solo a los estudiantes de primera matrícula."],
          ["Continuos","Aplica la vista solo a estudiantes que continúan sus estudios."],
          ["Totales (activo)","Muestra el perfil del total de estudiantes (nuevos + continuos). Vista por defecto."],
        ],2600),
        espacio(140),
 
        h3("2.5.2 Tarjetas de Distribución por Género"),
        cuerpo("La parte superior del dashboard muestra tres tarjetas de género para la Sede en total, la modalidad Presencial y la modalidad Distancia:"),
        espacio(80),
        new Table({ width:{size:9360,type:WidthType.DXA}, columnWidths:[3120,3120,3120],
          rows:[new TableRow({children:[
            new TableCell({borders:borde(C.azulMedio),shading:{fill:"0D2461",type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:140,right:140},children:[
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"SEDE",font:"Arial",size:20,bold:true,color:"00BFFF"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"23.559 estudiantes",font:"Arial",size:18,color:"CCCCCC"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"57% Femenino  |  43% Masculino",font:"Arial",size:19,bold:true,color:"FF69B4"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"13.377 Fem.  |  10.182 Masc.",font:"Arial",size:17,color:"AAAAAA"})]}),
            ]}),
            new TableCell({borders:borde(C.azulMedio),shading:{fill:"0A2A0A",type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:140,right:140},children:[
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"PRESENCIAL",font:"Arial",size:20,bold:true,color:"00FF7F"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"15.131 estudiantes",font:"Arial",size:18,color:"CCCCCC"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"52% Femenino  |  48% Masculino",font:"Arial",size:19,bold:true,color:"FF69B4"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"7.914 Fem.  |  7.217 Masc.",font:"Arial",size:17,color:"AAAAAA"})]}),
            ]}),
            new TableCell({borders:borde(C.azulMedio),shading:{fill:"2A0A2A",type:ShadingType.CLEAR},margins:{top:120,bottom:120,left:140,right:140},children:[
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"DISTANCIA",font:"Arial",size:20,bold:true,color:"C77DFF"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"8.428 estudiantes",font:"Arial",size:18,color:"CCCCCC"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"65% Femenino  |  35% Masculino",font:"Arial",size:19,bold:true,color:"FF69B4"})]}),
              new Paragraph({alignment:AlignmentType.CENTER,children:[new TextRun({text:"5.463 Fem.  |  2.965 Masc.",font:"Arial",size:17,color:"AAAAAA"})]}),
            ]}),
          ]})]
        }),
        espacio(100),
        nota("En la modalidad Distancia hay una mayor proporción de mujeres (65%) comparada con Presencial (52%). Este dato es clave para estrategias de bienestar, retención y comunicación diferenciada."),
        espacio(140),
 
        h3("2.5.3 Gráfico de Distribución por Edad"),
        cuerpo("El gráfico de barras agrupadas en la parte inferior muestra la distribución de estudiantes por rango de edad y modalidad. Las tres series representan Sede (azul oscuro), Presencial (verde) y Distancia (morado):"),
        espacio(80),
        tablaDesc([
          ["16 – 26 años","El grupo más numeroso: 16.232 en Sede total (12.445 Presencial + 3.787 Distancia). Es la población estudiantil tradicional."],
          ["27 – 36 años","Segundo grupo: 5.414 en Sede (2.229 Presencial + 3.185 Distancia). Nótese que en Distancia supera a Presencial en este rango: adultos que trabajan."],
          ["37 – 46 años","1.288 en Sede (291 Presencial + 997 Distancia). Predominantemente en Distancia: estudiantes trabajadores de mediana edad."],
          ["Mayores de 46","347 en Sede (93 Presencial + 254 Distancia). Aunque es el grupo más pequeño, la Distancia triplica a Presencial."],
        ],2200),
        espacio(100),
        nota("El análisis por Estrato (pestaña 'Estrato') muestra la distribución socioeconómica. Permite a la institución diseñar políticas de permanencia y apoyos focalizados según el perfil económico del estudiantado."),
        espacio(100),
        h3("2.5.4 Filtros del Sub-módulo Caracterización"),
        espacio(80),
        tablaDesc([
          ["Centro Universitario","Analiza el perfil sociodemográfico de un centro específico."],
          ["Año","Filtra por año para ver cómo ha cambiado el perfil estudiantil con el tiempo."],
          ["Periodo","Selección múltiple de periodos académicos."],
          ["Modalidad","Presencial, Distancia o todas."],
          ["Facultad","Restringe el análisis al perfil de una facultad específica."],
          ["Nivel Académico","Semestre o cuatrimestre del estudiante."],
          ["Nivel de Formación","Técnico, Tecnológico, Universitario, Especialización o Maestría."],
        ],2400),
        espacio(180),
 
        // ─ SUB-MÓDULO PERMANENCIA ─
        h2("2.6 Sub-módulo: Permanencia Estudiantil"),
        cuerpo("El sub-módulo de Permanencia es el instrumento más poderoso del 360° para analizar la retención y el riesgo de deserción. Monitorea la Tasa de Ausentismo y la Tasa de Deserción por periodo desde 2020, y analiza los graduados versus desertores por cohorte de ingreso."),
        espacio(100),
        ...imagen(
          "Permanencia",
          "Figura 6. Sub-módulo Permanencia — Ausentismo, Deserción y Análisis por Cohorte."
        ),
        
        h3("2.6.1 Búsquedas Superiores"),
        espacio(80),
        tablaDesc([
          ["Buscar por SNIES","Campo de búsqueda libre para filtrar los indicadores de permanencia de un programa específico por su código SNIES."],
          ["Buscar por Programa Académico","Desplegable para seleccionar el programa del catálogo y ver sus tasas de permanencia individuales."],
        ],2800),
        espacio(100),
        nota("Cuando estos campos tienen el valor 'Todas', los gráficos muestran el consolidado de toda la Sede. Al seleccionar un programa específico, todos los indicadores se recalculan para ese programa."),
        espacio(140),
 
        h3("2.6.2 Gráfico: Tasa de Ausentismo y Deserción por Periodo"),
        cuerpo("El gráfico de líneas superior muestra la evolución histórica desde 2020-Q1 hasta 2026-S1 de dos métricas críticas:"),
        espacio(80),
        tablaDesc([
          ["Tasa de Ausentismo (línea azul)","Porcentaje de estudiantes que no se presentan a clases o actividades en un periodo, sin llegar a retirarse formalmente. Indica el nivel de desenganche del estudiante."],
          ["Tasa de Deserción (línea amarilla)","Porcentaje de estudiantes que abandonan definitivamente sus estudios en un periodo. Es la métrica de mayor impacto en la sostenibilidad académica y financiera."],
        ],2800),
        espacio(100),
        cajaInfo("Tasas de Ausentismo y Deserción recientes (últimos periodos)",[
          "2026-S1: Ausentismo 11,41%  |  Deserción 8,87%",
          "2026-Q2: Ausentismo 14,52%  |  Deserción: (en proceso)",
          "2025-S2: Ausentismo 12,59%  |  Deserción 9,51%",
          "2025-S1: Ausentismo 13,92%  |  Deserción 9,69%",
          "2025-Q3: Ausentismo 15,46%  |  Deserción 12,01%",
          "2025-Q2: Ausentismo 15,97%  |  Deserción 12,62%",
          "2025-Q1: Ausentismo 18,65%  |  Deserción 12,07%",
          "📉  Tendencia positiva: las tasas de 2026-S1 son las más bajas de los últimos años.",
        ],C.verdeClaro,C.verde),
        espacio(140),
 
        h3("2.6.3 Gráfico: Graduados por Cohorte vs Deserción por Cohorte"),
        cuerpo("El gráfico de barras apiladas inferior analiza el destino final de cada cohorte de ingreso: qué porcentaje se gradúa y qué porcentaje deserta. Cada barra representa una cohorte (ej. 2020-Q2, 2021-S1, etc.):"),
        espacio(80),
        tablaDesc([
          ["Graduación Cohorte (azul)","Porcentaje de los estudiantes que ingresaron en esa cohorte y que lograron graduarse."],
          ["Deserción Cohorte (amarillo)","Porcentaje de los que ingresaron en esa cohorte y que abandonaron sus estudios sin graduarse."],
        ],2800),
        espacio(100),
        nota("Las cohortes más recientes (2025-S1, 2025-S2) no tienen datos completos de graduación porque los estudiantes aún están cursando. Los valores aparecen con un ícono de advertencia en la tabla lateral."),
        espacio(140),
 
        h3("2.6.4 Tabla Lateral: Tasas de Graduación y Deserción por Cohorte"),
        cuerpo("A la derecha del gráfico de cohortes, la tabla muestra el resumen numérico por cada cohorte. Los iconos ✅ y ❌ indican si la tasa de graduación es satisfactoria o está por debajo del umbral esperado."),
        espacio(100),
        h3("2.6.5 Filtros del Sub-módulo Permanencia"),
        espacio(80),
        tablaDesc([
          ["Centro Universitario","Analiza permanencia de un centro específico para comparar con otros."],
          ["Periodo","Filtra por periodos de interés."],
          ["Modalidad","Presencial vs Distancia: las tasas de deserción suelen ser diferentes entre modalidades."],
          ["Facultad","Ver permanencia por facultad permite identificar cuáles tienen mayor desafío de retención."],
          ["Nivel de Formación","Los programas T&T suelen tener tasas de ausentismo diferentes a los profesionales."],
          ["Nivel Académico","Semestre o cuatrimestre específico del estudiante."],
        ],2400),
        espacio(180),
 
        // ─ RESUMEN MÓDULO ESTUDIANTES ─
        h2("2.7 Resumen: ¿Cuándo usar cada sub-módulo de Estudiantes?"),
        espacio(80),
        tablaConHeader("Guía de uso por necesidad dentro del módulo Estudiantes",[
          ["Ver el total de estudiantes y KPIs principales de la Sede",           "Ir a: Inicio del módulo → Tarjetas de Resumen General"],
          ["Ver evolución histórica de la matrícula (desde 2020)",                "Ir a: Sub-módulo Sede → Gráfico de barras históricas"],
          ["Comparar matrícula entre centros universitarios",                     "Ir a: Sub-módulo Centro Universitario → Tabla por Centro"],
          ["Ver qué facultad tiene más o menos estudiantes",                     "Ir a: Sub-módulo Centro Universitario → Tabla por Escuela"],
          ["Analizar un programa académico específico",                           "Ir a: Sub-módulo Programas → Buscar por Programa"],
          ["Ver si un programa está por encima/debajo del Punto de Equilibrio",  "Ir a: Sub-módulo Programas → Indicador P. Equilibrio"],
          ["Conocer el perfil de género y edad del estudiantado",                "Ir a: Sub-módulo Caracterización → Pestañas Edad / Estrato"],
          ["Monitorear deserción y ausentismo",                                  "Ir a: Sub-módulo Permanencia → Gráfico Ausentismo y Deserción"],
          ["Analizar el destino de una cohorte de ingreso",                      "Ir a: Sub-módulo Permanencia → Gráfico Graduados vs Deserción por Cohorte"],
        ],3800),
        espacio(180),
        lineaDivisora(),
 

        // ───────────────────────────────────────
        // MÓDULO 3 – PROFESORES
        // ───────────────────────────────────────
        h1("Módulo 3 — Profesores"),
        cuerpo("El módulo de Profesores permite analizar la planta docente de la Sede desde múltiples perspectivas: nivel académico, tipo de dedicación, escalafón, funciones sustantivas y tipo de contrato. Cada sub-módulo presenta una tabla de datos históricos, un gráfico de barras por periodo y una tendencia histórica de línea."),
        espacio(100),
        ...imagen("Profesores", "Figura 3. Módulo de Profesores — Vista por Nivel Académico."),

        h2("3.1 Sub-módulos del Módulo de Profesores"),
        cuerpo("Los sub-módulos se acceden desde los botones de la barra superior (el botón activo aparece resaltado en amarillo):"),
        espacio(80),
        tablaDesc([
          ["Dedicación",           "Muestra la distribución de docentes por tipo de dedicación: Tiempo Completo, Medio Tiempo y Hora Cátedra. Permite ver la evolución por periodo."],
          ["Escalafón",            "Clasifica a los profesores según su categoría en el escalafón docente institucional. Útil para planificación de carrera docente."],
          ["Funciones Sustantivas","Distribución de profesores por las tres funciones sustantivas: Docencia, Investigación y Proyección Social."],
          ["Nivel Académico",      "Clasifica la planta docente por su máximo nivel de formación: Doctorado, Maestría, Especialización, Universitario y Técnico. Vista activa en la captura."],
          ["Tipo de Contrato",     "Muestra la distribución entre contratos de planta fija, término fijo y por prestación de servicios (hora cátedra)."],
        ], 2600),
        espacio(140),

        h2("3.2 Panel de Filtros (Columna Izquierda)"),
        espacio(80),
        tablaDesc([
          ["Tipo de Trabajador", "Permite filtrar entre Profesores, Investigadores u otros tipos de personal académico con actividades de docencia."],
          ["Modalidad",          "Filtra por la modalidad en que dictan sus clases: Presencial, Distancia u otra."],
          ["Género",             "Segmenta la planta docente por género para análisis de equidad e inclusión."],
        ], 2500),
        espacio(140),

        h2("3.3 Lectura de las Visualizaciones"),
        cuerpo("Cada sub-módulo presenta dos gráficos complementarios:"),
        espacio(80),
        tablaDesc([
          ["Tabla superior",         "Muestra los valores numéricos por categoría y periodo (ej: 2024-1, 2024-2, 2025-1, 2025-2, 2026-1). Permite hacer comparaciones directas entre periodos."],
          ["Barras por período",     "Gráfico de barras agrupadas que muestra la distribución de cada categoría por periodo académico reciente. Ideal para ver cambios de corto plazo."],
          ["Tendencia Histórica",    "Gráfico de líneas que muestra la evolución desde 2024-1 hasta el periodo más reciente. Permite identificar tendencias, picos y caídas en la planta docente."],
        ], 2600),
        espacio(80),
        nota("Los valores en color azul (hipervínculo) dentro de la tabla superior son clicables y filtran automáticamente las visualizaciones hacia ese valor específico."),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 4 – INVESTIGACIÓN
        // ───────────────────────────────────────
        h1("Módulo 4 — Investigación (I+D)"),
        cuerpo("El módulo de Investigación, administrado por la Dirección de Investigación I+D de la Vicerrectoría Académica, ofrece una vista completa de la actividad investigativa de la Sede. Muestra grupos, investigadores, productos y semilleros con capacidad de filtrado avanzado."),
        espacio(100),
        ...imagen("Investigación", "Figura 4. Módulo de Investigación I+D — Vista de grupos y productos."),

        h2("4.1 Indicadores Resumen (Tarjetas Superiores)"),
        espacio(80),
        tablaDesc([
          ["Cantidad de Grupos",   "Número total de grupos de investigación activos registrados en la Sede (ej. 46 grupos)."],
          ["Investigadores",       "Número total de investigadores vinculados a los grupos activos (ej. 492 investigadores)."],
          ["Cantidad de Productos","Número total de productos académicos e investigativos registrados (ej. 24.979 productos)."],
        ], 2600),
        espacio(140),

        h2("4.2 Panel de Filtros Superiores"),
        espacio(80),
        tablaDesc([
          ["Nombre del Grupo",          "Filtra para ver los datos de un grupo de investigación específico."],
          ["Escuela",                   "Filtra por la escuela o facultad a la que pertenece el grupo."],
          ["Nivel de Formación",        "Restringe a grupos asociados con programas de un nivel de formación específico."],
          ["Categoría Líder",           "Categoría Minciencias del investigador líder del grupo (A1, A, B, C, Reconocido)."],
          ["Categoría Grupo 2024",      "Clasificación actual del grupo según la última convocatoria Minciencias."],
          ["Tipo y Tipología Producto", "Filtra por categoría de producto (Divulgación Pública, Generación de Nuevo Conocimiento, Formación de Recursos Humanos, etc.)."],
          ["Año",                       "Selecciona el año de registro o publicación de los productos para análisis temporal."],
        ], 2600),
        espacio(140),

        h2("4.3 Visualizaciones Principales"),
        espacio(80),
        tablaDesc([
          ["Porcentaje Tipología (barras horizontales)",    "Distribución porcentual de productos por tipología. La 'Divulgación Pública de la Ciencia' concentra el mayor porcentaje (ej. 12.998 productos)."],
          ["Productos por Grupos de Investigación",        "Ranking de grupos ordenados por su número total de productos registrados. Permite identificar los grupos más productivos."],
          ["Tipos de Producto (barras horizontales)",      "Desglose detallado por tipo específico: Eventos científicos, Tesis de pregrado, Artículos en revistas, Capítulos de libro, etc."],
          ["Tabla inferior de grupos y tipologías",        "Matriz cruzada que muestra por cada grupo de investigación cuántos productos tiene en cada tipología, con totales por fila y columna."],
          ["Panel 'Información Producto' (derecha)",       "Al hacer clic en una barra o fila, aparece el listado de títulos específicos de los productos del grupo o tipología seleccionada."],
        ], 2900),
        espacio(140),

        h2("4.4 Sub-módulos de Navegación Superior (Investigación)"),
        espacio(80),
        tablaDesc([
          ["Inicio",                   "Regresa a la vista principal del módulo de Investigación con el resumen general."],
          ["Grupos de Investigación",  "Vista detallada de cada grupo: integrantes, líder, clasificación y líneas de investigación."],
          ["Productos",                "Exploración de todos los productos registrados con capacidad de búsqueda y filtrado por tipo."],
          ["Investigadores",           "Perfil y datos de cada investigador: grupos en que participa, productos a su nombre y categoría."],
          ["Semilleros",               "Información sobre semilleros de investigación activos: número de integrantes, proyectos y vinculación."],
        ], 2600),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 5 – PROYECCIÓN SOCIAL
        // ───────────────────────────────────────
        h1("Módulo 5 — Proyección Social"),
        cuerpo("Al hacer clic en 'Proyección Social' en el menú principal, el sistema te muestra una pantalla de navegación con tres sub-módulos diferenciados, cada uno con su propio dashboard. Desde aquí seleccionas en cuál quieres profundizar."),
        espacio(100),
        ...imagen("Proyección Social", "Figura 5. Pantalla de selección del módulo Proyección Social."),

        h2("5.1 Tres Sub-módulos de Proyección Social"),
        espacio(80),
        tablaTarjetas([
          { icono: "🎓", titulo: "Egresados", subtitulo: "Seguimiento · Impacto", desc: "Seguimiento a graduados, vinculación laboral e impacto en la sociedad. Incluye filtros por Bogotá, modalidad, nivel académico, nivel de formación, sede, centro universitario, escuela, programas, año y periodo de grado, títulos y territorio.", estado: "activo" },
          { icono: "🌐", titulo: "Centro de Educación para el Desarrollo (CED)", subtitulo: "Educación · Desarrollo", desc: "Programas CED, educación continua y desarrollo comunitario. Dashboard con indicadores de alcance e impacto de los programas de extensión.", estado: "activo" },
          { icono: "🛡️", titulo: "Centro Progresa", subtitulo: "Bienestar · Orientación", desc: "Orientación vocacional, emprendimiento y acompañamiento estudiantil. Indicadores de atención, programas de bienestar y apoyo al proyecto de vida del estudiante.", estado: "activo" },
        ]),
        espacio(140),

        h2("5.2 Sub-módulo: Egresados (Detalles)"),
        cuerpo("El dashboard de Egresados es el más robusto dentro de Proyección Social. Cuenta con los siguientes filtros y vistas:"),
        espacio(80),
        tablaDesc([
          ["Filtro Bogotá",               "Por defecto el sistema carga los datos de Bogotá. Puedes cambiar para ver Cundinamarca, Boyacá u otras sedes."],
          ["Modalidad",                   "Filtra por tipo de modalidad de estudio del egresado."],
          ["Nivel Académico",             "Segmenta por el nivel académico cursado (semestre o cuatrimestre del egresado)."],
          ["Nivel de Formación",          "Filtra entre Técnico, Tecnológico, Profesional, Especialización o Maestría."],
          ["Por Sede / Centro Universitario / Escuelas", "Navegación jerárquica desde la Sede hasta la escuela específica de la que egresó el graduado."],
          ["Nivel AC y Programas",        "Permite ver egresados de un programa académico concreto dentro de una escuela."],
          ["Gráfico Programas",           "Visualización de los programas con mayor número de egresados."],
          ["Año y Periodo de Grado",      "Análisis histórico de graduaciones: cuántos estudiantes se graduaron en cada año y periodo."],
          ["Títulos y Territorio",        "Distribución geográfica de los títulos otorgados: en qué territorios están los egresados de UNIMINUTO."],
        ], 3000),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 6 – SEGUIMIENTO A LA ESTRATEGIA
        // ───────────────────────────────────────
        h1("Módulo 6 — Seguimiento a la Estrategia"),
        cuerpo("Al hacer clic en 'Seg. Estrategia' en el menú, el sistema muestra una pantalla de selección con cinco herramientas de seguimiento estratégico. Cada tarjeta indica si está disponible (ABRIR ENLACE), si se puede copiar el enlace o si está próximamente disponible."),
        espacio(100),
        ...imagen("Seguimiento a la Estrategia", "Figura 6. Pantalla de selección del módulo Seguimiento a la Estrategia."),


        h2("6.1 Cinco Componentes del Módulo Estratégico"),
        espacio(80),
        tablaConHeader("Componentes de Seguimiento a la Estrategia", [
          ["Plan Estratégico (INSTITUCIONAL)",          "PRÓXIMAMENTE. Cubrirá objetivos institucionales, metas y KPIs estratégicos de la Sede. Aún en desarrollo."],
          ["Plan de Centros Universitarios (TERRITORIAL)", "DISPONIBLE. Metas por centro, seguimiento territorial y avance por centro universitario. Botones: ABRIR ENLACE / COPIAR ENLACE."],
          ["Proyecta UNIMINUTO (INDICADORES – KPI)",    "DISPONIBLE. Tablero de Real vs Presupuesto vs Proyección de los principales indicadores financieros y académicos. Botones: ABRIR ENLACE / COPIAR ENLACE."],
          ["Indicadores (KPIs – MÉTRICAS)",             "PRÓXIMAMENTE. Tablero de control con avance de metas institucionales. En desarrollo."],
          ["Sede en Cifras (RESUMEN – DATOS)",          "DISPONIBLE. Datos clave de la Sede, comparativo anual y boletín estadístico. Botones: ABRIR ENLACE / COPIAR ENLACE."],
        ], 3200),
        espacio(140),
        nota("Los componentes con 'ABRIR ENLACE' te redirigen directamente al reporte en Power BI. 'COPIAR ENLACE' te da el URL para compartir con otros usuarios."),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 7 – GESTIÓN POR PROCESOS
        // ───────────────────────────────────────
        h1("Módulo 7 — Gestión por Procesos"),
        cuerpo("Al hacer clic en 'Gestión por Procesos' desde la barra inferior de la pantalla de inicio, el sistema abre el chatbot institucional. Actualmente se llama AVA y próximamente será renombrado como EMA (Ecosistema de Modelos de IA Aplicada)."),
        espacio(80),
        cajaInfo("¿Qué es AVA / EMA?", [
          "🤖  AVA es el asistente virtual institucional del 360°. Puedes hacerle preguntas sobre la Sede, los módulos del sistema y los datos disponibles.",
          "🔄  PRÓXIMAMENTE: AVA será renombrado EMA (Ecosistema de Modelos de IA Aplicada), con capacidades ampliadas.",
          "💬  Interactúas con el asistente a través de una interfaz de chat. Escribe tu pregunta y el bot responde con información institucional.",
          "📋  Casos de uso: consulta de indicadores, orientación sobre qué módulo usar, respuestas a preguntas frecuentes de la Sede.",
        ], C.amarilloClaro, C.amarillo),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 8 – PLANTA FÍSICA
        // ───────────────────────────────────────
        h1("Módulo 8 — Ocupación Planta Física"),
        cuerpo("El módulo de Planta Física permite monitorear en tiempo real la ocupación de salones y edificios en cada centro universitario, por jornada y día de la semana. Es clave para la optimización de espacios físicos y la planificación académica."),
        espacio(100),
        ...imagen("Planta Física", "Figura 7. Módulo de Ocupación Planta Física."),

        h2("8.1 Panel de Filtros (Columna Izquierda)"),
        espacio(80),
        tablaDesc([
          ["Centro Universitario", "Filtra la vista para un centro universitario específico (ej. Engativá, Perdomo, Kennedy, Santa Fe, Usaquén)."],
          ["Año",                  "Selecciona el año de análisis. Por defecto muestra el año más reciente (ej. 2025)."],
          ["Periodo",              "Filtra por periodo académico dentro del año seleccionado."],
          ["Edificio",             "Restringe la vista a un edificio específico dentro del centro universitario (ej. CBRGH, CIPAD, CL90PC, etc.)."],
          ["Jornada",              "Filtra por jornada: 1-Mañana, 2-Tarde, 3-Noche."],
          ["Día de la Semana",     "Permite ver la ocupación de un día específico: Lunes a Sábado."],
        ], 2500),
        espacio(140),

        h2("8.2 Panel de Capacidad — Tabla de Edificios y Salones"),
        cuerpo("En la parte central-izquierda verás la tabla de planta física con el inventario de edificios y número de salones disponibles por edificio. Los edificios clave para Bogotá son:"),
        espacio(80),
        cajaInfo("Edificios principales registrados en el sistema", [
          "CBRGH — 46 salones  |  CIPAD — 5 salones  |  CL90PC — 18 salones  |  CLL-90 — 14 salones",
          "COLSL — 24 salones  |  DJC — 32 salones  |  RGH — 41 salones  |  SJE — 43 salones",
          "SEDAL — 21 salones  |  SAA — 4 salones  |  SAB — 5 salones  |  COPRO — 7 salones",
          "TOTAL registrado: 277 salones activos en la Sede.",
        ], C.gris),
        espacio(140),

        h2("8.3 Tablas de Ocupación por Jornada"),
        cuerpo("El sistema presenta dos tablas de ocupación ubicadas en la parte central-derecha:"),
        espacio(80),
        tablaDesc([
          ["Ocupación salones por jornada (por EDIFICIO)", "Muestra para cada edificio el porcentaje de uso en mañana, tarde, noche y el total. Ejemplo: CBRGH tiene 46 salones con uso de 43-Mañana, 25-Tarde, 34-Noche."],
          ["Ocupación salones por jornada (por SALÓN)",    "Tabla detallada que muestra el porcentaje de ocupación individual por salón. Ejemplo: Salón 1001 con 84,67% en la mañana. Porcentajes por encima del 100% indican sobre-utilización."],
        ], 3200),
        espacio(140),

        h2("8.4 Gráfico: % de Uso por Día y Jornada"),
        cuerpo("El gráfico de barras en la parte inferior muestra el porcentaje de uso de salones para cada combinación de día de la semana y jornada (Lunes Mañana, Lunes Tarde, Lunes Noche, ..., Sábado Noche). Permite identificar cuáles días y jornadas tienen mayor o menor demanda de espacios para tomar decisiones de asignación."),
        espacio(80),
        nota("La información de este módulo proviene del Sistema de Registro Académico con fecha de corte indicada en la parte inferior de la pantalla (ej. 20 de noviembre de 2025)."),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 9 – MERCADEO (SALESFORCE)
        // ───────────────────────────────────────
        h1("Módulo 9 — Mercadeo: Solicitudes de Inscripción Salesforce"),
        cuerpo("El módulo de Mercadeo consolida las solicitudes de inscripción registradas en Salesforce para la Sede Bogotá. Es la herramienta principal para analizar el embudo de captación: desde la primera inscripción hasta la matrícula efectiva. Permite comparar modalidades Distancia vs Presencial."),
        espacio(100),
        ...imagen("Mercadeo", "Figura 8. Módulo de Mercadeo — Solicitudes de Inscripción Salesforce."),

        h2("9.1 Indicadores Resumen (Tarjetas Superiores)"),
        espacio(80),
        tablaDesc([
          ["Total General",  "Número total de solicitudes de inscripción registradas en el periodo visualizado (ej. 13.780 solicitudes)."],
          ["Cuatrimestral",  "Total de solicitudes correspondientes a programas cuatrimestrales (ej. 6.282)."],
          ["Semestral",      "Total de solicitudes correspondientes a programas semestrales (ej. 7.497)."],
          ["Solicitudes",    "Botón que lleva a la vista detallada de todas las solicitudes individuales."],
        ], 2200),
        espacio(140),

        h2("9.2 Pestañas de Modalidad"),
        cuerpo("En la parte superior del dashboard hay dos botones que segmentan todo el análisis por modalidad:"),
        espacio(80),
        tablaDesc([
          ["Distancia",   "Muestra solo las solicitudes de programas en modalidad a distancia (virtual/tradicional)."],
          ["Presencial",  "Muestra solo las solicitudes de programas en modalidad presencial."],
        ], 2000),
        espacio(140),

        h2("9.3 Gráfico de Captación (Línea de Tendencia)"),
        cuerpo("El gráfico de línea muestra la evolución mensual del número de solicitudes desde los últimos meses. La tabla adjunta muestra el detalle por mes con variaciones:"),
        espacio(80),
        tablaDesc([
          ["Solicitudes del mes",  "Número absoluto de inscripciones recibidas en cada mes."],
          ["Var Año %",            "Variación porcentual respecto al mismo mes del año anterior. Un valor positivo indica crecimiento en captación."],
          ["Var Mes %",            "Variación porcentual respecto al mes inmediatamente anterior. Indica la tendencia de corto plazo."],
          ["Dif Abs Mes",          "Diferencia absoluta en número de solicitudes respecto al mes anterior."],
        ], 2500),
        espacio(140),

        h2("9.4 Visualizaciones del Panel Central"),
        espacio(80),
        tablaConHeader("Secciones del dashboard de Mercadeo", [
          ["Programa de Formación (barras horizontales)",   "Ranking de programas ordenados de mayor a menor número de solicitudes. Los programas más demandados aparecen en la cima (ej. Contaduría Pública 1.392, Administración de Empresas 1.291)."],
          ["Centro Universitario (barras verticales)",      "Distribución de solicitudes por centro. Muestra cuáles centros atraen más prospectos (ej. Engativá 9.130, Perdomo 3.105, Kennedy 816)."],
          ["Estado de Candidato (dona)",                    "Distribución porcentual de las solicitudes según su estado en el embudo: Matriculado (5%), Pre-Admitido (46%), Inscrito (22%), Admitido (9%), Pre-inscrito (16%), Desistido, Retirado."],
          ["Matrículas por Campaña",                        "Tabla cruzada que muestra matrículas efectivas por Jornada (Mañana, Tarde, Noche, Distancia) y periodo (Q1, Q2, S1, S2)."],
          ["Tasa de Inscripción Estudiantes Nuevos",        "Gráfico de líneas que muestra la tasa de conversión desde inscripción hasta matrícula por jornada y cohorte."],
          ["Sub-Estado de Etapa",                           "Tabla con el detalle de por qué las solicitudes están en cada estado (ej. 'Ya realizó el pago', 'Va a pagar de contado', 'Retirado por gestión administrativa')."],
        ], 3200),
        espacio(140),

        h2("9.5 Panel de Filtros (Columna Izquierda)"),
        espacio(80),
        tablaDesc([
          ["Año, Mes, Día",         "Selecciona un rango temporal específico para analizar periodos de mayor captación."],
          ["Periodo",               "Filtra por periodo académico (Q1, Q2, Q3, S1, S2)."],
          ["Centro Universitario",  "Restringe el análisis a un punto de atención específico."],
          ["Etapa de la Admisión",  "Filtra por la etapa del proceso (inscrito, admitido, pre-admitido, matriculado, etc.)."],
          ["Nivel de Formación",    "Filtra entre niveles de formación de los programas ofertados."],
          ["Tipo de Vinculación",   "Diferencia entre nuevos estudiantes o transferencias."],
          ["Programa Académico",    "Análisis granular de un programa específico en el embudo de captación."],
        ], 2400),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 10 – CONTEXTO EXTERNO
        // ───────────────────────────────────────
        h1("Módulo 10 — Contexto Externo"),
        cuerpo("El módulo de Contexto Externo reúne información de fuentes externas que permiten a la Sede entender su posición en el ecosistema de educación superior colombiano. Es fundamental para decisiones de oferta académica, precios y posicionamiento. Al ingresar verás un menú con 9 sub-módulos accesibles como tarjetas."),
        espacio(100),
        ...imagen("Contexto Externo", "Figura 9. Menú del módulo Contexto Externo con sus nueve sub-módulos."),

        h2("10.1 Los 9 Sub-módulos de Contexto Externo"),
        espacio(80),
        tablaConHeader("Sub-módulos de Contexto Externo", [
          ["Resumen cifras SNIES – SPADIES",  "Vista consolidada de los principales indicadores de educación superior en Colombia (matrícula nacional, deserción) según SNIES y SPADIES. Punto de partida recomendado."],
          ["Precios Externos UNIMINUTO (P)",  "Comparativo de precios de programas de UNIMINUTO frente a otras IES competidoras en la región. Esencial para decisiones de matrícula y tarifas."],
          ["Demografía",                      "Datos demográficos de Bogotá, Cundinamarca y Boyacá: distribución por edades, niveles educativos y potencial de demanda de educación superior."],
          ["Data SNIES",                      "Base de datos completa del Sistema Nacional de Información de Educación Superior. Matrícula, graduados, programas y IES a nivel nacional."],
          ["Observatorio Ocupación SENA",     "Datos del Observatorio Laboral y Ocupacional del SENA: tendencias del mercado laboral, ocupaciones más demandadas y relación con los programas de UNIMINUTO."],
          ["Geolocalización por IES",         "Mapa interactivo con la ubicación geográfica de las IES competidoras en Bogotá, Cundinamarca y Boyacá. Identifica zonas de alta concentración de oferta educativa."],
          ["Data SPADIES",                    "Datos detallados del sistema SPADIES sobre deserción estudiantil: tasas por programa, institución, periodo y características sociodemográficas."],
          ["Data Programas Externos (P)",     "Detalle de los programas académicos ofrecidos por las IES competidoras: nombre, precio, modalidad y ubicación."],
          ["Geolocalización por Colegios",    "Mapa con la ubicación de colegios de bachillerato que alimentan la matrícula de educación superior. Útil para estrategias de captación territorial."],
        ], 2800),
        espacio(80),
        nota("La etiqueta '(P)' junto al nombre de algunos sub-módulos indica que el acceso puede requerir permisos especiales. Consulta con la Dirección de Planeación si no puedes acceder."),
        espacio(80),
        nota("El botón 'Menú Principal' en la parte izquierda de la pantalla te regresa a la vista de selección de sub-módulos de Contexto Externo."),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 11 – INTERNACIONALIZACIÓN
        // ───────────────────────────────────────
        h1("Módulo 11 — Internacionalización"),
        cuerpo("Al hacer clic en 'Internacionalización' desde la barra inferior del inicio, el sistema muestra la pantalla de este módulo con un indicador claro de su estado actual."),
        espacio(100),
        ...imagen("Internacionalización", "Figura 10. Módulo de Internacionalización — Actualmente en construcción."),

        espacio(60),
        enConstruccion(),
        espacio(80),
        cajaInfo("¿Qué cubrirá el módulo de Internacionalización?", [
          "✈️  Movilidad académica entrante y saliente: estudiantes y docentes en intercambio.",
          "🤝  Convenios internacionales: listado, estado y tipo de los acuerdos con instituciones del exterior.",
          "🌍  Cooperación global: proyectos conjuntos, co-investigación y programas con universidades socias.",
          "📊  KPIs de internacionalización: índices de movilidad, idiomas, doble titulación.",
          "📅  Estado: El equipo de Asuntos Internacionales está preparando el módulo. Progreso: 'En desarrollo'.",
        ], C.rojoClaro, C.rojo),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // MÓDULO 12 – CIUDAD VAC
        // ───────────────────────────────────────
        h1("Módulo 12 — Ciudad VAC (Vicerrectoría Académica)"),
        cuerpo("Ciudad VAC es un portal visual interactivo de la Vicerrectoría Académica de UNIMINUTO. Al ingresar verás una ilustración tipo 'ciudad' donde cada edificio representa una dirección o unidad de la Vicerrectoría. Haciendo clic en cada etiqueta accedes a la información de esa dirección."),
        espacio(100),
        ...imagen("Ciudad VAC", "Figura 11. Portal Ciudad VAC — Vicerrectoría Académica UNIMINUTO."),

        h2("12.1 Menú de Navegación de Ciudad VAC"),
        espacio(80),
        tablaDesc([
          ["Principal",      "Pantalla de inicio con el mapa visual de la ciudad donde cada edificio representa una dirección académica. Es la vista que se muestra en la captura."],
          ["BotVAC",         "Chatbot de la Vicerrectoría Académica. Asistente virtual para consultas relacionadas con procesos académicos."],
          ["Mia & Leo",      "Asistentes o personajes guía de la Vicerrectoría Académica para orientación a la comunidad universitaria."],
          ["About Us",       "Información institucional sobre la Vicerrectoría Académica: misión, visión, equipo directivo y funciones."],
          ["Servicios VAC",  "Catálogo de servicios que ofrece la Vicerrectoría Académica a estudiantes, docentes y directivos."],
        ], 2300),
        espacio(140),

        h2("12.2 Direcciones Accesibles desde el Mapa Visual"),
        cuerpo("En el mapa de Ciudad VAC puedes hacer clic en las etiquetas de cada edificio para acceder a la información de cada dirección:"),
        espacio(80),
        tablaConHeader("Direcciones de la Vicerrectoría Académica en Ciudad VAC", [
          ["Dirección de Investigación I+D",                  "Acceso a datos de grupos, investigadores y productos. Conectado con el módulo de Investigación del 360°."],
          ["Dirección Currículo, Enseñanza y Evaluación",     "Información sobre diseño curricular, evaluación de programas y estándares pedagógicos."],
          ["Dirección de Innovación y Transformación Académica", "Proyectos de innovación educativa, transformación digital y nuevos modelos pedagógicos."],
          ["Dirección Internacionalización Académica – Asuntos Globales", "Información sobre convenios, movilidad y cooperación internacional (articula con el módulo de Internacionalización del 360°)."],
          ["Observatorio de Datos",                           "Repositorio de datos y análisis institucionales. Fuente de información para indicadores del 360°."],
          ["Dirección Calidad Académica",                     "Procesos de acreditación, autoevaluación y mejora continua de programas académicos."],
        ], 3300),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // GUÍA RÁPIDA DE NAVEGACIÓN
        // ───────────────────────────────────────
        h1("Guía Rápida: ¿Qué necesito? ¿Dónde lo encuentro?"),
        cuerpo("Esta tabla te ayuda a identificar rápidamente en qué módulo del 360° encontrarás la información que necesitas:"),
        espacio(100),
        tablaConHeader("Guía de necesidades vs. módulo del 360°", [
          ["¿Cuántos estudiantes nuevos ingresaron este semestre?",           "→ Módulo Estudiantes  |  Tipo de Vista: Nuevos  |  Filtro: Periodo"],
          ["¿Cómo está la retención estudiantil?",                           "→ Módulo Estudiantes  |  Tipo de Vista: Continuos  |  Comparar periodos"],
          ["¿Cuántos docentes con doctorado tiene la Sede?",                 "→ Módulo Profesores  |  Sub-módulo: Nivel Académico  |  Ver tabla"],
          ["¿Cuántos grupos de investigación hay activos?",                  "→ Módulo Investigación  |  Tarjeta 'Cantidad de Grupos' en el resumen"],
          ["¿Cuántos egresados se han graduado por programa?",               "→ Módulo Proyección Social  |  Sub-módulo: Egresados  |  Gráfico Programas"],
          ["¿Cómo va el avance del Plan de Centros Universitarios?",         "→ Módulo Seg. Estrategia  |  Tarjeta: Plan de Centros Universitarios"],
          ["¿Cuántas solicitudes de inscripción hay para el próximo periodo?","→ Módulo Mercadeo  |  Gráfico Captación  |  Tabla Estado de Candidato"],
          ["¿Qué tan ocupados están los salones de un edificio?",            "→ Módulo Planta Física  |  Filtro: Edificio + Jornada"],
          ["¿Cuánto cobra la competencia por un programa similar?",          "→ Módulo Contexto Externo  |  Sub-módulo: Precios Externos UNIMINUTO"],
          ["¿Cuánta deserción hay a nivel nacional en mi programa?",         "→ Módulo Contexto Externo  |  Sub-módulo: Data SPADIES"],
          ["¿Dónde están ubicados los colegios que nos envían estudiantes?", "→ Módulo Contexto Externo  |  Sub-módulo: Geolocalización por Colegios"],
          ["¿Cuál es la demanda laboral para los egresados de mi programa?", "→ Módulo Contexto Externo  |  Sub-módulo: Observatorio Ocupación SENA"],
          ["Necesito hablar con el asistente virtual de la Sede",            "→ Módulo Gestión por Procesos  |  Chatbot AVA/EMA"],
          ["Quiero ver los indicadores consolidados de la Vicerrectoría",    "→ Ciudad VAC  |  Observatorio de Datos o módulos específicos"],
        ], 4200),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // BUENAS PRÁCTICAS
        // ───────────────────────────────────────
        h1("Buenas Prácticas y Consejos de Uso"),
        espacio(80),
        cajaInfo("10 consejos para sacar el máximo provecho al 360°", [
          "1.  Siempre verifica qué filtros están activos antes de interpretar un gráfico. Un filtro olvidado puede mostrar datos parciales y llevar a conclusiones incorrectas.",
          "2.  Usa el botón 'Limpiar Filtros' cada vez que quieras volver a la vista general de la institución sin segmentaciones.",
          "3.  Para análisis comparativos, abre el mismo módulo en dos pestañas del navegador con diferentes filtros activos.",
          "4.  Los valores en azul (hipervínculo) en las tablas son clicables: al hacer clic filtran automáticamente todas las visualizaciones de la pantalla.",
          "5.  Para exportar un gráfico o datos, usa el menú de tres puntos (…) que aparece al pasar el cursor sobre cada visual en Power BI.",
          "6.  En gráficos con muchos periodos, usa la barra de desplazamiento horizontal para ver todos los datos disponibles.",
          "7.  La app Power BI Mobile (iOS/Android) permite acceder al 360° desde tu celular con la misma cuenta institucional.",
          "8.  Los datos tienen una fecha de corte. Revisa la nota al pie de cada pantalla para saber cuándo fue la última actualización.",
          "9.  Si un porcentaje de ocupación supera el 100% en Planta Física, indica que hay más clases asignadas que capacidad disponible. Es una señal de alerta.",
          "10. Para reportar errores en datos o solicitar nuevos accesos, contacta a la Dirección de Planeación y Desarrollo.",
        ], C.verdeClaro, C.verde),
        espacio(180),
        lineaDivisora(),

        // ───────────────────────────────────────
        // SOPORTE
        // ───────────────────────────────────────
        h1("Soporte Técnico y Contacto"),
        cuerpo("Para consultas, reportes de errores, solicitudes de acceso o sugerencias de mejora del sistema 360°, comunícate con la Dirección de Planeación y Desarrollo de la Sede."),
        espacio(80),
        cajaInfo("Dirección de Planeación y Desarrollo — UNIMINUTO Sede Bogotá – Cundinamarca – Boyacá", [
          "📍  Sede Bogotá – Cundinamarca – Boyacá",
          "🌐  Acceso al sistema: https://uniminuto0.sharepoint.com",
          "⏰  Horario de atención: lunes a viernes, 8:00 a.m. – 5:00 p.m.",
          "📧  Para solicitar acceso a módulos restringidos: contacta a tu director o coordinador de área.",
          "🐛  Para reportar errores en los datos: indica el módulo, la fecha, los filtros activos y el valor incorrecto observado.",
        ], C.azulClaro),
        espacio(200),

        // Pie de página del documento
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { before: 200, after: 80 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.amarillo, space: 4 } },
          children: [new TextRun({ text: "UNIMINUTO — Educación de calidad al alcance de todos", font: "Arial", size: 22, italics: true, bold: true, color: C.azulOscuro })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "360° Ecosistema de Inteligencia  |  Dirección de Planeación y Desarrollo  |  Sede Bogotá – Cundinamarca – Boyacá  |  2025", font: "Arial", size: 18, color: "888888" })],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('./SuperManual_360_UNIMINUTO2.docx', buffer);
  console.log('✅ Super manual generado correctamente.');
});