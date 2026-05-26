// ============================================================
// PROSPECCIÓN CASTELLÓN — Creador automático de Google Sheet
// Instrucciones:
//   1. Ve a script.google.com
//   2. Crea un proyecto nuevo
//   3. Pega este código completo
//   4. Pulsa ▶ Ejecutar (función: crearSheetProspeccion)
//   5. Acepta los permisos
//   6. ¡Listo! Se abrirá el Google Sheet creado
// ============================================================

function crearSheetProspeccion() {

  const ss = SpreadsheetApp.create("Prospectos Castellón de la Plana — Diseño Web");
  const sheet = ss.getActiveSheet();
  sheet.setName("Prospectos");

  // ── DATOS ──────────────────────────────────────────────────
  const datos = [
    [1,  "Luis Sos Fontanería",                "Fontanería",          "",             "",                              "Castellón de la Plana",                       "",                      "Facebook: /fontaneria.luissos",   "sin_web",          9.5, "ALTA",       "Más de 60 años. Sin web propia. Solo Facebook."],
    [2,  "Joyería Font",                        "Joyería / Relojería", "964 242 618",  "joyeriafontsl@hotmail.com",     "C/ Alloza, 68 Bajo, 12001 Castellón",         "",                      "Facebook",                        "solo_redes",       9.2, "ALTA",       "Fundada 1935. Email @hotmail. Distribuidor Montblanc."],
    [3,  "Ana Esteve Estilistas",               "Peluquería",          "615 210 373",  "",                              "Avda. Tombatossals, 99, 12004 Castellón",      "",                      "",                                "sin_web",          9.0, "ALTA",       "Cero presencia digital. Sin web ni redes."],
    [4,  "Dani Vozme Hair Studio",              "Peluquería",          "622 152 595",  "",                              "Plaza Juan XXIII, 3, 12004 Castellón",         "",                      "",                                "sin_web",          8.8, "ALTA",       "Nombre premium sin presencia digital."],
    [5,  "Sandra Zahonero Estilistas",          "Peluquería / Spa",    "964 063 858",  "",                              "Plaza Real, 21, 12001 Castellón",              "",                      "Instagram: @sandrazahoneroestilistas","solo_redes",    8.7, "ALTA",       "Solo Instagram. Sin web. Perfil activo."],
    [6,  "Talleres Ultimateauto",               "Taller mecánico",     "628 678 720",  "",                              "Pol. Rda. Sur, Nave 35, 12006 Castellón",      "",                      "",                                "sin_web",          8.5, "ALTA",       "Sin web. Solo en directorios."],
    [7,  "Peluquería Julia Garrido & Julio Valls","Peluquería",        "964 231 733",  "",                              "Av. Casalduch, 6 Bajo, 12005 Castellón",       "",                      "Facebook",                        "solo_redes",       8.3, "MEDIA-ALTA", "Solo Facebook. Sin web propia."],
    [8,  "Autoescuela Sensal",                  "Autoescuela",         "964 049 338",  "",                              "C/ Cronista Revest, 13, 12005 Castellón",      "autoescolasensal.com",  "Facebook, Instagram",             "web_desactualizada",8.0,"MEDIA-ALTA", "Web sin actualizar desde dic. 2020. Diseño antiguo."],
    [9,  "VH Estilistas",                       "Peluquería",          "964 737 699",  "",                              "C/ Honori Garcia i Garcia, 11, 12006 Castellón","",                     "",                                "sin_web",          7.8, "MEDIA",      "Sin web ni redes detectadas."],
    [10, "Talleres Cardona Benages",            "Taller mecánico",     "964 226 208",  "",                              "Plaza Escultor Adsuara, 9, 12003 Castellón",   "",                      "",                                "sin_web",          7.8, "MEDIA",      "Sin web. Solo en Páginas Amarillas."],
    [11, "AK Peluqueros",                       "Peluquería",          "964 250 242",  "",                              "Av. Enrique Gimeno, 101, 12006 Castellón",     "",                      "Facebook: /akpeluqueros.es",       "solo_redes",       7.5, "MEDIA",      "Solo Facebook. Sin web."],
    [12, "Cadu Taller",                         "Taller mecánico",     "642 594 689",  "",                              "C/ Montanejos, 83 Bajo, 12005 Castellón",      "",                      "",                                "sin_web",          7.5, "MEDIA",      "Sin web ni redes."],
    [13, "Reyes Agut Peluquería",               "Peluquería",          "643 997 099",  "",                              "C/ Ronda Magdalena, 72, 12004 Castellón",      "",                      "",                                "sin_web",          7.3, "MEDIA",      "Sin web ni redes."],
    [14, "Saints Barbershop",                   "Barbería",            "964 819 935",  "",                              "C/ Falcó, 5, 12001 Castellón",                 "",                      "",                                "sin_web",          7.2, "MEDIA",      "Nombre moderno sin presencia digital."],
    [15, "Peluquería Gloria Mollón",            "Peluquería",          "964 200 030",  "",                              "Av. Dr. Clara, 20 Bajo, 12002 Castellón",      "",                      "",                                "sin_web",          7.0, "MEDIA",      "Sin web ni redes."],
    [16, "AUTODICAR S.L.",                      "Taller mecánico",     "964 280 553",  "",                              "C/ Ciudadela, 24-26, 12100 Castellón",         "",                      "",                                "sin_web",          7.0, "MEDIA",      "Sociedad limitada sin web."],
    [17, "Sweet World Heladería",               "Hostelería",          "602 385 215",  "",                              "C/ San Vicente, 3, 12002 Castellón",           "",                      "",                                "sin_web",          6.8, "MEDIA",      "Sin web. Nombre moderno no desarrollado."],
    [18, "Rubén Duat Peluquería",               "Peluquería",          "964 047 717",  "",                              "C/ Estatuto, 21, 12004 Castellón",             "",                      "",                                "sin_web",          6.7, "MEDIA-BAJA", "Sin web ni redes."],
    [19, "Talleres Lucas",                      "Taller mecánico",     "637 865 880",  "",                              "C/ Taxida, 11, 12003 Castellón",               "",                      "",                                "sin_web",          6.5, "MEDIA-BAJA", "Sin web. Solo en directorio."],
    [20, "Carrillo Taller",                     "Taller mecánico",     "964 245 658",  "",                              "Ctra. Borriol, 14, 12006 Castellón",           "",                      "",                                "sin_web",          6.3, "MEDIA-BAJA", "Sin web. Solo en Páginas Amarillas."],
  ];

  // ── CABECERA ───────────────────────────────────────────────
  const cabecera = [
    "#", "Negocio", "Sector", "Teléfono", "Email", "Dirección",
    "Web actual", "Redes sociales", "Estado digital", "Puntuación", "Prioridad", "Notas"
  ];

  sheet.getRange(1, 1, 1, cabecera.length).setValues([cabecera]);

  // Estilo cabecera
  const headerRange = sheet.getRange(1, 1, 1, cabecera.length);
  headerRange.setBackground("#1a1a2e");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setFontSize(11);
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  sheet.setRowHeight(1, 40);

  // ── FILAS DE DATOS ─────────────────────────────────────────
  sheet.getRange(2, 1, datos.length, cabecera.length).setValues(datos);

  // Altura uniforme de filas
  for (let i = 2; i <= datos.length + 1; i++) {
    sheet.setRowHeight(i, 32);
  }

  // ── FORMATO POR ESTADO DIGITAL (columna 9) ─────────────────
  const colEstado = 9;
  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    const estado = datos[i][colEstado - 1];
    const cell = sheet.getRange(fila, colEstado);
    cell.setHorizontalAlignment("center");
    cell.setFontWeight("bold");
    cell.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);

    if (estado === "sin_web") {
      cell.setBackground("#fee2e2");
      cell.setFontColor("#b91c1c");
    } else if (estado === "solo_redes") {
      cell.setBackground("#fef3c7");
      cell.setFontColor("#92400e");
    } else if (estado === "web_desactualizada") {
      cell.setBackground("#dbeafe");
      cell.setFontColor("#1e40af");
    }
  }

  // ── FORMATO PUNTUACIÓN (columna 10) ───────────────────────
  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    const puntuacion = datos[i][9];
    const cell = sheet.getRange(fila, 10);
    cell.setHorizontalAlignment("center");
    cell.setFontWeight("bold");
    cell.setFontSize(12);

    if (puntuacion >= 9.0) {
      cell.setFontColor("#b91c1c");
    } else if (puntuacion >= 8.0) {
      cell.setFontColor("#d97706");
    } else {
      cell.setFontColor("#374151");
    }
  }

  // ── FORMATO PRIORIDAD (columna 11) ────────────────────────
  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    const prioridad = datos[i][10];
    const cell = sheet.getRange(fila, 11);
    cell.setHorizontalAlignment("center");
    cell.setFontWeight("bold");

    if (prioridad === "ALTA") {
      cell.setBackground("#dc2626");
      cell.setFontColor("#ffffff");
      cell.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    } else if (prioridad === "MEDIA-ALTA") {
      cell.setBackground("#ea580c");
      cell.setFontColor("#ffffff");
      cell.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    } else if (prioridad === "MEDIA") {
      cell.setBackground("#ca8a04");
      cell.setFontColor("#ffffff");
      cell.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      cell.setBackground("#6b7280");
      cell.setFontColor("#ffffff");
      cell.setBorder(true, true, true, true, false, false, "#ffffff", SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  // ── FRANJAS ALTERNAS ──────────────────────────────────────
  for (let i = 0; i < datos.length; i++) {
    const fila = i + 2;
    if (i % 2 === 0) {
      const rowRange = sheet.getRange(fila, 1, 1, cabecera.length);
      rowRange.setBackground("#f8fafc");
    }
  }

  // ── ANCHOS DE COLUMNA ─────────────────────────────────────
  sheet.setColumnWidth(1, 40);   // #
  sheet.setColumnWidth(2, 240);  // Negocio
  sheet.setColumnWidth(3, 140);  // Sector
  sheet.setColumnWidth(4, 120);  // Teléfono
  sheet.setColumnWidth(5, 200);  // Email
  sheet.setColumnWidth(6, 270);  // Dirección
  sheet.setColumnWidth(7, 160);  // Web actual
  sheet.setColumnWidth(8, 220);  // Redes
  sheet.setColumnWidth(9, 160);  // Estado digital
  sheet.setColumnWidth(10, 100); // Puntuación
  sheet.setColumnWidth(11, 110); // Prioridad
  sheet.setColumnWidth(12, 320); // Notas

  // ── COLUMNA NEGOCIO EN NEGRITA ─────────────────────────────
  sheet.getRange(2, 2, datos.length, 1).setFontWeight("bold");

  // ── CENTRAR COLUMNA # ─────────────────────────────────────
  sheet.getRange(2, 1, datos.length, 1).setHorizontalAlignment("center");

  // ── FIJAR FILA CABECERA ───────────────────────────────────
  sheet.setFrozenRows(1);

  // ── FILTROS ───────────────────────────────────────────────
  sheet.getRange(1, 1, datos.length + 1, cabecera.length).createFilter();

  // ── SEGUNDA HOJA: RESUMEN ──────────────────────────────────
  const resumen = ss.insertSheet("Resumen");

  const resumenData = [
    ["RESUMEN DE PROSPECCIÓN"],
    [""],
    ["Zona analizada",         "Castellón de la Plana"],
    ["Servicio ofrecido",      "Diseño y renovación web"],
    ["Fecha del informe",      "26 de mayo de 2026"],
    [""],
    ["ESTADÍSTICAS"],
    ["Total negocios analizados", 20],
    ["Sin página web",            13],
    ["Solo redes sociales",        4],
    ["Web desactualizada",         2],
    ["Web aceptable",              1],
    [""],
    ["SECTORES CON MÁS OPORTUNIDAD"],
    ["Peluquerías / Estética",  "12 negocios — más del 80% sin web"],
    ["Talleres mecánicos",      "6 negocios — 100% sin web"],
    ["Fontanería",              "1 empresa clave — 60 años sin web"],
    ["Joyería",                 "1 negocio histórico sin presencia digital"],
  ];

  resumen.getRange(1, 1, resumenData.length, 2).setValues(resumenData);

  // Estilo título resumen
  resumen.getRange(1, 1).setFontSize(16).setFontWeight("bold").setFontColor("#1a1a2e");
  resumen.getRange(7, 1).setFontSize(13).setFontWeight("bold").setBackground("#e94560").setFontColor("#ffffff");
  resumen.getRange(14, 1).setFontSize(13).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");
  resumen.setColumnWidth(1, 280);
  resumen.setColumnWidth(2, 360);

  // ── TERCERA HOJA: MENSAJES DE CONTACTO ────────────────────
  const mensajes = ss.insertSheet("Mensajes de Contacto");

  const mensajesData = [
    ["Negocio", "Canal", "Mensaje listo para enviar"],
    [
      "Joyería Font",
      "Llamada / WhatsApp",
      "Hola, buenos días. Me llamo [tu nombre] y me dedico al diseño web en Castellón.\n\nHe visto que Joyería Font lleva desde 1935 aquí en la ciudad — es impresionante esa trayectoria. Buscando en Google \"joyería Castellón\" me doy cuenta de que no aparecéis con web propia, y eso es una oportunidad enorme que estáis perdiendo cada día.\n\nOs propongo crear una web elegante que muestre vuestro catálogo, vuestra historia y que aparezca en Google cuando alguien busque una joyería en Castellón.\n\n¿Tendríais 15 minutos esta semana para que os cuente cómo funciona?"
    ],
    [
      "Joyería Font",
      "Email → joyeriafontsl@hotmail.com",
      "Asunto: Vuestra web en Google — propuesta para Joyería Font\n\nBuenos días,\n\nMe llamo [tu nombre] y soy diseñador web en Castellón.\n\nHe visto que Joyería Font lleva desde 1935 en la ciudad y tiene una reputación excelente. Sin embargo, cuando busco \"joyería Castellón\" en Google, no aparecéis con una página web propia.\n\nOs propongo:\n• Una web profesional y elegante con vuestro catálogo\n• Posicionamiento en Google para \"joyería Castellón\"\n• Diseño adaptado a móviles y tablets\n• Formulario de contacto y WhatsApp directo\n\nPrecio cerrado, sin permanencia.\n\n¿Puedo enviaros una propuesta sin compromiso?\n\nUn saludo,\n[Tu nombre]\n[Tu teléfono]"
    ],
    [
      "Luis Sos Fontanería",
      "Facebook Messenger / llamada",
      "Hola, he visto que Luis Sos Fontanería lleva más de 60 años en Castellón y me ha impresionado vuestra trayectoria. Sin embargo, cuando busco \"fontanero Castellón\" en Google, no os encuentro porque no tenéis página web.\n\nOs estáis perdiendo muchos clientes nuevos que buscan online.\n\nOs propongo crear una web sencilla y profesional que os posicione en Google desde el primer mes.\n\n¿Podemos hablar 10 minutos esta semana?"
    ],
    [
      "Peluquerías (Ana Esteve, Dani Vozme, VH Estilistas, etc.)",
      "WhatsApp / Llamada",
      "Hola [nombre], soy [tu nombre], diseñador web en Castellón.\n\nHe buscado vuestra peluquería online y me he dado cuenta de que no tenéis página web. Las peluquerías que más reservas consiguen hoy en día son las que tienen web con:\n\n✅ Galería de trabajos y transformaciones\n✅ Reserva de cita online (sin llamadas)\n✅ Aparición en Google cuando buscan \"peluquería [barrio/Castellón]\"\n\nOs creo la web en 2 semanas, con diseño moderno y reserva de citas incluida.\n\n¿Os puedo llamar esta semana para contaros más?"
    ],
    [
      "Talleres mecánicos (Ultimateauto, Cardona, Cadu, etc.)",
      "WhatsApp / Llamada",
      "Hola, soy [tu nombre], diseñador web en Castellón.\n\nHe buscado vuestro taller en Google y no aparecéis con web propia. Cuando alguien busca \"taller mecánico Castellón\" desde el móvil, va al primero que encuentra.\n\nOs propongo una web sencilla con:\n✅ Vuestros servicios y precios\n✅ Cita previa online (os ahorra llamadas)\n✅ Aparición en Google Maps y búsquedas\n✅ Reseñas y valoraciones visibles\n\nPrecio fijo, sin sorpresas.\n\n¿Hablamos esta semana?"
    ],
  ];

  mensajes.getRange(1, 1, mensajesData.length, 3).setValues(mensajesData);

  const headerMensajes = mensajes.getRange(1, 1, 1, 3);
  headerMensajes.setBackground("#1a1a2e").setFontColor("#ffffff").setFontWeight("bold").setFontSize(11);
  mensajes.setColumnWidth(1, 260);
  mensajes.setColumnWidth(2, 220);
  mensajes.setColumnWidth(3, 600);
  mensajes.getRange(2, 3, mensajesData.length - 1, 1).setWrap(true);
  for (let i = 2; i <= mensajesData.length; i++) {
    mensajes.setRowHeight(i, 160);
  }
  mensajes.setFrozenRows(1);

  // ── ABRE EL SHEET EN EL NAVEGADOR ─────────────────────────
  const url = ss.getUrl();
  Logger.log("✅ Google Sheet creado: " + url);

  // Muestra la URL en un popup
  SpreadsheetApp.getUi ? null : null;
  const ui = SpreadsheetApp.getUi ? null : null;
  Logger.log("Abre este enlace: " + url);

  // Activa la primera hoja
  ss.setActiveSheet(sheet);

  Browser.msgBox(
    "✅ Sheet creado correctamente",
    "Tu Google Sheet está listo.\n\nURL:\n" + url,
    Browser.Buttons.OK
  );
}
