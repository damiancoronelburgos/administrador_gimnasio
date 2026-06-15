const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, Header, Footer
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    children: [new TextRun({ text, bold: true, size: 36, font: "Arial", color: "1B4F72" })]
  });
}

function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, font: "Arial", color: "2E86C1" })]
  });
}

function h3(text) {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, font: "Arial", color: "1A5276" })]
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...opts })]
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function tableRow(cells, isHeader = false) {
  return new TableRow({
    children: cells.map((text, i) => new TableCell({
      borders,
      margins: cellMargins,
      width: { size: Math.floor(9360 / cells.length), type: WidthType.DXA },
      shading: isHeader ? { fill: "1B4F72", type: ShadingType.CLEAR } : (i % 2 === 0 ? { fill: "EBF5FB", type: ShadingType.CLEAR } : { fill: "FFFFFF", type: ShadingType.CLEAR }),
      children: [new Paragraph({
        children: [new TextRun({ text, size: 20, font: "Arial", bold: isHeader, color: isHeader ? "FFFFFF" : "000000" })]
      })]
    }))
  });
}

function makeTable(headers, rows) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: headers.map(() => Math.floor(9360 / headers.length)),
    rows: [
      tableRow(headers, true),
      ...rows.map(r => tableRow(r))
    ]
  });
}

function spacer() {
  return new Paragraph({ spacing: { after: 160 }, children: [] });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } }
        }]
      }
    ]
  },
  styles: {
    default: {
      document: { run: { font: "Arial", size: 22 } }
    }
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1B4F72", space: 1 } },
          children: [new TextRun({ text: "Proyecto: Administrador de Gimnasio", size: 18, font: "Arial", color: "1B4F72" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "1B4F72", space: 1 } },
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Pagina ", size: 18, font: "Arial", color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "Arial", color: "888888" }),
            new TextRun({ text: " de ", size: 18, font: "Arial", color: "888888" }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18, font: "Arial", color: "888888" }),
          ]
        })]
      })
    },
    children: [

      // PORTADA
      spacer(), spacer(), spacer(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "ADMINISTRADOR DE GIMNASIO", bold: true, size: 56, font: "Arial", color: "1B4F72" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: "Plan de Proyecto — App Mobile + Backend", size: 28, font: "Arial", color: "2E86C1" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 800 },
        children: [new TextRun({ text: "Proyecto de aprendizaje basado en la arquitectura de Changayaa", size: 22, font: "Arial", color: "888888", italics: true })]
      }),
      spacer(), spacer(), spacer(),

      // SECCION 1: OBJETIVO
      h1("1. Objetivo del Proyecto"),
      p("Construir una aplicacion mobile para administrar un gimnasio que cubra los mismos conceptos tecnicos de Changayaa pero con una logica de negocio simple y comprensible. El resultado sera una APK funcional para Android lista para portafolio."),
      spacer(),
      p("Al terminar este proyecto vas a entender como funcionan en conjunto:", { bold: true }),
      bullet("React + Ionic para construir interfaces mobile"),
      bullet("Capacitor para convertir una web app en una APK nativa de Android"),
      bullet("Node.js + Express para construir una API REST"),
      bullet("MySQL + Sequelize para modelar y persistir datos"),
      bullet("JWT para autenticacion y sesiones"),
      bullet("Multer para subida de archivos (fotos de socios)"),
      bullet("Socket.IO para notificaciones en tiempo real"),
      bullet("Docker para correr todo el entorno local"),
      spacer(),

      // SECCION 2: ARQUITECTURA
      h1("2. Arquitectura del Proyecto"),
      p("La arquitectura replica exactamente la estructura de Changayaa, simplificando el dominio. Tres servicios: frontend mobile, backend API, y base de datos, todos orquestados con Docker."),
      spacer(),
      makeTable(
        ["Carpeta", "Tecnologia", "Descripcion"],
        [
          ["frontend/", "React + Ionic + Capacitor", "App mobile — genera la APK de Android"],
          ["backend/", "Node.js + Express + Sequelize", "API REST — logica de negocio"],
          ["frontend/android/", "Gradle + Capacitor", "Proyecto nativo generado automaticamente"],
          [".docker/", "Docker Compose", "Orquesta todos los servicios"],
        ]
      ),
      spacer(),

      // SECCION 3: FUNCIONALIDADES
      h1("3. Funcionalidades de la App"),

      h2("3.1 Modulo de Autenticacion"),
      bullet("Login con email y contrasenia — retorna JWT"),
      bullet("Registro de administrador (unico usuario del sistema)"),
      bullet("Persistencia de sesion en localStorage"),
      bullet("Cierre de sesion"),
      spacer(),

      h2("3.2 Modulo de Socios"),
      bullet("Listar todos los socios con foto y estado (activo / vencido)"),
      bullet("Agregar nuevo socio (nombre, apellido, DNI, telefono, foto)"),
      bullet("Editar datos de un socio"),
      bullet("Eliminar un socio"),
      bullet("Subir foto con camara del celular via Capacitor Camera"),
      spacer(),

      h2("3.3 Modulo de Membresías"),
      bullet("Registrar pago de membresia mensual para un socio"),
      bullet("Ver historial de pagos por socio"),
      bullet("Calcular automaticamente fecha de vencimiento (30 dias)"),
      bullet("Marcar socios como vencidos cuando expira la membresia"),
      spacer(),

      h2("3.4 Modulo de Asistencia"),
      bullet("Registrar entrada de un socio al gimnasio"),
      bullet("Ver historial de asistencia del dia actual"),
      bullet("Ver historial de asistencia por socio"),
      spacer(),

      h2("3.5 Notificaciones en Tiempo Real (Socket.IO)"),
      bullet("Notificar cuando un socio tiene la membresia proxima a vencer (menos de 5 dias)"),
      bullet("Panel de notificaciones en la app"),
      spacer(),

      h2("3.6 Dashboard"),
      bullet("Total de socios activos vs vencidos"),
      bullet("Socios que asistieron hoy"),
      bullet("Proximos vencimientos (proximos 7 dias)"),
      spacer(),

      // SECCION 4: MODELO DE DATOS
      h1("4. Modelo de Base de Datos"),
      p("La base de datos es MySQL manejada con Sequelize. Cuatro tablas principales que replican los patrones de Changayaa (modelos, asociaciones, relaciones has-many/belongs-to)."),
      spacer(),

      h3("Tabla: users"),
      makeTable(
        ["Campo", "Tipo", "Descripcion"],
        [
          ["id", "INT (PK)", "Identificador unico"],
          ["name", "VARCHAR(100)", "Nombre del administrador"],
          ["email", "VARCHAR(100)", "Email — usado para login"],
          ["password", "VARCHAR(255)", "Hash bcrypt de la contrasenia"],
          ["created_at", "DATETIME", "Fecha de creacion"],
        ]
      ),
      spacer(),

      h3("Tabla: members (socios)"),
      makeTable(
        ["Campo", "Tipo", "Descripcion"],
        [
          ["id", "INT (PK)", "Identificador unico"],
          ["name", "VARCHAR(100)", "Nombre del socio"],
          ["last_name", "VARCHAR(100)", "Apellido del socio"],
          ["dni", "VARCHAR(20)", "Documento de identidad"],
          ["phone", "VARCHAR(20)", "Telefono de contacto"],
          ["photo_url", "VARCHAR(255)", "Ruta a la foto subida con Multer"],
          ["status", "ENUM('active','expired')", "Estado de la membresia"],
          ["created_at", "DATETIME", "Fecha de alta"],
        ]
      ),
      spacer(),

      h3("Tabla: memberships (membresias)"),
      makeTable(
        ["Campo", "Tipo", "Descripcion"],
        [
          ["id", "INT (PK)", "Identificador unico"],
          ["member_id", "INT (FK)", "Referencia al socio"],
          ["start_date", "DATE", "Fecha de inicio de la membresia"],
          ["end_date", "DATE", "Fecha de vencimiento (start + 30 dias)"],
          ["amount_paid", "DECIMAL(10,2)", "Monto abonado"],
          ["created_at", "DATETIME", "Fecha de pago"],
        ]
      ),
      spacer(),

      h3("Tabla: attendances (asistencias)"),
      makeTable(
        ["Campo", "Tipo", "Descripcion"],
        [
          ["id", "INT (PK)", "Identificador unico"],
          ["member_id", "INT (FK)", "Referencia al socio"],
          ["checked_in_at", "DATETIME", "Fecha y hora de entrada"],
        ]
      ),
      spacer(),

      // SECCION 5: API ENDPOINTS
      h1("5. Endpoints de la API"),

      h2("Autenticacion"),
      makeTable(
        ["Metodo", "Endpoint", "Descripcion"],
        [
          ["POST", "/api/auth/login", "Login — retorna JWT"],
          ["POST", "/api/auth/register", "Registro del administrador"],
        ]
      ),
      spacer(),

      h2("Socios"),
      makeTable(
        ["Metodo", "Endpoint", "Descripcion"],
        [
          ["GET", "/api/members", "Listar todos los socios"],
          ["GET", "/api/members/:id", "Obtener un socio por ID"],
          ["POST", "/api/members", "Crear nuevo socio (con foto via Multer)"],
          ["PUT", "/api/members/:id", "Editar datos de un socio"],
          ["DELETE", "/api/members/:id", "Eliminar un socio"],
        ]
      ),
      spacer(),

      h2("Membresias"),
      makeTable(
        ["Metodo", "Endpoint", "Descripcion"],
        [
          ["GET", "/api/memberships/:memberId", "Historial de pagos de un socio"],
          ["POST", "/api/memberships", "Registrar nuevo pago de membresia"],
        ]
      ),
      spacer(),

      h2("Asistencias"),
      makeTable(
        ["Metodo", "Endpoint", "Descripcion"],
        [
          ["GET", "/api/attendance/today", "Asistencias del dia actual"],
          ["GET", "/api/attendance/:memberId", "Historial de asistencia por socio"],
          ["POST", "/api/attendance", "Registrar entrada de un socio"],
        ]
      ),
      spacer(),

      h2("Dashboard"),
      makeTable(
        ["Metodo", "Endpoint", "Descripcion"],
        [
          ["GET", "/api/dashboard/stats", "Estadisticas generales (socios activos, vencidos, hoy)"],
          ["GET", "/api/dashboard/expiring", "Socios que vencen en los proximos 7 dias"],
        ]
      ),
      spacer(),

      // SECCION 6: PANTALLAS
      h1("6. Pantallas de la App (Frontend)"),
      makeTable(
        ["Pantalla", "Ruta", "Descripcion"],
        [
          ["Login", "/login", "Formulario de email y contrasenia"],
          ["Dashboard", "/dashboard", "Estadisticas y proximos vencimientos"],
          ["Lista de Socios", "/members", "Grid con foto, nombre y badge de estado"],
          ["Detalle de Socio", "/members/:id", "Info completa, membresia activa, asistencia reciente"],
          ["Nuevo/Editar Socio", "/members/form", "Formulario con captura de foto via camara"],
          ["Registrar Pago", "/memberships/new", "Seleccionar socio y registrar pago"],
          ["Asistencia", "/attendance", "Lista del dia + boton para registrar entrada"],
          ["Notificaciones", "/notifications", "Panel de socios proximos a vencer"],
        ]
      ),
      spacer(),

      // SECCION 7: ESTRUCTURA DE CARPETAS
      h1("7. Estructura de Carpetas"),
      p("Replica la misma convencion que Changayaa para que la comparacion sea directa:"),
      spacer(),

      h2("Backend"),
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({
          text: [
            "backend/",
            "  app.js                  (entry point — Express + Socket.IO)",
            "  routers/",
            "    authRoutes.js",
            "    memberRoutes.js",
            "    membershipRoutes.js",
            "    attendanceRoutes.js",
            "    dashboardRoutes.js",
            "  controllers/",
            "    authController.js",
            "    memberController.js",
            "    membershipController.js",
            "    attendanceController.js",
            "    dashboardController.js",
            "  sequelize/",
            "    models/",
            "      User.js",
            "      Member.js",
            "      Membership.js",
            "      Attendance.js",
            "    config/",
            "      database.js",
            "  middleware/",
            "    authMiddleware.js     (verifica JWT)",
            "  multer/",
            "    memberPhoto.js        (config de subida de foto)",
            "  uploads/               (fotos guardadas aqui)",
          ].join("\n"),
          font: "Courier New", size: 18, color: "1A5276"
        })]
      }),
      spacer(),

      h2("Frontend"),
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({
          text: [
            "frontend/",
            "  src/",
            "    pages/",
            "      Login/Login.tsx",
            "      Dashboard/Dashboard.tsx",
            "      Members/",
            "        MemberList.tsx",
            "        MemberDetail.tsx",
            "        MemberForm.tsx",
            "      Memberships/MembershipForm.tsx",
            "      Attendance/AttendancePage.tsx",
            "      Notifications/NotificationsPage.tsx",
            "    components/",
            "      MemberCard/MemberCard.tsx",
            "      Navbar/Navbar.tsx",
            "      StatusBadge/StatusBadge.tsx",
            "    contexts/",
            "      AuthContext.tsx     (user, login, logout)",
            "    services/",
            "      api.ts              (axios instance con JWT)",
            "      memberService.ts",
            "      membershipService.ts",
            "    router/",
            "      AppRouter.tsx       (rutas protegidas)",
            "  capacitor.config.ts",
            "  ionic.config.json",
          ].join("\n"),
          font: "Courier New", size: 18, color: "1A5276"
        })]
      }),
      spacer(),

      // SECCION 8: ROADMAP
      h1("8. Roadmap de Desarrollo"),
      p("Orden sugerido para construir el proyecto paso a paso, de menor a mayor complejidad:"),
      spacer(),
      makeTable(
        ["Fase", "Que construir", "Conceptos que se aprenden"],
        [
          ["Fase 1", "Setup: Docker + MySQL + Express arrancando", "Docker Compose, variables de entorno, conexion a DB"],
          ["Fase 2", "Modelos Sequelize + migraciones", "ORM, modelos, asociaciones, sync"],
          ["Fase 3", "Auth: register + login + JWT", "bcrypt, JWT, middleware de autenticacion"],
          ["Fase 4", "CRUD de socios + Multer para fotos", "REST API completa, subida de archivos"],
          ["Fase 5", "Membresias y calculo de vencimiento", "Logica de negocio, fechas, relaciones FK"],
          ["Fase 6", "Asistencias + Dashboard stats", "Queries con Sequelize, agregaciones"],
          ["Fase 7", "Setup frontend: Ionic + React + Capacitor", "Ionic pages, routing, AuthContext"],
          ["Fase 8", "Pantallas de socios + consumo de API", "fetch/axios, formularios, Ionic components"],
          ["Fase 9", "Camara nativa con Capacitor Camera", "Plugins nativos de Capacitor"],
          ["Fase 10", "Socket.IO: notificaciones de vencimiento", "WebSockets, eventos del servidor al cliente"],
          ["Fase 11", "Build APK con Capacitor + Gradle", "npx cap sync, Android Studio, firma del APK"],
        ]
      ),
      spacer(),

      // SECCION 9: COMPARACION CON CHANGAYAA
      h1("9. Paralelo con Changayaa"),
      p("Esta tabla muestra como cada concepto de este proyecto se mapea directamente a lo que hace Changayaa en produccion:"),
      spacer(),
      makeTable(
        ["Este proyecto", "Changayaa (equivalente)"],
        [
          ["Socio con foto (Member)", "Usuario con avatar (User)"],
          ["Pago de membresia (Membership)", "Facturacion y suscripciones (Billing, Subscription)"],
          ["Registro de asistencia (Attendance)", "Registro de contratos y tareas (Contract, Task)"],
          ["JWT en AuthContext", "JWT en UserContext"],
          ["Multer para foto de socio", "Multer para avatar, documentos, imagenes de trabajo"],
          ["Socket.IO para vencimientos", "Socket.IO para chat y notificaciones"],
          ["Dashboard de stats", "Panel de admin con metricas"],
          ["Capacitor Camera", "Capacitor Camera (mismo plugin)"],
          ["Build APK Capacitor", "Build APK Capacitor (mismo flujo)"],
          ["Docker Compose local", "Docker Compose en produccion"],
        ]
      ),
      spacer(),

      // SECCION 10: STACK TECNOLOGICO
      h1("10. Stack Tecnologico Completo"),
      makeTable(
        ["Capa", "Tecnologia", "Version sugerida"],
        [
          ["Runtime JS", "Node.js", "20.x LTS"],
          ["Framework backend", "Express", "4.x"],
          ["ORM", "Sequelize", "6.x"],
          ["Base de datos", "MySQL", "8.0"],
          ["Autenticacion", "jsonwebtoken + bcrypt", "9.x / 5.x"],
          ["Subida de archivos", "Multer", "1.x"],
          ["Tiempo real", "Socket.IO", "4.x"],
          ["Framework frontend", "React", "18.x"],
          ["Framework mobile", "Ionic", "7.x"],
          ["Bridge nativo", "Capacitor", "6.x"],
          ["Lenguaje frontend", "TypeScript", "5.x"],
          ["Build frontend", "Vite", "5.x"],
          ["HTTP Client", "Axios", "1.x"],
          ["Contenedores", "Docker + Docker Compose", "latest"],
        ]
      ),
      spacer(),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("C:\\Users\\damian\\Documents\\GitHub\\administrador_gimnasio\\plan_proyecto.docx", buffer);
  console.log("Documento creado exitosamente.");
});
