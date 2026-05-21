/**
 * Market Dashboard / Next.js Research – Report & Presentation generator
 * Run:  node generate-docs.cjs
 */
"use strict";

const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, PageBreak, TableOfContents,
  BorderStyle, Table, TableRow, TableCell,
  WidthType, Header, Footer,
  PageNumber, LevelFormat, convertInchesToTwip, LineRuleType,
} = require("./node_modules/docx/dist/index.umd.cjs");

const pptxgen = require("./node_modules/pptxgenjs/dist/pptxgen.cjs.js");
const { writeFileSync, mkdirSync } = require("fs");

mkdirSync("docs", { recursive: true });

// ── Metadata ──────────────────────────────────────────────────────────────

const TITLE    = "Next.js: Research and Development of a Web Application";
const COURSE   = "Advanced Web Programming";
const YEAR     = "2025/2026 - 1st Semester";
const UNIV     = "ISLA Gaia";
const DEPT     = "Information Systems and Technologies";
const PROF     = "Jose Joaquim Magalhaes Moreira";
const STUDENTS = [
  { name: "Raphael Malburg", id: "a22500074" },
  { name: "Andre Neves",     id: "a22508912" },
];

// ── DOCX helpers ──────────────────────────────────────────────────────────

const FONT    = "Times New Roman";
const FS      = 24;   // 12 pt in half-points
const SPACING = { after: 200, line: 276, lineRule: LineRuleType.AUTO };

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: FS, ...opts });
}
function para(text, paraOpts = {}) {
  const child = typeof text === "string" ? run(text) : text;
  return new Paragraph({ spacing: SPACING, alignment: AlignmentType.JUSTIFIED, children: [child], ...paraOpts });
}
function paraRuns(runs, paraOpts = {}) {
  return new Paragraph({ spacing: SPACING, alignment: AlignmentType.JUSTIFIED, children: runs, ...paraOpts });
}
function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 480, after: 240 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true })] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true })] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FS, bold: true, italics: true })] });
}
function bullet(text, level = 0) {
  return new Paragraph({ bullet: { level }, spacing: { after: 100, line: 276, lineRule: LineRuleType.AUTO },
    alignment: AlignmentType.JUSTIFIED, children: [run(text)] });
}
function center(text, opts = {}) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 },
    children: [new TextRun({ text, font: FONT, size: FS, ...opts })] });
}
function empty() {
  return new Paragraph({ spacing: { after: 160 }, children: [run("")] });
}
function br() {
  return new Paragraph({ children: [new PageBreak()] });
}
function bibEntry(text) {
  return new Paragraph({ indent: { left: 720, hanging: 720 }, spacing: SPACING,
    alignment: AlignmentType.JUSTIFIED, children: [run(text)] });
}
function captionPara(text) {
  return new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new TextRun({ text, font: FONT, size: 20, italics: true })] });
}

// ── DOCX Content ──────────────────────────────────────────────────────────

const COVER = [
  empty(), empty(), empty(),
  center(UNIV,    { bold: true, size: 28 }),
  center(DEPT),
  empty(), empty(),
  center(TITLE,   { bold: true, size: 32 }),
  empty(),
  center("Practical Assignment – " + COURSE, { italics: true, size: 26 }),
  empty(), empty(),
  ...STUDENTS.map(s => center(`${s.name}  |  ${s.id}`)),
  empty(),
  center(`Lecturer: ${PROF}`),
  empty(), empty(),
  center(YEAR),
  center("May 2026"),
  br(),
];

const ABSTRACT = [
  h1("Abstract"),
  para("This work, developed within the scope of the Advanced Web Programming course, aims to investigate the Next.js framework created by Vercel, analysing its state of the art, advantages, disadvantages and main areas of use. Alongside the research, a demonstration web application called Market Dashboard was developed using the same framework, illustrating its capabilities in the context of real-time financial data visualisation. The application integrates the unofficial Yahoo Finance API, presents candlestick charts with timeframe selection, and supports configurable technical indicators, serving as a proof of concept for Next.js features in a production-like environment."),
  empty(),
  paraRuns([run("Keywords: ", { bold: true }), run("Next.js, React, web framework, hybrid rendering, web application, financial visualisation.")]),
  br(),
];

const INDEX = [
  h1("Table of Contents"),
  para("Note: In Microsoft Word, right-click the table of contents and select \"Update Field\" to generate page numbers.", { italics: true }),
  new Paragraph({
    spacing: { after: 200 },
    children: [new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" })],
  }),
  br(),
];

const FIGURES = [
  h1("List of Figures and Tables"),
  empty(),
  paraRuns([run("Figures", { bold: true })]),
  para("Figure 1 – Next.js hybrid rendering architecture ........................... 5"),
  para("Figure 2 – Rendering approach comparison (CSR vs SSR vs SSG) .............. 6"),
  para("Figure 3 – Market Dashboard application folder structure ................... 9"),
  para("Figure 4 – Market dashboard interface ...................................... 10"),
  para("Figure 5 – Candlestick chart with technical indicators ..................... 11"),
  empty(),
  paraRuns([run("Tables", { bold: true })]),
  para("Table 1 – Framework comparison: Next.js vs alternatives .................... 7"),
  para("Table 2 – Technologies used in the implementation .......................... 9"),
  para("Table 3 – Internal API endpoints ........................................... 10"),
  br(),
];

const INTRO = [
  h1("1. Introduction"),
  para("Web application development has evolved at an accelerated pace over the last two decades, driven by increasing demands for fast, accessible, and dynamic user experiences. In this context, web frameworks emerge as fundamental tools for organising and abstracting development complexity, allowing teams to focus on business logic instead of solving infrastructure problems repeatedly."),
  para("This work focuses on Next.js, an open-source React framework created by Vercel in 2016, which has distinguished itself through its hybrid rendering model, capable of combining Server-Side Rendering (SSR), Static Site Generation (SSG), and Client-Side Rendering (CSR) within a single application. This model gives Next.js unique flexibility in the landscape of modern web frameworks."),
  h2("1.1 Objectives"),
  bullet("Investigate the state of the art of Next.js, including its history, architecture, and ecosystem;"),
  bullet("Identify the advantages, disadvantages, and main areas of use of the framework;"),
  bullet("Evaluate the development experience with Next.js and its practical benefits;"),
  bullet("Develop a demonstration application – Market Dashboard – using the investigated framework."),
  h2("1.2 Report Structure"),
  para("The report is organised as follows: Section 2 presents the state of the art of Next.js; Section 3 describes the problem-solving methodology; Section 4 details the application implementation; Section 5 presents conclusions, difficulties and proposed improvements; Section 6 lists the bibliographic references."),
  br(),
];

const SOA = [
  h1("2. State of the Art – Next.js"),
  h2("2.1 History and Evolution"),
  para("Next.js was publicly released in October 2016 by Guillermo Rauch and the Vercel team (then called Zeit), as a response to a concrete problem: building React applications with server-side rendering in a simple and uncomplicated way. The first version introduced six core principles: server rendering by default, route-level code loading, CSS support, hot-reloading in development, a simple build cycle, and minimal configuration (Vercel, Inc., 2026)."),
  para("Since then, the framework has evolved consistently. Version 9 (2019) introduced automatic dynamic routing; version 10 (2020) brought integrated image optimisation; version 12 (2021) introduced the Rust-based SWC compiler and edge-level middleware; version 13 (2022) launched the App Router based on React Server Components; version 14 (2023) stabilised the App Router and introduced Server Actions; versions 15 and beyond continued to refine the component model and performance (Vercel, Inc., 2026)."),
  h2("2.2 Architecture and Main Features"),
  para("Next.js is built on React and extends it with a set of features covering both the frontend and backend of a web application:"),
  h3("2.2.1 Rendering Models"),
  para("One of Next.js's most distinctive characteristics is its ability to combine different rendering strategies within the same application:"),
  bullet("Server-Side Rendering (SSR): the page is generated on the server per request, ensuring always up-to-date data;"),
  bullet("Static Site Generation (SSG): pages are pre-generated at build time, offering maximum performance;"),
  bullet("Incremental Static Regeneration (ISR): allows updating static pages in the background without recompiling the entire application;"),
  bullet("Client-Side Rendering (CSR): traditional browser-side rendering, suitable for highly interactive components."),
  h3("2.2.2 App Router and React Server Components"),
  para("From version 13 onwards, Next.js introduced the App Router, based on React Server Components (RSC). RSC allow components to run exclusively on the server, reducing the JavaScript bundle sent to the client. This architecture clearly distinguishes between server components (no interactivity, no browser access) and client components (marked with the 'use client' directive), which hydrate in the browser and support state and events."),
  h3("2.2.3 API Routes and Full-Stack"),
  para("Next.js includes native support for API Routes, enabling REST endpoints to be created directly within the application structure, without the need for a separate server. This feature positions the framework as a full-stack solution, integrating frontend and backend into a single project."),
  h3("2.2.4 Built-in Optimisations"),
  para("The framework includes a set of automatic optimisations: image optimisation (next/image), optimised font loading (next/font), automatic route prefetching, per-route code splitting, and Edge Runtime support for functions close to the user."),
  h2("2.3 Ecosystem and Adoption"),
  para("Next.js is currently one of the most widely adopted web frameworks worldwide. According to the annual State of JavaScript surveys (2023), Next.js is consistently ranked as the most used React framework in production environments, with a high satisfaction rate among developers. Companies such as Netflix, TikTok, Twitch, Hulu, and GitHub use Next.js in their platforms (Vercel, Inc., 2026)."),
  para("The Next.js ecosystem benefits from the vast React community, as well as native integration with the Vercel deployment platform, which offers features such as automatic PR previews, integrated analytics, and global Edge Network support."),
  h2("2.4 Advantages"),
  bullet("Hybrid rendering: ability to choose the most appropriate rendering strategy per route or component;"),
  bullet("Developer Experience (DX): minimal configuration, hot-reloading, integrated TypeScript, and detailed error messages;"),
  bullet("Performance: automatic code splitting, route prefetching, image and font optimisation;"),
  bullet("Full-stack in a single codebase: API Routes and Server Actions eliminate the need for a separate backend for common use cases;"),
  bullet("SEO: SSR and SSG ensure content is indexable by search engines, unlike pure SPA applications;"),
  bullet("Community and ecosystem: large community, extensive documentation, and integration with the React ecosystem;"),
  bullet("Scalability: suitable for both small projects and large-scale production applications."),
  h2("2.5 Disadvantages"),
  bullet("Learning curve: the App Router model with React Server Components introduces new concepts requiring deep understanding;"),
  bullet("Growing complexity: the separation between server and client components can make architecture difficult to manage in large projects;"),
  bullet("Vercel dependency: although open-source, Next.js is developed by Vercel, whose commercial interests may influence the framework's direction;"),
  bullet("Cold starts in serverless: serverless functions may experience initial latency on platforms that do not maintain warm instances;"),
  bullet("Build times: projects with many static pages can have high compilation times;"),
  bullet("Potential vendor lock-in: some advanced features (such as optimised Edge Middleware) work best on Vercel's infrastructure."),
  h2("2.6 Main Areas of Use"),
  bullet("E-commerce platforms (Shopify Hydrogen, Nike.com): benefit from SSG and ISR for high-performance product pages;"),
  bullet("SaaS applications and dashboards: the App Router and Server Components simplify building complex interfaces with dynamic data;"),
  bullet("Content and media portals (blogs, online newspapers): SSG generates static pages with maximum performance;"),
  bullet("Full-stack APIs and applications: API Routes allow building simple backends without additional infrastructure;"),
  bullet("Financial and real-time data applications: rendering flexibility and support for incremental updates make Next.js suitable for live data dashboards."),
  h2("2.7 Comparison with Alternatives"),
  para("Table 1 compares Next.js with alternative frameworks in the same market space:"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Framework", "Base", "Rendering", "Full-Stack", "Strength"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["Next.js",   "React",  "SSR/SSG/CSR/ISR", "Yes (API Routes)", "Flexibility and DX"],
        ["Nuxt.js",   "Vue 3",  "SSR/SSG/CSR",     "Yes",              "Vue ecosystem"],
        ["SvelteKit", "Svelte", "SSR/SSG/CSR",     "Yes",              "Performance / small bundle"],
        ["Remix",     "React",  "SSR",             "Yes",              "Web standards / forms"],
        ["Astro",     "Multi",  "SSG/MPA",         "Partial",          "Static content sites"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Table 1 – Framework comparison: Next.js vs alternatives"),
  empty(),
  br(),
];

const METHOD = [
  h1("3. Problem-Solving Methodology"),
  h2("3.1 Research Approach"),
  para("The investigation into Next.js was conducted based on primary sources (official framework documentation, version changelogs, GitHub repository) and secondary sources (academic articles, technical publications, community surveys). This analysis enabled the structuring of the state of the art presented in Section 2 and the identification of the most relevant use cases for the framework."),
  h2("3.2 Selection of the Demonstration Application"),
  para("To demonstrate Next.js capabilities in a real context, an application called Market Dashboard was developed — a real-time financial data visualisation platform. This choice was motivated by the need to exercise key framework features:"),
  bullet("API Routes (full-stack): implementation of REST endpoints to access external data;"),
  bullet("App Router with server and client components: efficient management of hybrid rendering;"),
  bullet("ISR and data caching: cache invalidation policies for data with different update frequencies;"),
  bullet("Interactive client components: charts, filters, and real-time updates."),
  h2("3.3 Development Process"),
  para("Development followed an iterative approach divided into three phases:"),
  bullet("Phase 1 – Base dashboard: asset listing page with updated prices via Yahoo Finance API, with fallback to simulated data on failure;"),
  bullet("Phase 2 – Detail pages: dynamic routes /market/[ticker] with per-asset metrics (open, high, low, volume);"),
  bullet("Phase 3 – Advanced visualisation: candlestick chart with timeframe selection, infinite scroll for historical data, 15-second polling and configurable technical indicators."),
  br(),
];

const IMPL = [
  h1("4. Description of the Implementation"),
  h2("4.1 Technologies Used"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Technology", "Version", "Role in the Application"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["Next.js",            "16.2.6",      "Main framework (App Router, SSR, API Routes)"],
        ["React",              "19.2.4",      "UI library"],
        ["TypeScript",         "5.x",         "Static typing"],
        ["Tailwind CSS",       "4.x",         "Utility-first styling"],
        ["lightweight-charts", "5.2.0",       "Financial charts (TradingView)"],
        ["Yahoo Finance API",  "Unofficial",  "Market data source"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Table 2 – Technologies used in the implementation"),
  empty(),
  h2("4.2 Application Architecture"),
  para("The application makes use of the App Router in Next.js 16, which distinguishes between server components and client components. Server components execute on the server and produce HTML without sending JavaScript to the browser; client components (marked with 'use client') hydrate in the browser and support interactivity."),
  para("The route structure is as follows:"),
  bullet("app/market/page.tsx – server component; asset listing page;"),
  bullet("app/market/[ticker]/page.tsx – server component; asset detail page;"),
  bullet("app/api/stocks/route.ts – API Route; current prices (30-second ISR cache);"),
  bullet("app/api/ohlcv/[ticker]/route.ts – API Route; historical OHLCV data;"),
  bullet("components/candlestick-chart.tsx – wrapper using next/dynamic ssr:false;"),
  bullet("components/candlestick-chart-inner.tsx – client component; interactive chart."),
  h2("4.3 API Layer"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Endpoint", "Parameters", "Description"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["GET /api/stocks",         "–",                "Current prices for 6 assets (30-second ISR); simulated fallback"],
        ["GET /api/ohlcv/[ticker]", "interval, before", "~400 OHLCV candles from Yahoo Finance; pagination via 'before'"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Table 3 – Internal API endpoints"),
  empty(),
  h2("4.4 Candlestick Chart and Indicators"),
  para("The CandlestickChartInner component uses TradingView's lightweight-charts v5 library and implements four behaviours:"),
  bullet("Initial loading: ~400 candles per timeframe (1m, 5m, 15m, 1H, 1D, 1W);"),
  bullet("Infinite scroll: detected via subscribeVisibleLogicalRangeChange; loads older data when reaching the left edge;"),
  bullet("15-second polling: updates the current candle and appends new candles; guarded by pollInFlightRef to prevent concurrency;"),
  bullet("Technical indicators: MA 20, MA 50, and Bollinger Bands (20 periods, sigma=2), calculated on the client and drawn as overlaid line series."),
  h2("4.5 Highlighted Next.js Features Used"),
  bullet("App Router with Server and Client Components: clear separation of responsibilities and bundle optimisation;"),
  bullet("API Routes as proxy: avoids exposing external calls to the client and centralises error handling;"),
  bullet("ISR (revalidate: 30): price response caching with automatic invalidation every 30 seconds;"),
  bullet("next/dynamic with ssr: false: dynamic import for DOM-accessing libraries (lightweight-charts);"),
  bullet("generateMetadata: dynamic per-page metadata for SEO."),
  br(),
];

const CONCLUSIONS = [
  h1("5. Conclusions"),
  h2("5.1 Difficulties Encountered"),
  bullet("Yahoo Finance API: unofficial and undocumented API, with null fields in OHLCV series and risk of unavailability; resolved with a simulated data fallback;"),
  bullet("SSR vs client components: lightweight-charts accesses the DOM on import, causing hydration errors; resolved with next/dynamic and ssr: false;"),
  bullet("Migration to lightweight-charts v5: API changed (addCandlestickSeries removed, replaced by addSeries with series definition as argument); identified through source code analysis;"),
  bullet("Asynchronous state management: coexistence of polling, scroll pagination and timeframe switching required the use of mutable refs (useRef) and a pollInFlightRef guard to avoid race conditions."),
  h2("5.2 Limitations"),
  bullet("Unofficial API dependency: risk of breaking without notice;"),
  bullet("Fixed asset set: limited to six predefined tickers;"),
  bullet("Polling vs real-time: up to 15 seconds of latency compared to WebSocket;"),
  bullet("No persistence: no authentication, profiles, or portfolio management;"),
  bullet("Limited historical data by interval: the API restricts 1-minute data to the last 7 days."),
  h2("5.3 Alternatives and Improvements"),
  bullet("WebSocket or Server-Sent Events for genuinely real-time updates;"),
  bullet("Official API (Polygon.io, Finnhub) for greater reliability and more extensive historical data;"),
  bullet("Support for searching any ticker and dynamically adding assets;"),
  bullet("Additional indicators: RSI, MACD, Stochastic;"),
  bullet("Authentication and portfolio management with price alerts;"),
  bullet("Automated testing (Vitest + Playwright) to ensure stability."),
  h2("5.4 Feedback on Next.js"),
  para("The development experience with Next.js proved to be very positive. The separation between server and client components, although initially confusing, represents a powerful abstraction that enables significant performance optimisations without sacrificing productivity. The full-stack approach, with integrated API Routes, substantially simplified the project architecture, eliminating the need for a separate backend server."),
  para("The official documentation is extensive and well organised, and the development experience in development mode (hot-reloading, detailed error messages) is remarkably smooth. On the other hand, the App Router with React Server Components has a considerable learning curve, and the boundary between what can and cannot be done in a server component is not always immediately obvious."),
  para("In summary, Next.js proves to be a solid choice for modern web applications that require performance, SEO, and scalability, with a very favourable balance between ease of use and technical capability."),
  br(),
];

const BIBLIOGRAPHY = [
  h1("6. Bibliography"),
  para("References formatted according to the APA 7th edition standard."),
  empty(),
  bibEntry("Fielding, R. T. (2000). Architectural styles and the design of network-based software architectures [Doctoral dissertation, University of California, Irvine]. https://roy.gbiv.com/pubs/dissertation/top.htm"),
  bibEntry("Meta Platforms, Inc. (2026). React: A JavaScript library for building user interfaces. https://react.dev"),
  bibEntry("Microsoft. (2012-2026). TypeScript: JavaScript with syntax for types. https://www.typescriptlang.org"),
  bibEntry("Statista. (2024). Most used web frameworks among developers worldwide. https://www.statista.com/statistics/1124699/worldwide-developer-survey-most-used-frameworks-web/"),
  bibEntry("Tailwind Labs Inc. (2026). Tailwind CSS: Rapidly build modern websites without ever leaving your HTML. https://tailwindcss.com"),
  bibEntry("TradingView, Inc. (2026). Lightweight Charts: Performant financial charts built with HTML5 canvas. https://tradingview.github.io/lightweight-charts/"),
  bibEntry("Tufte, E. R. (1983). The visual display of quantitative information. Graphics Press."),
  bibEntry("Vercel, Inc. (2026). Next.js documentation. https://nextjs.org/docs"),
  bibEntry("Yahoo Finance. (2026). Yahoo Finance – Financial market data. https://finance.yahoo.com"),
];

// ── Build DOCX ────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullet-list",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "·",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  styles: {
    default: {
      document: {
        run: { font: FONT, size: FS },
        paragraph: { spacing: { after: 200 } },
      },
    },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: FONT }, paragraph: { spacing: { before: 480, after: 240 } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT }, paragraph: { spacing: { before: 360, after: 180 } } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: FS, bold: true, italics: true, font: FONT }, paragraph: { spacing: { before: 240, after: 120 } } },
    ],
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top:    convertInchesToTwip(1),
          right:  convertInchesToTwip(1),
          bottom: convertInchesToTwip(1),
          left:   convertInchesToTwip(1.18),
        },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "aaaaaa", space: 1 } },
            children: [new TextRun({ text: `Next.js – ${COURSE}`, font: FONT, size: 18, color: "666666" })],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "aaaaaa", space: 1 } },
            children: [
              new TextRun({ text: "Page ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "666666" }),
              new TextRun({ text: " of ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: "666666" }),
            ],
          }),
        ],
      }),
    },
    children: [
      ...COVER, ...ABSTRACT, ...INDEX, ...FIGURES,
      ...INTRO, ...SOA, ...METHOD, ...IMPL, ...CONCLUSIONS, ...BIBLIOGRAPHY,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  writeFileSync("docs/relatorio.docx", buf);
  console.log("OK  docs/relatorio.docx");
}).catch(e => console.error("ERROR DOCX:", e.message));

// ═══════════════════════════════════════════════════════════════════════════
// PPTX – light mode, professional, no emojis, English
// ═══════════════════════════════════════════════════════════════════════════

const prs = new pptxgen();
prs.layout  = "LAYOUT_16x9";   // 10 x 5.625 inches
prs.author  = STUDENTS.map(s => s.name).join(", ");
prs.subject = COURSE;
prs.title   = "Next.js – " + COURSE;

// Colour palette (light / academic)
const C = {
  BG:      "FFFFFF",
  PANEL:   "F1F5F9",
  PANEL2:  "EFF6FF",
  BORDER:  "CBD5E1",
  BLUE:    "1D4ED8",
  BLUE_LT: "DBEAFE",
  TEXT:    "1E293B",
  MUTED:   "64748B",
  GREEN:   "15803D",
  RED:     "B91C1C",
  BLACK:   "0F172A",
};

// Slide dimensions (LAYOUT_WIDE)
const SW = 10;     // slide width  (inches)
const SH = 5.625;  // slide height (inches)

// ── Helpers ───────────────────────────────────────────────────────────────

function addSlide(titleText, subText) {
  const s = prs.addSlide();
  s.background = { color: C.BG };

  // Blue rule at top
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.08, fill: { color: C.BLUE }, line: { color: C.BLUE } });

  if (titleText) {
    s.addText(titleText, {
      x: 0.45, y: 0.15, w: SW - 0.9, h: 0.62,
      fontSize: 22, bold: true, color: C.BLACK, fontFace: "Calibri",
      valign: "middle",
    });
  }
  if (subText) {
    s.addText(subText, {
      x: 0.45, y: 0.75, w: SW - 0.9, h: 0.3,
      fontSize: 12, color: C.MUTED, fontFace: "Calibri", italic: true,
    });
  }

  // Thin bottom rule
  const footerY = SH - 0.22;
  s.addShape(prs.ShapeType.rect, { x: 0, y: footerY - 0.04, w: SW, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });
  s.addText(`${COURSE}  |  ${YEAR}`, {
    x: 0.45, y: footerY, w: 6.5, h: 0.22,
    fontSize: 9, color: C.MUTED, fontFace: "Calibri",
  });
  s.addText(STUDENTS.map(st => st.id).join("  /  "), {
    x: SW - 2.8, y: footerY, w: 2.5, h: 0.22,
    fontSize: 9, color: C.MUTED, fontFace: "Calibri", align: "right",
  });

  return s;
}

function panel(s, x, y, w, h, color) {
  s.addShape(prs.ShapeType.rect, {
    x, y, w, h,
    fill: { color: color || C.PANEL },
    line: { color: C.BORDER, width: 0.5 },
  });
}

function blueLabel(s, text, x, y, w) {
  s.addText(text, { x, y, w, h: 0.3, fontSize: 11, bold: true, color: C.BLUE, fontFace: "Calibri" });
}

// Content area: from y=1.1 to y=SH-0.32
const CONTENT_TOP = 1.1;
const CONTENT_H   = SH - CONTENT_TOP - 0.32;   // ~4.2 in

// ── Slide 1: Cover ─────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.BG };

  // Full-width blue top bar
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SW, h: 1.55, fill: { color: C.BLUE }, line: { color: C.BLUE } });

  s.addText("Next.js", {
    x: 0, y: 0.06, w: SW, h: 0.82,
    fontSize: 52, bold: true, color: "FFFFFF", fontFace: "Calibri", align: "center", valign: "middle",
  });
  s.addText("Research and Development of a Web Application", {
    x: 0, y: 0.9, w: SW, h: 0.52,
    fontSize: 14, color: "DBEAFE", fontFace: "Calibri", align: "center", valign: "middle",
  });

  s.addText(COURSE, {
    x: 0, y: 1.72, w: SW, h: 0.38,
    fontSize: 13, color: C.MUTED, fontFace: "Calibri", align: "center", italic: true,
  });

  // Divider
  s.addShape(prs.ShapeType.rect, { x: 2.5, y: 2.26, w: 5.0, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  // Students
  STUDENTS.forEach((st, i) => {
    s.addText(`${st.name}  (${st.id})`, {
      x: 0, y: 2.42 + i * 0.44, w: SW, h: 0.38,
      fontSize: 14, color: C.TEXT, fontFace: "Calibri", align: "center",
    });
  });

  s.addText(`Lecturer: ${PROF}`, {
    x: 0, y: 3.42, w: SW, h: 0.34,
    fontSize: 12, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
  s.addText(`${UNIV}  |  ${YEAR}  |  May 2026`, {
    x: 0, y: 3.86, w: SW, h: 0.3,
    fontSize: 11, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
}

// ── Slide 2: Agenda ────────────────────────────────────────────────────────
{
  const s = addSlide("Agenda");

  const items = [
    ["01", "What is Next.js?"],
    ["02", "History and Evolution"],
    ["03", "Architecture and Main Features"],
    ["04", "Advantages and Disadvantages"],
    ["05", "Areas of Use"],
    ["06", "Demonstration Application – Market Dashboard"],
    ["07", "Difficulties, Limitations and Improvements"],
    ["08", "Conclusions and Feedback"],
  ];

  const colW = (SW - 0.9 - 0.2) / 2;   // ~4.45 in
  const startX = 0.45;
  const gap = 0.2;
  const rowH = 0.72;
  const startY = CONTENT_TOP;

  items.forEach(([num, label], i) => {
    const col = i < 4 ? 0 : 1;
    const row = i < 4 ? i : i - 4;
    const x = startX + col * (colW + gap);
    const y = startY + row * (rowH + 0.08);

    panel(s, x, y, colW, rowH);
    s.addText(num, {
      x: x + 0.12, y: y + 0.12, w: 0.55, h: 0.48,
      fontSize: 15, bold: true, color: C.BLUE, fontFace: "Calibri", valign: "middle",
    });
    s.addText(label, {
      x: x + 0.68, y: y + 0.12, w: colW - 0.82, h: 0.48,
      fontSize: 12, color: C.TEXT, fontFace: "Calibri", valign: "middle",
    });
  });
}

// ── Slide 3: What is Next.js ───────────────────────────────────────────────
{
  const s = addSlide("What is Next.js?", "Definition and context");

  panel(s, 0.45, CONTENT_TOP, SW - 0.9, 0.88, C.BLUE_LT);
  s.addText(
    "Next.js is an open-source React framework created by Vercel in 2016. It extends React with server-side rendering, static site generation, file-system-based routing, and integrated API Routes — all within a single codebase.",
    { x: 0.62, y: CONTENT_TOP + 0.1, w: SW - 1.22, h: 0.68, fontSize: 13, color: C.TEXT, fontFace: "Calibri", valign: "middle" }
  );

  const cols = [
    { title: "Created by",   body: "Vercel, Inc.\n(2016 – Guillermo Rauch)" },
    { title: "Based on",     body: "React\n(Meta's UI library)" },
    { title: "Licence",      body: "MIT (open-source)" },
    { title: "Version",      body: "16.x (2026)\n~130 k GitHub stars" },
  ];
  const colW = (SW - 0.9 - 0.15 * 3) / 4;   // 4 equal cols with gaps
  cols.forEach((col, i) => {
    const x = 0.45 + i * (colW + 0.15);
    panel(s, x, CONTENT_TOP + 1.05, colW, 1.45);
    blueLabel(s, col.title, x + 0.12, CONTENT_TOP + 1.14, colW - 0.2);
    s.addText(col.body, {
      x: x + 0.12, y: CONTENT_TOP + 1.44, w: colW - 0.2, h: 0.9,
      fontSize: 12, color: C.TEXT, fontFace: "Calibri",
    });
  });
}

// ── Slide 4: History ───────────────────────────────────────────────────────
{
  const s = addSlide("History and Evolution", "Key milestones of the framework");

  const events = [
    { year: "2016", label: "v1.0 – Launch (SSR, zero config, 6 core principles)" },
    { year: "2019", label: "v9   – Automatic dynamic routes, SSG" },
    { year: "2020", label: "v10  – Image optimisation (next/image), ISR" },
    { year: "2021", label: "v12  – Rust compiler (SWC), Edge Middleware" },
    { year: "2022", label: "v13  – App Router, React Server Components" },
    { year: "2023", label: "v14  – Server Actions, Turbopack (stable)" },
    { year: "2026", label: "v16  – RSC refinements, performance improvements" },
  ];

  const lineX = 1.18;
  const dotW  = 0.26;
  s.addShape(prs.ShapeType.rect, { x: lineX, y: CONTENT_TOP, w: 0.05, h: CONTENT_H, fill: { color: C.BLUE }, line: { color: C.BLUE } });

  events.forEach((ev, i) => {
    const y = CONTENT_TOP + i * (CONTENT_H / events.length);
    s.addShape(prs.ShapeType.ellipse, {
      x: lineX - dotW / 2 + 0.025, y: y + 0.04, w: dotW, h: dotW,
      fill: { color: C.BLUE }, line: { color: C.BLUE },
    });
    s.addText(ev.year, { x: 0.45, y, w: 0.68, h: 0.38, fontSize: 10, bold: true, color: C.BLUE, fontFace: "Calibri" });
    const isLast = i === events.length - 1;
    panel(s, lineX + 0.32, y, SW - lineX - 0.77, 0.42, isLast ? C.BLUE_LT : C.PANEL);
    s.addText(ev.label, { x: lineX + 0.46, y: y + 0.05, w: SW - lineX - 0.95, h: 0.32, fontSize: 11.5, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 5: Architecture ─────────────────────────────────────────────────
{
  const s = addSlide("Architecture and Main Features");

  const features = [
    { title: "Hybrid Rendering",              body: "SSR · SSG · ISR · CSR – choose per route or component" },
    { title: "App Router + Server Components",body: "Components run on server; no extra JS sent to client" },
    { title: "API Routes (Full-Stack)",        body: "Integrated REST endpoints — no separate server needed" },
    { title: "Automatic Optimisations",        body: "Code splitting · route prefetch · next/image · next/font" },
    { title: "Edge Runtime",                   body: "Functions close to the user; reduced latency globally" },
    { title: "SWC Compiler (Rust)",            body: "Builds up to 17x faster than Babel" },
  ];

  const cols = 2;
  const rows = Math.ceil(features.length / cols);
  const colW = (SW - 0.9 - 0.2) / cols;
  const rowH = CONTENT_H / rows - 0.1;

  features.forEach((f, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = 0.45 + col * (colW + 0.2);
    const y = CONTENT_TOP + row * (rowH + 0.1);
    panel(s, x, y, colW, rowH);
    blueLabel(s, f.title, x + 0.14, y + 0.1, colW - 0.26);
    s.addText(f.body, { x: x + 0.14, y: y + 0.42, w: colW - 0.26, h: rowH - 0.5, fontSize: 11, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 6: Advantages ───────────────────────────────────────────────────
{
  const s = addSlide("Advantages of Next.js");

  const items = [
    "Hybrid rendering – choose the ideal strategy per route (SSR, SSG, ISR, CSR)",
    "Excellent Developer Experience (DX) – zero config, hot-reload, native TypeScript",
    "Performance – automatic code splitting, route prefetch, image/font optimisation",
    "Full-stack in a single codebase – API Routes and Server Actions",
    "Native SEO – indexable content via SSR/SSG, unlike pure SPA applications",
    "Scalability – used in production by Netflix, TikTok, Hulu, GitHub",
    "Mature ecosystem – React community, extensive documentation, Vercel integration",
  ];

  const itemH = CONTENT_H / items.length - 0.06;

  items.forEach((item, i) => {
    const y = CONTENT_TOP + i * (itemH + 0.06);
    panel(s, 0.45, y, SW - 0.9, itemH);
    s.addShape(prs.ShapeType.rect, { x: 0.45, y, w: 0.12, h: itemH, fill: { color: C.GREEN }, line: { color: C.GREEN } });
    s.addText(item, { x: 0.7, y: y + 0.05, w: SW - 1.2, h: itemH - 0.1, fontSize: 11.5, color: C.TEXT, fontFace: "Calibri", valign: "middle" });
  });
}

// ── Slide 7: Disadvantages ────────────────────────────────────────────────
{
  const s = addSlide("Disadvantages and Limitations");

  const items = [
    "Learning curve – App Router and React Server Components require a paradigm shift",
    "Growing complexity – server/client boundary can be hard to manage in large projects",
    "Vercel dependency – commercial interests may influence the framework's direction",
    "Cold starts in serverless – initial latency on platforms without warm instances",
    "Build times – projects with many static pages can have lengthy compilation times",
    "Potential vendor lock-in – advanced features optimised for Vercel infrastructure",
  ];

  const itemH = CONTENT_H / items.length - 0.07;

  items.forEach((item, i) => {
    const y = CONTENT_TOP + i * (itemH + 0.07);
    panel(s, 0.45, y, SW - 0.9, itemH);
    s.addShape(prs.ShapeType.rect, { x: 0.45, y, w: 0.12, h: itemH, fill: { color: C.RED }, line: { color: C.RED } });
    s.addText(item, { x: 0.7, y: y + 0.05, w: SW - 1.2, h: itemH - 0.1, fontSize: 11.5, color: C.TEXT, fontFace: "Calibri", valign: "middle" });
  });
}

// ── Slide 8: Areas of Use ─────────────────────────────────────────────────
{
  const s = addSlide("Main Areas of Use");

  const areas = [
    { area: "E-commerce",           desc: "SSG + ISR for high-performance product pages (Shopify, Nike)" },
    { area: "SaaS Platforms",       desc: "App Router and RSC for dashboards and complex interfaces" },
    { area: "Content Portals",      desc: "SSG for blogs, newspapers, and documentation sites" },
    { area: "Full-Stack APIs",      desc: "API Routes for simple backends without extra infrastructure" },
    { area: "Data Dashboards",      desc: "Hybrid rendering for live data with ISR and polling" },
    { area: "Internal Applications",desc: "DX and native TypeScript accelerate development cycles" },
  ];

  const cols = 2;
  const colW = (SW - 0.9 - 0.2) / cols;
  const rows = Math.ceil(areas.length / cols);
  const rowH = CONTENT_H / rows - 0.1;

  areas.forEach((a, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = 0.45 + col * (colW + 0.2);
    const y = CONTENT_TOP + row * (rowH + 0.1);
    panel(s, x, y, colW, rowH);
    blueLabel(s, a.area, x + 0.14, y + 0.1, colW - 0.26);
    s.addText(a.desc, { x: x + 0.14, y: y + 0.42, w: colW - 0.26, h: rowH - 0.5, fontSize: 11, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 9: Comparison ───────────────────────────────────────────────────
{
  const s = addSlide("Comparison with Alternatives");

  const headers = ["Framework", "Base", "Rendering", "Full-Stack", "Strength"];
  const rows = [
    ["Next.js",   "React",  "SSR/SSG/ISR/CSR", "Yes",     "Flexibility and DX"],
    ["Nuxt.js",   "Vue 3",  "SSR/SSG/CSR",     "Yes",     "Vue ecosystem"],
    ["SvelteKit", "Svelte", "SSR/SSG/CSR",     "Yes",     "Performance / small bundle"],
    ["Remix",     "React",  "SSR",             "Yes",     "Web standards"],
    ["Astro",     "Multi",  "SSG/MPA",         "Partial", "Static content sites"],
  ];

  const colW = [1.7, 1.1, 2.15, 1.15, 3.6];  // sums to SW - 0.9
  const tableX = 0.45;
  const headerH = 0.4;
  const rowH    = (CONTENT_H - headerH - 0.12) / rows.length;

  // Header row
  let cx = tableX;
  headers.forEach((h, i) => {
    s.addShape(prs.ShapeType.rect, { x: cx, y: CONTENT_TOP, w: colW[i], h: headerH, fill: { color: C.BLUE }, line: { color: C.BLUE } });
    s.addText(h, { x: cx + 0.06, y: CONTENT_TOP + 0.04, w: colW[i] - 0.1, h: headerH - 0.08, fontSize: 11, bold: true, color: "FFFFFF", fontFace: "Calibri", valign: "middle" });
    cx += colW[i];
  });

  rows.forEach((row, ri) => {
    cx = tableX;
    const bg = ri === 0 ? C.BLUE_LT : (ri % 2 === 0 ? C.PANEL : C.BG);
    const isNext = ri === 0;
    row.forEach((cell, ci) => {
      const ry = CONTENT_TOP + headerH + 0.06 + ri * rowH;
      s.addShape(prs.ShapeType.rect, { x: cx, y: ry, w: colW[ci], h: rowH - 0.04, fill: { color: bg }, line: { color: C.BORDER, width: 0.3 } });
      s.addText(cell, { x: cx + 0.06, y: ry + 0.04, w: colW[ci] - 0.1, h: rowH - 0.12, fontSize: 11, bold: isNext, color: isNext ? C.BLUE : C.TEXT, fontFace: "Calibri", valign: "middle" });
      cx += colW[ci];
    });
  });

  s.addText("Next.js highlighted in blue", { x: 0.45, y: SH - 0.32, w: SW - 0.9, h: 0.22, fontSize: 10, italic: true, color: C.MUTED, fontFace: "Calibri" });
}

// ── Slide 10: Demonstration Application ──────────────────────────────────
{
  const s = addSlide("Market Dashboard", "Demonstration application built with Next.js");

  panel(s, 0.45, CONTENT_TOP, SW - 0.9, 0.72, C.BLUE_LT);
  s.addText("A real-time financial data visualisation platform, developed to demonstrate Next.js capabilities in a production-like context.", {
    x: 0.62, y: CONTENT_TOP + 0.1, w: SW - 1.22, h: 0.54, fontSize: 13, color: C.TEXT, fontFace: "Calibri", valign: "middle",
  });

  const feats = [
    { label: "Market Dashboard",      body: "6 assets (BTC, ETH, AAPL, MSFT, TSLA, NVDA) with live prices and search filter" },
    { label: "Detail Pages",          body: "Dynamic routes /market/[ticker] with open, high, low, volume metrics" },
    { label: "Candlestick Chart",     body: "Timeframes 1m to 1W; infinite scroll for history; 15-second polling via API Route" },
    { label: "Technical Indicators",  body: "MA 20, MA 50, Bollinger Bands (20 periods, sigma=2); computed on client" },
  ];

  const cols = 2;
  const colW = (SW - 0.9 - 0.2) / cols;
  const rowH = (CONTENT_H - 0.82 - 0.1) / 2;

  feats.forEach((f, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = 0.45 + col * (colW + 0.2);
    const y = CONTENT_TOP + 0.82 + row * (rowH + 0.1);
    panel(s, x, y, colW, rowH);
    blueLabel(s, f.label, x + 0.14, y + 0.1, colW - 0.26);
    s.addText(f.body, { x: x + 0.14, y: y + 0.42, w: colW - 0.26, h: rowH - 0.5, fontSize: 11, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 11: Next.js Features in Practice ────────────────────────────────
{
  const s = addSlide("Next.js Features in Practice", "How the framework was used in the application");

  const map = [
    { feature: "App Router + RSC",       usage: "/market and /market/[ticker] are Server Components – no extra JS on client" },
    { feature: "API Routes",             usage: "GET /api/stocks and GET /api/ohlcv/[ticker] – proxy to Yahoo Finance, no separate backend" },
    { feature: "ISR (revalidate: 30 s)", usage: "Price response caching; automatic invalidation without recompiling" },
    { feature: "next/dynamic (ssr:false)",usage: "lightweight-charts (DOM access) imported dynamically; avoids hydration errors" },
    { feature: "generateMetadata",       usage: "Dynamic per-ticker metadata for SEO" },
    { feature: "notFound()",             usage: "Redirects invalid tickers to Next.js native 404 page" },
  ];

  const itemH = CONTENT_H / map.length - 0.06;

  map.forEach((m, i) => {
    const y = CONTENT_TOP + i * (itemH + 0.06);
    panel(s, 0.45, y, SW - 0.9, itemH);
    s.addText(m.feature, { x: 0.6, y: y + 0.06, w: 2.85, h: itemH - 0.12, fontSize: 11, bold: true, color: C.BLUE, fontFace: "Courier New", valign: "middle" });
    s.addShape(prs.ShapeType.rect, { x: 3.5, y: y + 0.1, w: 0.02, h: itemH - 0.2, fill: { color: C.BORDER }, line: { color: C.BORDER } });
    s.addText(m.usage, { x: 3.65, y: y + 0.06, w: SW - 4.1, h: itemH - 0.12, fontSize: 11, color: C.TEXT, fontFace: "Calibri", valign: "middle" });
  });
}

// ── Slide 12: Difficulties and Limitations ────────────────────────────────
{
  const s = addSlide("Difficulties and Limitations");

  const leftW  = (SW - 0.9 - 0.2) / 2;
  const rightW = leftW;
  const leftX  = 0.45;
  const rightX = leftX + leftW + 0.2;

  panel(s, leftX,  CONTENT_TOP, leftW,  CONTENT_H);
  panel(s, rightX, CONTENT_TOP, rightW, CONTENT_H);

  s.addText("Difficulties", { x: leftX + 0.14, y: CONTENT_TOP + 0.12, w: leftW - 0.26, h: 0.36, fontSize: 13, bold: true, color: C.TEXT, fontFace: "Calibri" });
  [
    "Unofficial Yahoo Finance API – null fields, unavailability risk",
    "lightweight-charts v5 – API changed; identified via source code",
    "SSR vs client – next/dynamic required for DOM-dependent libraries",
    "Race conditions – concurrent polling, scroll, and TF switching",
  ].forEach((d, i) => {
    s.addShape(prs.ShapeType.rect, { x: leftX + 0.14, y: CONTENT_TOP + 0.6 + i * 0.84, w: 0.1, h: 0.5, fill: { color: C.BLUE }, line: { color: C.BLUE } });
    s.addText(d, { x: leftX + 0.32, y: CONTENT_TOP + 0.62 + i * 0.84, w: leftW - 0.5, h: 0.5, fontSize: 11, color: C.TEXT, fontFace: "Calibri" });
  });

  s.addText("Limitations", { x: rightX + 0.14, y: CONTENT_TOP + 0.12, w: rightW - 0.26, h: 0.36, fontSize: 13, bold: true, color: C.TEXT, fontFace: "Calibri" });
  [
    "Unofficial API – risk of breaking without notice",
    "6 fixed tickers – no free search",
    "15-second polling – not genuine WebSocket",
    "No authentication or portfolio",
    "1-minute history limited to 7 days",
  ].forEach((l, i) => {
    s.addShape(prs.ShapeType.rect, { x: rightX + 0.14, y: CONTENT_TOP + 0.62 + i * 0.7, w: 0.1, h: 0.42, fill: { color: C.RED }, line: { color: C.RED } });
    s.addText(l, { x: rightX + 0.32, y: CONTENT_TOP + 0.63 + i * 0.7, w: rightW - 0.5, h: 0.42, fontSize: 11, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 13: Conclusions and Feedback ───────────────────────────────────
{
  const s = addSlide("Conclusions and Feedback on Next.js");

  panel(s, 0.45, CONTENT_TOP, SW - 0.9, 0.88, C.BLUE_LT);
  s.addText(
    "Next.js proves to be a solid choice for modern web applications: hybrid rendering, the full-stack model, and Developer Experience justify its growing adoption. The App Router learning curve is real, but the investment is rewarded by the quality of the final product.",
    { x: 0.62, y: CONTENT_TOP + 0.1, w: SW - 1.22, h: 0.7, fontSize: 12.5, color: C.TEXT, fontFace: "Calibri", valign: "middle" }
  );

  const items = [
    { label: "Strengths",         body: "Flexibility, DX, full-stack, performance, SEO" },
    { label: "Areas to improve",  body: "RSC learning curve, Vercel dependency" },
    { label: "Recommended for",   body: "E-commerce, SaaS, dashboards, content portals" },
    { label: "Next steps",        body: "WebSocket, official API, more indicators, portfolio" },
  ];

  const cols = 2;
  const colW = (SW - 0.9 - 0.2) / cols;
  const rowH = (CONTENT_H - 1.04) / 2;

  items.forEach((it, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = 0.45 + col * (colW + 0.2);
    const y = CONTENT_TOP + 1.0 + row * (rowH + 0.1);
    panel(s, x, y, colW, rowH);
    blueLabel(s, it.label, x + 0.14, y + 0.1, colW - 0.26);
    s.addText(it.body, { x: x + 0.14, y: y + 0.42, w: colW - 0.26, h: rowH - 0.5, fontSize: 11.5, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 14: Questions ────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.BG };

  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: SW, h: 0.08, fill: { color: C.BLUE }, line: { color: C.BLUE } });
  s.addShape(prs.ShapeType.rect, { x: 0, y: SH - 0.26, w: SW, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  s.addText("Questions and Discussion", {
    x: 0, y: 1.5, w: SW, h: 0.88,
    fontSize: 36, bold: true, color: C.BLACK, fontFace: "Calibri", align: "center", valign: "middle",
  });
  s.addText("Thank you for your attention", {
    x: 0, y: 2.5, w: SW, h: 0.52,
    fontSize: 18, color: C.MUTED, fontFace: "Calibri", align: "center", italic: true,
  });

  s.addShape(prs.ShapeType.rect, { x: 3.0, y: 3.2, w: SW - 6.0, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  STUDENTS.forEach((st, i) => {
    s.addText(`${st.name}  |  ${st.id}`, {
      x: 0, y: 3.38 + i * 0.44, w: SW, h: 0.38,
      fontSize: 13, color: C.TEXT, fontFace: "Calibri", align: "center",
    });
  });
  s.addText(`${COURSE}  |  ${YEAR}`, {
    x: 0, y: 4.42, w: SW, h: 0.3,
    fontSize: 11, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
}

prs.writeFile({ fileName: "docs/apresentacao.pptx" }).then(() => {
  console.log("OK  docs/apresentacao.pptx");
}).catch(e => console.error("ERROR PPTX:", e.message));
