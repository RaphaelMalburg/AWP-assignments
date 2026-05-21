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

const TITLE    = "Next.js: Investigacao e Desenvolvimento de uma Aplicacao Web";
const COURSE   = "Advanced Web Programming";
const YEAR     = "2025/2026 - 1. Semestre";
const UNIV     = "[Nome da Universidade]";
const DEPT     = "[Departamento / Escola]";
const PROF     = "[Nome do/a Docente]";
const STUDENTS = [
  { name: "[Nome do Estudante 1]", id: "N. XXXXXX" },
  { name: "[Nome do Estudante 2]", id: "N. XXXXXX" },
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
  center("Trabalho Pratico – " + COURSE, { italics: true, size: 26 }),
  empty(), empty(),
  ...STUDENTS.map(s => center(`${s.name}  |  ${s.id}`)),
  empty(),
  center(`Docente: ${PROF}`),
  empty(), empty(),
  center(YEAR),
  center("Maio de 2026"),
  br(),
];

const RESUMO = [
  h1("Resumo"),
  para("O presente trabalho, desenvolvido no ambito da unidade curricular de Advanced Web Programming, tem como objetivo investigar o framework Next.js, desenvolvido pela Vercel, analisando o seu estado da arte, vantagens, desvantagens e principais areas de utilizacao. Paralelamente, e apresentada uma aplicacao web de demonstracao denominada Market Dashboard, desenvolvida com o mesmo framework, que ilustra as suas capacidades no contexto de visualizacao de dados financeiros em tempo real. A aplicacao integra a API nao oficial do Yahoo Finance, apresenta graficos de candlestick com selecao de timeframe e suporta indicadores tecnicos configurAveis, servindo como prova de conceito das funcionalidades do Next.js em ambiente de producao."),
  empty(),
  paraRuns([run("Palavras-chave: ", { bold: true }), run("Next.js, React, framework web, renderizacao hibrida, aplicacao web, visualizacao financeira.")]),
  br(),
];

const INDEX = [
  h1("Indice"),
  para("Nota: No Microsoft Word, clique com o botao direito sobre o indice e selecione \"Atualizar Campo\" para gerar a numeracao de paginas.", { italics: true }),
  new Paragraph({
    spacing: { after: 200 },
    children: [new TableOfContents("Indice", { hyperlink: true, headingStyleRange: "1-3" })],
  }),
  br(),
];

const FIGURES = [
  h1("Lista de Figuras e Tabelas"),
  empty(),
  paraRuns([run("Figuras", { bold: true })]),
  para("Figura 1 – Arquitetura de renderizacao hibrida do Next.js ................ 5"),
  para("Figura 2 – Comparacao de abordagens de renderizacao (CSR vs SSR vs SSG) . 6"),
  para("Figura 3 – Estrutura de pastas da aplicacao Market Dashboard ............. 9"),
  para("Figura 4 – Interface do dashboard de mercado ............................. 10"),
  para("Figura 5 – Grafico de candlestick com indicadores tecnicos .............. 11"),
  empty(),
  paraRuns([run("Tabelas", { bold: true })]),
  para("Tabela 1 – Comparacao de frameworks: Next.js vs alternativas ............. 7"),
  para("Tabela 2 – Tecnologias utilizadas na implementacao ........................ 9"),
  para("Tabela 3 – Endpoints da API interna ..................................... 10"),
  br(),
];

const INTRO = [
  h1("1. Introducao"),
  para("O desenvolvimento de aplicacoes web tem evoluido a um ritmo acelerado nas ultimas duas decadas, impulsionado pela crescente exigencia de experiencias de utilizador rapidas, acessiveis e dinamicas. Neste contexto, os frameworks web emergem como ferramentas fundamentais para organizar e abstrair a complexidade do desenvolvimento, permitindo que as equipas se concentrem na logica de negocio em vez de resolver problemas de infraestrutura de forma repetida."),
  para("O presente trabalho debruca-se sobre o Next.js, um framework React de codigo aberto criado pela Vercel em 2016, que se tem destacado pelo seu modelo de renderizacao hibrida, capaz de combinar renderizacao no servidor (Server-Side Rendering, SSR), geracao de conteudo estatico (Static Site Generation, SSG) e renderizacao no cliente (Client-Side Rendering, CSR) numa unica aplicacao. Este modelo confere ao Next.js uma flexibilidade singular no panorama dos frameworks web modernos."),
  h2("1.1 Objetivos"),
  bullet("Investigar o estado da arte do Next.js, incluindo a sua historia, arquitetura e ecossistema;"),
  bullet("Identificar as vantagens, desvantagens e principais areas de utilizacao do framework;"),
  bullet("Avaliar a experiencia de desenvolvimento com Next.js e os seus beneficios praticos;"),
  bullet("Desenvolver uma aplicacao de demonstracao – Market Dashboard – utilizando o framework investigado."),
  h2("1.2 Estrutura do Relatorio"),
  para("O relatorio esta organizado da seguinte forma: a Seccao 2 apresenta o estado da arte do Next.js; a Seccao 3 descreve a metodologia de resolucao do problema; a Seccao 4 detalha a implementacao da aplicacao; a Seccao 5 apresenta as conclusoes, dificuldades e melhorias propostas; a Seccao 6 lista as referencias bibliograficas."),
  br(),
];

const SOA = [
  h1("2. Estado da Arte – Next.js"),
  h2("2.1 Historia e Evolucao"),
  para("O Next.js foi lancado publicamente em outubro de 2016 por Guillermo Rauch e a equipa da Vercel (entao denominada Zeit), como resposta a um problema concreto: construir aplicacoes React com renderizacao do lado do servidor de forma simples e sem configuracao manual complexa. A primeira versao introduziu seis principios fundamentais: renderizacao no servidor por defeito, carregamento de codigo por rota, suporte a CSS, hot-reloading em desenvolvimento, ciclo de compilacao simples e configuracao minima (Vercel, Inc., 2026)."),
  para("Desde entao, o framework tem evoluido de forma consistente. A versao 9 (2019) introduziu o roteamento dinamico automatico; a versao 10 (2020) trouxe otimizacao de imagens integrada; a versao 12 (2021) introduziu o compilador Rust (SWC) e middleware a nivel de borda; a versao 13 (2022) lancou o App Router baseado em React Server Components; a versao 14 (2023) estabilizou o App Router e introduziu Server Actions; a versao 15 e posteriores continuaram a refinar o modelo de componentes e a performance (Vercel, Inc., 2026)."),
  h2("2.2 Arquitetura e Funcionalidades Principais"),
  para("O Next.js baseia-se no React e estende-o com um conjunto de funcionalidades que cobrem tanto o frontend como o backend de uma aplicacao web:"),
  h3("2.2.1 Modelos de Renderizacao"),
  para("Uma das caracteristicas mais distintas do Next.js e a sua capacidade de combinar diferentes estrategias de renderizacao na mesma aplicacao:"),
  bullet("Server-Side Rendering (SSR): a pagina e gerada no servidor em cada pedido, garantindo dados sempre atualizados;"),
  bullet("Static Site Generation (SSG): as paginas sao pre-geradas em tempo de compilacao, oferecendo o maximo desempenho;"),
  bullet("Incremental Static Regeneration (ISR): permite atualizar paginas estaticas em segundo plano sem re-compilar toda a aplicacao;"),
  bullet("Client-Side Rendering (CSR): renderizacao tradicional no browser, adequada para componentes altamente interativos."),
  h3("2.2.2 App Router e React Server Components"),
  para("A partir da versao 13, o Next.js introduziu o App Router, baseado em React Server Components (RSC). Os RSC permitem que os componentes sejam executados exclusivamente no servidor, reduzindo o bundle JavaScript enviado ao cliente. Esta arquitetura distingue claramente entre componentes de servidor (sem interatividade, sem acesso ao browser) e componentes de cliente (marcados com a diretiva 'use client'), que hidratam no browser e suportam estado e eventos."),
  h3("2.2.3 API Routes e Full-Stack"),
  para("O Next.js inclui suporte nativo a API Routes, permitindo criar endpoints REST diretamente na estrutura da aplicacao, sem necessidade de um servidor separado. Esta funcionalidade posiciona o framework como uma solucao fullstack, integrando frontend e backend num unico projeto."),
  h3("2.2.4 Otimizacoes Integradas"),
  para("O framework inclui um conjunto de otimizacoes automaticas: otimizacao de imagens (next/image), carregamento de fontes otimizado (next/font), prefetching automatico de rotas, divisao de codigo (code splitting) por rota e suporte a Edge Runtime para funcoes proximas do utilizador."),
  h2("2.3 Ecossistema e Adocao"),
  para("O Next.js e atualmente um dos frameworks web mais adotados a nivel mundial. De acordo com as sondagens anuais da State of JavaScript (2023), o Next.js e consistentemente classificado como o framework React mais utilizado em ambiente de producao, com uma taxa de satisfacao elevada entre os utilizadores. Empresas como Netflix, TikTok, Twitch, Hulu e GitHub utilizam Next.js nas suas plataformas (Vercel, Inc., 2026)."),
  para("O ecossistema do Next.js beneficia da vasta comunidade React, bem como de integracao nativa com a plataforma de deploy da Vercel, que oferece funcionalidades como previews automaticos por pull request, analytics integrado e suporte a Edge Network global."),
  h2("2.4 Vantagens"),
  bullet("Renderizacao hibrida: possibilidade de escolher a estrategia de renderizacao mais adequada por rota ou componente;"),
  bullet("Developer Experience (DX): configuracao minima, hot-reloading, TypeScript integrado e mensagens de erro detalhadas;"),
  bullet("Performance: code splitting automatico, prefetching de rotas, otimizacao de imagens e fontes;"),
  bullet("Full-stack numa unica base de codigo: API Routes e Server Actions eliminam a necessidade de um backend separado para casos de uso comuns;"),
  bullet("SEO: SSR e SSG garantem que o conteudo e indexavel por motores de busca, ao contrario de aplicacoes SPA puras;"),
  bullet("Comunidade e ecossistema: ampla comunidade, documentacao extensa e integracao com o ecossistema React;"),
  bullet("Escalabilidade: adequado tanto para projetos pequenos como para aplicacoes de grande escala em producao."),
  h2("2.5 Desvantagens"),
  bullet("Curva de aprendizagem: o modelo de App Router com React Server Components introduz conceitos novos que exigem compreensao aprofundada;"),
  bullet("Complexidade crescente: a separacao entre componentes de servidor e de cliente pode tornar a arquitetura dificil de gerir em projetos grandes;"),
  bullet("Dependencia da Vercel: embora seja open-source, o Next.js e desenvolvido pela Vercel, cujos interesses comerciais podem influenciar a direcao do framework;"),
  bullet("Cold starts em serverless: funcoes serverless podem sofrer latencia inicial em plataformas que nao mantem instancias aquecidas;"),
  bullet("Build times: projetos com muitas paginas estaticas podem ter tempos de compilacao elevados;"),
  bullet("Vendor lock-in potencial: algumas funcionalidades avancadas (como Edge Middleware otimizado) funcionam melhor na infraestrutura da Vercel."),
  h2("2.6 Principais Areas de Utilizacao"),
  bullet("Plataformas de e-commerce (Shopify Hydrogen, Nike.com): beneficiam do SSG e ISR para paginas de produto altamente performativas;"),
  bullet("Aplicacoes SaaS e dashboards: o App Router e os Server Components simplificam a construcao de interfaces complexas com dados dinamicos;"),
  bullet("Portais de conteudo e media (blogs, jornais online): o SSG gera paginas estaticas com performance maxima;"),
  bullet("APIs e aplicacoes fullstack: as API Routes permitem construir backends simples sem infraestrutura adicional;"),
  bullet("Aplicacoes financeiras e de dados em tempo real: a flexibilidade de renderizacao e o suporte a atualizacoes incrementais tornam o Next.js adequado para dashboards de dados ao vivo."),
  h2("2.7 Comparacao com Alternativas"),
  para("A Tabela 1 compara o Next.js com frameworks alternativos no mesmo espaco de mercado:"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Framework", "Base", "Renderizacao", "Full-Stack", "Ponto Forte"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["Next.js",  "React",  "SSR/SSG/CSR/ISR", "Sim (API Routes)", "Flexibilidade e DX"],
        ["Nuxt.js",  "Vue 3",  "SSR/SSG/CSR",     "Sim",              "Ecossistema Vue"],
        ["SvelteKit","Svelte", "SSR/SSG/CSR",     "Sim",              "Performance / bundle reduzido"],
        ["Remix",    "React",  "SSR",             "Sim",              "Web standards / forms"],
        ["Astro",    "Multi",  "SSG/MPA",         "Parcial",          "Sites de conteudo estatico"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Tabela 1 – Comparacao de frameworks: Next.js vs alternativas"),
  empty(),
  br(),
];

const METHOD = [
  h1("3. Metodologia de Resolucao do Problema"),
  h2("3.1 Abordagem de Investigacao"),
  para("A investigacao sobre o Next.js foi conduzida com base em fontes primarias (documentacao oficial do framework, changelogs de versoes, repositorio GitHub) e secundarias (artigos academicos, publicacoes tecnicas, sondagens da comunidade). Esta analise permitiu estruturar o estado da arte apresentado na Seccao 2 e identificar os casos de uso mais relevantes do framework."),
  h2("3.2 Selecao da Aplicacao de Demonstracao"),
  para("Para demonstrar as capacidades do Next.js em contexto real, foi desenvolvida uma aplicacao denominada Market Dashboard, uma plataforma de visualizacao de dados financeiros em tempo real. Esta escolha foi motivada pela necessidade de exercitar funcionalidades-chave do framework:"),
  bullet("API Routes (full-stack): implementacao de endpoints REST para acesso a dados externos;"),
  bullet("App Router com componentes de servidor e de cliente: gestao eficiente do rendering hibrido;"),
  bullet("ISR e cache de dados: politicas de invalidacao de cache para dados com diferentes frequencias de atualizacao;"),
  bullet("Componentes de cliente interativos: graficos, filtros e atualizacoes em tempo real."),
  h2("3.3 Processo de Desenvolvimento"),
  para("O desenvolvimento seguiu uma abordagem iterativa dividida em tres fases:"),
  bullet("Fase 1 – Dashboard base: pagina de listagem de ativos com precos atualizados via API do Yahoo Finance, com fallback para dados simulados em caso de falha;"),
  bullet("Fase 2 – Paginas de detalhe: rotas dinamicas /market/[ticker] com metricas por ativo (abertura, maximo, minimo, volume);"),
  bullet("Fase 3 – Visualizacao avancada: grafico de candlestick com selecao de timeframe, scroll infinito para dados historicos, polling a 15 segundos e indicadores tecnicos configurAveis."),
  br(),
];

const IMPL = [
  h1("4. Descricao da Implementacao"),
  h2("4.1 Tecnologias Utilizadas"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Tecnologia", "Versao", "Papel na Aplicacao"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["Next.js",            "16.2.6",    "Framework principal (App Router, SSR, API Routes)"],
        ["React",              "19.2.4",    "Biblioteca de UI"],
        ["TypeScript",         "5.x",       "Tipagem estatica"],
        ["Tailwind CSS",       "4.x",       "Estilizacao utilitaria"],
        ["lightweight-charts", "5.2.0",     "Graficos financeiros (TradingView)"],
        ["Yahoo Finance API",  "Nao oficial","Fonte de dados de mercado"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Tabela 2 – Tecnologias utilizadas na implementacao"),
  empty(),
  h2("4.2 Arquitetura da Aplicacao"),
  para("A aplicacao explora o App Router do Next.js 16, que distingue componentes de servidor de componentes de cliente. Os componentes de servidor executam no servidor e produzem HTML, sem enviar JavaScript ao browser; os componentes de cliente (marcados com 'use client') hidratam no browser e suportam interatividade."),
  para("A estrutura de rotas e a seguinte:"),
  bullet("app/market/page.tsx – componente de servidor; lista de ativos;"),
  bullet("app/market/[ticker]/page.tsx – componente de servidor; detalhe do ativo;"),
  bullet("app/api/stocks/route.ts – API Route; precos atuais (cache ISR de 30 s);"),
  bullet("app/api/ohlcv/[ticker]/route.ts – API Route; dados OHLCV historicos;"),
  bullet("components/candlestick-chart.tsx – wrapper com next/dynamic ssr:false;"),
  bullet("components/candlestick-chart-inner.tsx – componente de cliente; grafico interativo."),
  h2("4.3 Camada de API"),
  empty(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: ["Endpoint", "Parametros", "Descricao"].map(h =>
          new TableCell({ children: [para(run(h, { bold: true }))] })
        ),
      }),
      ...[
        ["GET /api/stocks",         "–",               "Precos atuais dos 6 ativos (ISR 30 s); fallback simulado"],
        ["GET /api/ohlcv/[ticker]", "interval, before", "~400 velas OHLCV do Yahoo Finance; paginacao por 'before'"],
      ].map(row =>
        new TableRow({ children: row.map(c => new TableCell({ children: [para(c)] })) })
      ),
    ],
  }),
  empty(),
  captionPara("Tabela 3 – Endpoints da API interna"),
  empty(),
  h2("4.4 Grafico de Candlestick e Indicadores"),
  para("O componente CandlestickChartInner utiliza a biblioteca lightweight-charts v5 da TradingView e implementa quatro comportamentos:"),
  bullet("Carregamento inicial: ~400 velas por timeframe (1m, 5m, 15m, 1H, 1D, 1W);"),
  bullet("Scroll infinito: detetado via subscribeVisibleLogicalRangeChange; carrega dados mais antigos ao atingir a borda esquerda;"),
  bullet("Polling a 15 segundos: atualiza a vela em curso e apende novas velas; guardado por pollInFlightRef para evitar concorrencia;"),
  bullet("Indicadores tecnicos: MA 20, MA 50 e Bandas de Bollinger (20 periodos, sigma=2), calculados no cliente e desenhados como series de linha sobrepostas."),
  h2("4.5 Funcionalidades de Destaque do Next.js Utilizadas"),
  bullet("App Router com Server e Client Components: separacao clara de responsabilidades e otimizacao do bundle;"),
  bullet("API Routes como proxy: evita exposicao de chamadas externas ao cliente e centraliza tratamento de erros;"),
  bullet("ISR (revalidate: 30): cache das respostas de preco com invalidacao automatica a cada 30 segundos;"),
  bullet("next/dynamic com ssr: false: importacao dinamica para bibliotecas que acedem ao DOM (lightweight-charts);"),
  bullet("generateMetadata: metadados dinamicos por pagina para SEO."),
  br(),
];

const CONCLUSIONS = [
  h1("5. Conclusoes"),
  h2("5.1 Dificuldades Encontradas"),
  bullet("API do Yahoo Finance: API nao oficial e nao documentada, com campos nulos em series OHLCV e risco de indisponibilidade; resolvido com fallback de dados simulados;"),
  bullet("Componentes SSR vs. cliente: lightweight-charts acede ao DOM na importacao, causando erros de hidratacao; resolvido com next/dynamic e ssr: false;"),
  bullet("Migracao para lightweight-charts v5: API alterada (addCandlestickSeries removido, substituido por addSeries com definicao de serie como argumento); identificado por analise do codigo-fonte;"),
  bullet("Gestao de estado assincrono: coexistencia de polling, paginacao por scroll e troca de timeframe exigiu uso de refs mutaveis (useRef) e guarda pollInFlightRef para evitar condicoes de corrida."),
  h2("5.2 Limitacoes"),
  bullet("Dependencia de API nao oficial: risco de quebra sem aviso;"),
  bullet("Conjunto fixo de ativos: limitado a seis tickers predefinidos;"),
  bullet("Polling vs. tempo real: latencia de ate 15 segundos comparativamente a WebSocket;"),
  bullet("Ausencia de persistencia: sem autenticacao, perfis ou gestao de portfolio;"),
  bullet("Dados historicos limitados por intervalo: a API restringe dados de 1 minuto aos ultimos 7 dias."),
  h2("5.3 Alternativas e Melhorias"),
  bullet("WebSocket ou Server-Sent Events para atualizacoes genuinamente em tempo real;"),
  bullet("API oficial (Polygon.io, Finnhub) para maior fiabilidade e historico mais extenso;"),
  bullet("Suporte a pesquisa de qualquer ticker e adicao dinamica de ativos;"),
  bullet("Indicadores adicionais: RSI, MACD, Estocastico;"),
  bullet("Autenticacao e gestao de portfolio com alertas de preco;"),
  bullet("Testes automatizados (Vitest + Playwright) para garantir estabilidade."),
  h2("5.4 Feedback sobre o Next.js"),
  para("A experiencia de desenvolvimento com o Next.js revelou-se muito positiva. A separacao entre componentes de servidor e de cliente, embora inicialmente confusa, representa uma abstratcao poderosa que permite otimizacoes significativas de performance sem sacrificar a produtividade. A abordagem full-stack, com API Routes integradas, simplificou substancialmente a arquitetura do projeto, eliminando a necessidade de um servidor backend separado."),
  para("A documentacao oficial e extensa e bem organizada, e a experiencia de desenvolvimento em modo de desenvolvimento (hot-reloading, mensagens de erro detalhadas) e notavelmente fluida. Em contrapartida, o App Router com React Server Components tem uma curva de aprendizagem consideravel, e a fronteira entre o que pode e o que nao pode ser feito num componente de servidor nem sempre e imediatamente obvio."),
  para("Em suma, o Next.js revela-se uma escolha solida para aplicacoes web modernas que exijam performance, SEO e escalabilidade, com um equilibrio muito favoravel entre facilidade de uso e capacidade tecnica."),
  br(),
];

const BIBLIOGRAPHY = [
  h1("6. Bibliografia"),
  para("Referencias formatadas segundo a norma APA 7.a edicao."),
  empty(),
  bibEntry("Fielding, R. T. (2000). Architectural styles and the design of network-based software architectures [Dissertacao de doutoramento, University of California, Irvine]. https://roy.gbiv.com/pubs/dissertation/top.htm"),
  bibEntry("Meta Platforms, Inc. (2026). React: A JavaScript library for building user interfaces. https://react.dev"),
  bibEntry("Microsoft. (2012-2026). TypeScript: JavaScript with syntax for types. https://www.typescriptlang.org"),
  bibEntry("Statista. (2024). Most used web frameworks among developers worldwide. https://www.statista.com/statistics/1124699/worldwide-developer-survey-most-used-frameworks-web/"),
  bibEntry("Tailwind Labs Inc. (2026). Tailwind CSS: Rapidly build modern websites without ever leaving your HTML. https://tailwindcss.com"),
  bibEntry("TradingView, Inc. (2026). Lightweight Charts: Performant financial charts built with HTML5 canvas. https://tradingview.github.io/lightweight-charts/"),
  bibEntry("Tufte, E. R. (1983). The visual display of quantitative information. Graphics Press."),
  bibEntry("Vercel, Inc. (2026). Next.js documentation. https://nextjs.org/docs"),
  bibEntry("Yahoo Finance. (2026). Yahoo Finance – Dados de mercado financeiro. https://finance.yahoo.com"),
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
              new TextRun({ text: "Pagina ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "666666" }),
              new TextRun({ text: " de ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: "666666" }),
            ],
          }),
        ],
      }),
    },
    children: [
      ...COVER, ...RESUMO, ...INDEX, ...FIGURES,
      ...INTRO, ...SOA, ...METHOD, ...IMPL, ...CONCLUSIONS, ...BIBLIOGRAPHY,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  writeFileSync("docs/relatorio.docx", buf);
  console.log("OK  docs/relatorio.docx");
}).catch(e => console.error("ERRO DOCX:", e.message));

// ═══════════════════════════════════════════════════════════════════════════
// PPTX – light mode, professional, no emojis
// ═══════════════════════════════════════════════════════════════════════════

const prs = new pptxgen();
prs.layout  = "LAYOUT_WIDE";
prs.author  = STUDENTS.map(s => s.name).join(", ");
prs.subject = COURSE;
prs.title   = "Next.js – " + COURSE;

// Colour palette (light / academic)
const C = {
  BG:      "FFFFFF",   // slide background
  PANEL:   "F1F5F9",   // card / panel fill
  PANEL2:  "EFF6FF",   // blue-tinted panel
  BORDER:  "CBD5E1",   // subtle border
  BLUE:    "1D4ED8",   // primary accent
  BLUE_LT: "DBEAFE",   // light blue highlight
  TEXT:    "1E293B",   // dark text
  MUTED:   "64748B",   // secondary text
  GREEN:   "15803D",   // positive indicator
  RED:     "B91C1C",   // negative indicator
  BLACK:   "0F172A",   // headings
};

// ── Helpers ───────────────────────────────────────────────────────────────

function addSlide(titleText, subText) {
  const s = prs.addSlide();
  s.background = { color: C.BG };

  // Blue rule at top
  s.addShape(prs.ShapeType.rect, {
    x: 0, y: 0, w: 10, h: 0.08,
    fill: { color: C.BLUE }, line: { color: C.BLUE },
  });

  if (titleText) {
    s.addText(titleText, {
      x: 0.45, y: 0.18, w: 9.1, h: 0.6,
      fontSize: 24, bold: true, color: C.BLACK, fontFace: "Calibri",
    });
  }
  if (subText) {
    s.addText(subText, {
      x: 0.45, y: 0.78, w: 9.1, h: 0.32,
      fontSize: 13, color: C.MUTED, fontFace: "Calibri", italic: true,
    });
  }

  // Thin bottom rule
  s.addShape(prs.ShapeType.rect, {
    x: 0, y: 5.42, w: 10, h: 0.04,
    fill: { color: C.BORDER }, line: { color: C.BORDER },
  });
  // Footer text
  s.addText(`${COURSE}  |  ${YEAR}`, {
    x: 0.45, y: 5.46, w: 7.0, h: 0.28,
    fontSize: 9, color: C.MUTED, fontFace: "Calibri",
  });
  s.addText(STUDENTS.map(st => st.id).join("  /  "), {
    x: 7.5, y: 5.46, w: 2.2, h: 0.28,
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
  s.addText(text, { x, y, w, h: 0.32, fontSize: 11, bold: true, color: C.BLUE, fontFace: "Calibri" });
}

// ── Slide 1: Capa ──────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.BG };

  // Full-width top bar
  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 1.5, fill: { color: C.BLUE }, line: { color: C.BLUE } });
  s.addText("Next.js", {
    x: 0.5, y: 0.1, w: 9.0, h: 0.75,
    fontSize: 52, bold: true, color: "FFFFFF", fontFace: "Calibri", align: "center",
  });
  s.addText("Investigacao e Desenvolvimento de uma Aplicacao Web", {
    x: 0.5, y: 0.85, w: 9.0, h: 0.55,
    fontSize: 14, color: "DBEAFE", fontFace: "Calibri", align: "center",
  });

  s.addText(COURSE, {
    x: 0.5, y: 1.75, w: 9.0, h: 0.42,
    fontSize: 14, color: C.MUTED, fontFace: "Calibri", align: "center", italic: true,
  });

  s.addShape(prs.ShapeType.rect, { x: 2.5, y: 2.35, w: 5.0, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  STUDENTS.forEach((st, i) => {
    s.addText(`${st.name}  (${st.id})`, {
      x: 0.5, y: 2.55 + i * 0.42, w: 9.0, h: 0.38,
      fontSize: 14, color: C.TEXT, fontFace: "Calibri", align: "center",
    });
  });

  s.addText(`Docente: ${PROF}`, {
    x: 0.5, y: 3.55, w: 9.0, h: 0.35,
    fontSize: 12, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
  s.addText(`${UNIV}  |  ${YEAR}  |  Maio 2026`, {
    x: 0.5, y: 4.05, w: 9.0, h: 0.32,
    fontSize: 11, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
}

// ── Slide 2: Agenda ────────────────────────────────────────────────────────
{
  const s = addSlide("Agenda");
  const items = [
    ["01", "O que e o Next.js?"],
    ["02", "Historia e Evolucao"],
    ["03", "Arquitetura e Funcionalidades Principais"],
    ["04", "Vantagens e Desvantagens"],
    ["05", "Areas de Utilizacao"],
    ["06", "Aplicacao de Demonstracao – Market Dashboard"],
    ["07", "Dificuldades, Limitacoes e Melhorias"],
    ["08", "Conclusoes e Feedback"],
  ];
  items.forEach(([num, label], i) => {
    const col = i < 4 ? 0 : 1;
    const row = col === 0 ? i : i - 4;
    panel(s, 0.45 + col * 4.9, 1.15 + row * 0.93, 4.65, 0.75);
    s.addText(num, {
      x: 0.6 + col * 4.9, y: 1.23 + row * 0.93, w: 0.55, h: 0.55,
      fontSize: 15, bold: true, color: C.BLUE, fontFace: "Calibri",
    });
    s.addText(label, {
      x: 1.12 + col * 4.9, y: 1.26 + row * 0.93, w: 3.8, h: 0.5,
      fontSize: 13, color: C.TEXT, fontFace: "Calibri",
    });
  });
}

// ── Slide 3: O que e o Next.js ─────────────────────────────────────────────
{
  const s = addSlide("O que e o Next.js?", "Definicao e contexto");

  panel(s, 0.45, 1.18, 9.1, 0.9, C.BLUE_LT);
  s.addText(
    "Next.js e um framework React de codigo aberto criado pela Vercel em 2016, que estende o React com renderizacao do lado do servidor, geracao de conteudo estatico, roteamento baseado no sistema de ficheiros e API Routes integradas, numa unica base de codigo.",
    { x: 0.65, y: 1.26, w: 8.7, h: 0.75, fontSize: 13, color: C.TEXT, fontFace: "Calibri" }
  );

  const cols = [
    { title: "Criado por", body: "Vercel, Inc.\n(2016 – Guillermo Rauch)" },
    { title: "Base",       body: "React\n(biblioteca de UI da Meta)" },
    { title: "Licenca",    body: "MIT (open-source)" },
    { title: "Versao",     body: "16.x (2026)\n~130 k estrelas GitHub" },
  ];
  cols.forEach((col, i) => {
    panel(s, 0.45 + i * 2.3, 2.3, 2.15, 1.35);
    blueLabel(s, col.title, 0.6 + i * 2.3, 2.38, 2.0);
    s.addText(col.body, {
      x: 0.6 + i * 2.3, y: 2.68, w: 1.95, h: 0.82,
      fontSize: 12, color: C.TEXT, fontFace: "Calibri",
    });
  });
}

// ── Slide 4: Historia ──────────────────────────────────────────────────────
{
  const s = addSlide("Historia e Evolucao", "Principais marcos do framework");

  const events = [
    { year: "2016", label: "v1.0 – Lancamento (SSR simples, zero config)" },
    { year: "2019", label: "v9   – Rotas dinamicas automaticas, SSG" },
    { year: "2020", label: "v10  – Otimizacao de imagens (next/image), ISR" },
    { year: "2021", label: "v12  – Compilador Rust (SWC), Middleware" },
    { year: "2022", label: "v13  – App Router, React Server Components" },
    { year: "2023", label: "v14  – Server Actions, Turbopack (stable)" },
    { year: "2026", label: "v16  – Refinamentos RSC, melhorias de performance" },
  ];

  // Timeline line
  s.addShape(prs.ShapeType.rect, { x: 1.2, y: 1.2, w: 0.05, h: 3.9, fill: { color: C.BLUE }, line: { color: C.BLUE } });

  events.forEach((ev, i) => {
    const y = 1.18 + i * 0.58;
    // Dot
    s.addShape(prs.ShapeType.ellipse, { x: 1.08, y: y + 0.04, w: 0.28, h: 0.28, fill: { color: C.BLUE }, line: { color: C.BLUE } });
    // Year
    s.addText(ev.year, { x: 0.45, y: y, w: 0.6, h: 0.36, fontSize: 11, bold: true, color: C.BLUE, fontFace: "Calibri" });
    // Label
    panel(s, 1.55, y, 7.9, 0.42, i === events.length - 1 ? C.BLUE_LT : C.PANEL);
    s.addText(ev.label, { x: 1.7, y: y + 0.05, w: 7.6, h: 0.32, fontSize: 12, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 5: Arquitetura ───────────────────────────────────────────────────
{
  const s = addSlide("Arquitetura e Funcionalidades Principais");

  const features = [
    { title: "Renderizacao Hibrida",          body: "SSR · SSG · ISR · CSR – escolha por rota ou componente" },
    { title: "App Router + Server Components", body: "Componentes executados no servidor; sem JS extra no cliente" },
    { title: "API Routes (Full-Stack)",        body: "Endpoints REST integrados sem servidor separado" },
    { title: "Otimizacoes automaticas",        body: "Code splitting · prefetch de rotas · next/image · next/font" },
    { title: "Edge Runtime",                   body: "Funcoes proximas do utilizador; latencia reduzida" },
    { title: "Compilador SWC (Rust)",          body: "Builds ate 17x mais rapidos vs. Babel" },
  ];

  features.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    panel(s, 0.45 + col * 4.85, 1.15 + row * 1.2, 4.65, 1.0);
    blueLabel(s, f.title, 0.62 + col * 4.85, 1.22 + row * 1.2, 4.3);
    s.addText(f.body, {
      x: 0.62 + col * 4.85, y: 1.53 + row * 1.2, w: 4.3, h: 0.52,
      fontSize: 11.5, color: C.MUTED, fontFace: "Calibri",
    });
  });
}

// ── Slide 6: Vantagens ─────────────────────────────────────────────────────
{
  const s = addSlide("Vantagens do Next.js");

  const items = [
    "Renderizacao hibrida – escolha da estrategia ideal por rota (SSR, SSG, ISR, CSR)",
    "Developer Experience (DX) excelente – zero config, hot-reload, TypeScript nativo",
    "Performance – code splitting automatico, prefetch, otimizacao de imagens/fontes",
    "Full-stack numa unica base de codigo – API Routes e Server Actions",
    "SEO nativo – conteudo indexavel via SSR/SSG, ao contrario de SPA puras",
    "Escalabilidade – usado por Netflix, TikTok, Hulu, GitHub em producao",
    "Ecossistema maduro – comunidade React, documentacao extensa, integracao Vercel",
  ];

  items.forEach((item, i) => {
    panel(s, 0.45, 1.15 + i * 0.59, 9.1, 0.48);
    s.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.15 + i * 0.59, w: 0.12, h: 0.48, fill: { color: C.GREEN }, line: { color: C.GREEN } });
    s.addText(item, { x: 0.72, y: 1.22 + i * 0.59, w: 8.6, h: 0.36, fontSize: 12, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 7: Desvantagens ─────────────────────────────────────────────────
{
  const s = addSlide("Desvantagens e Limitacoes");

  const items = [
    "Curva de aprendizagem – App Router e React Server Components exigem mudanca de paradigma",
    "Complexidade crescente – fronteira servidor/cliente pode ser dificil de gerir em projetos grandes",
    "Dependencia da Vercel – interesses comerciais podem influenciar a direcao do framework",
    "Cold starts em serverless – latencia inicial em plataformas sem instancias aquecidas",
    "Build times – projetos com muitas paginas estaticas podem ter compilacoes demoradas",
    "Vendor lock-in potencial – funcionalidades avancadas otimizadas para infraestrutura Vercel",
  ];

  items.forEach((item, i) => {
    panel(s, 0.45, 1.15 + i * 0.69, 9.1, 0.56);
    s.addShape(prs.ShapeType.rect, { x: 0.45, y: 1.15 + i * 0.69, w: 0.12, h: 0.56, fill: { color: C.RED }, line: { color: C.RED } });
    s.addText(item, { x: 0.72, y: 1.23 + i * 0.69, w: 8.6, h: 0.4, fontSize: 12, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 8: Areas de Utilizacao ──────────────────────────────────────────
{
  const s = addSlide("Principais Areas de Utilizacao");

  const areas = [
    { area: "E-commerce",           desc: "SSG + ISR para paginas de produto de alta performance (Shopify, Nike)" },
    { area: "Plataformas SaaS",     desc: "App Router e RSC para dashboards e interfaces complexas" },
    { area: "Portais de Conteudo",  desc: "SSG para blogs, jornais e sites de documentacao" },
    { area: "APIs Full-Stack",      desc: "API Routes para backends simples sem infraestrutura adicional" },
    { area: "Dashboards de Dados",  desc: "Renderizacao hibrida para dados ao vivo com ISR e polling" },
    { area: "Aplicacoes Internas",  desc: "Developer Experience e TypeScript nativo aceleram o desenvolvimento" },
  ];

  areas.forEach((a, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    panel(s, 0.45 + col * 4.85, 1.15 + row * 1.25, 4.65, 1.05);
    blueLabel(s, a.area, 0.62 + col * 4.85, 1.22 + row * 1.25, 4.3);
    s.addText(a.desc, { x: 0.62 + col * 4.85, y: 1.54 + row * 1.25, w: 4.3, h: 0.55, fontSize: 11.5, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 9: Comparacao com Alternativas ──────────────────────────────────
{
  const s = addSlide("Comparacao com Alternativas");

  const headers = ["Framework", "Base", "Renderizacao", "Full-Stack", "Ponto Forte"];
  const rows = [
    ["Next.js",   "React",  "SSR/SSG/ISR/CSR", "Sim",     "Flexibilidade e DX"],
    ["Nuxt.js",   "Vue 3",  "SSR/SSG/CSR",     "Sim",     "Ecossistema Vue"],
    ["SvelteKit", "Svelte", "SSR/SSG/CSR",     "Sim",     "Performance / bundle reduzido"],
    ["Remix",     "React",  "SSR",             "Sim",     "Web standards"],
    ["Astro",     "Multi",  "SSG/MPA",         "Parcial", "Sites de conteudo estatico"],
  ];

  const colW = [1.7, 1.2, 2.2, 1.2, 2.2];
  const startX = 0.45;
  let cx = startX;

  // Header row
  headers.forEach((h, i) => {
    s.addShape(prs.ShapeType.rect, { x: cx, y: 1.15, w: colW[i], h: 0.42, fill: { color: C.BLUE }, line: { color: C.BLUE } });
    s.addText(h, { x: cx + 0.06, y: 1.2, w: colW[i] - 0.1, h: 0.32, fontSize: 11, bold: true, color: "FFFFFF", fontFace: "Calibri" });
    cx += colW[i];
  });

  rows.forEach((row, ri) => {
    cx = startX;
    const bg = ri === 0 ? C.BLUE_LT : (ri % 2 === 0 ? C.PANEL : C.BG);
    row.forEach((cell, ci) => {
      s.addShape(prs.ShapeType.rect, { x: cx, y: 1.57 + ri * 0.6, w: colW[ci], h: 0.58, fill: { color: bg }, line: { color: C.BORDER, width: 0.3 } });
      s.addText(cell, { x: cx + 0.06, y: 1.63 + ri * 0.6, w: colW[ci] - 0.1, h: 0.44, fontSize: ri === 0 ? 11 : 11, bold: ri === 0, color: ri === 0 ? C.BLUE : C.TEXT, fontFace: "Calibri" });
      cx += colW[ci];
    });
  });

  s.addText("Next.js destacado a azul", { x: 0.45, y: 4.65, w: 9.1, h: 0.28, fontSize: 10, italic: true, color: C.MUTED, fontFace: "Calibri" });
}

// ── Slide 10: Aplicacao de Demonstracao ───────────────────────────────────
{
  const s = addSlide("Market Dashboard", "Aplicacao desenvolvida com Next.js");

  panel(s, 0.45, 1.15, 9.1, 0.62, C.BLUE_LT);
  s.addText("Plataforma web de visualizacao de dados financeiros em tempo real, desenvolvida para demonstrar as capacidades do Next.js em contexto de producao.", {
    x: 0.62, y: 1.22, w: 8.8, h: 0.5, fontSize: 12.5, color: C.TEXT, fontFace: "Calibri",
  });

  const feats = [
    { label: "Dashboard de Mercado",    body: "Lista de 6 ativos (BTC, ETH, AAPL, MSFT, TSLA, NVDA) com precos ao vivo e filtro de pesquisa" },
    { label: "Paginas de Detalhe",      body: "Rotas dinamicas /market/[ticker] com metricas de abertura, maximo, minimo e volume" },
    { label: "Grafico de Candlestick",  body: "Timeframes 1m a 1W; scroll infinito para historico; polling a 15 s via API Route" },
    { label: "Indicadores Tecnicos",    body: "MA 20, MA 50, Bandas de Bollinger (20 per., sigma=2); calculados no cliente" },
  ];

  feats.forEach((f, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    panel(s, 0.45 + col * 4.85, 1.95 + row * 1.3, 4.65, 1.1);
    blueLabel(s, f.label, 0.62 + col * 4.85, 2.02 + row * 1.3, 4.3);
    s.addText(f.body, { x: 0.62 + col * 4.85, y: 2.36 + row * 1.3, w: 4.3, h: 0.6, fontSize: 11, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 11: Funcionalidades Next.js na Pratica ───────────────────────────
{
  const s = addSlide("Funcionalidades do Next.js na Pratica", "Como o framework foi utilizado na aplicacao");

  const map = [
    { feature: "App Router + RSC",       usage: "Paginas /market e /market/[ticker] sao Server Components – sem JS extra no cliente" },
    { feature: "API Routes",             usage: "GET /api/stocks e GET /api/ohlcv/[ticker] – proxy para Yahoo Finance, sem backend separado" },
    { feature: "ISR (revalidate: 30 s)", usage: "Cache das respostas de preco; invalidacao automatica sem re-compilar" },
    { feature: "next/dynamic (ssr:false)", usage: "lightweight-charts (acede ao DOM) importado dinamicamente; evita erros de hidratacao" },
    { feature: "generateMetadata",       usage: "Metadados dinamicos por ticker para SEO" },
    { feature: "notFound()",             usage: "Redireciona tickers invalidos para pagina 404 nativa do Next.js" },
  ];

  map.forEach((m, i) => {
    panel(s, 0.45, 1.15 + i * 0.7, 9.1, 0.58);
    s.addText(m.feature, { x: 0.62, y: 1.22 + i * 0.7, w: 3.0, h: 0.44, fontSize: 11.5, bold: true, color: C.BLUE, fontFace: "Calibri", fontFace: "Courier New" });
    s.addText(m.usage, { x: 3.7, y: 1.22 + i * 0.7, w: 5.7, h: 0.44, fontSize: 11.5, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 12: Dificuldades e Limitacoes ───────────────────────────────────
{
  const s = addSlide("Dificuldades e Limitacoes");

  panel(s, 0.45, 1.15, 4.55, 3.92);
  s.addText("Dificuldades", { x: 0.62, y: 1.22, w: 4.2, h: 0.38, fontSize: 14, bold: true, color: C.TEXT, fontFace: "Calibri" });
  [
    "API nao oficial do Yahoo Finance – campos nulos, risco de indisponibilidade",
    "lightweight-charts v5 – API alterada, identificada via codigo-fonte",
    "SSR vs. cliente – next/dynamic necessario para bibliotecas DOM",
    "Condicoes de corrida – polling + scroll + troca de TF em simultaneo",
  ].forEach((d, i) => {
    s.addShape(prs.ShapeType.rect, { x: 0.62, y: 1.7 + i * 0.82, w: 0.1, h: 0.5, fill: { color: C.BLUE }, line: { color: C.BLUE } });
    s.addText(d, { x: 0.82, y: 1.72 + i * 0.82, w: 3.95, h: 0.5, fontSize: 11, color: C.TEXT, fontFace: "Calibri" });
  });

  panel(s, 5.1, 1.15, 4.45, 3.92);
  s.addText("Limitacoes", { x: 5.28, y: 1.22, w: 4.0, h: 0.38, fontSize: 14, bold: true, color: C.TEXT, fontFace: "Calibri" });
  [
    "API nao oficial – risco de quebra sem aviso",
    "6 tickers fixos – sem pesquisa livre",
    "Polling 15 s – nao e WebSocket genuino",
    "Sem autenticacao nem portfolio",
    "Historico 1m limitado a 7 dias",
  ].forEach((l, i) => {
    s.addShape(prs.ShapeType.rect, { x: 5.28, y: 1.72 + i * 0.68, w: 0.1, h: 0.4, fill: { color: C.RED }, line: { color: C.RED } });
    s.addText(l, { x: 5.5, y: 1.74 + i * 0.68, w: 3.85, h: 0.38, fontSize: 11, color: C.TEXT, fontFace: "Calibri" });
  });
}

// ── Slide 13: Conclusoes e Feedback ───────────────────────────────────────
{
  const s = addSlide("Conclusoes e Feedback sobre o Next.js");

  panel(s, 0.45, 1.15, 9.1, 1.05, C.BLUE_LT);
  s.addText(
    "O Next.js revela-se uma escolha solida para aplicacoes web modernas: a renderizacao hibrida, o modelo full-stack e a Developer Experience justificam a sua adocao crescente. A curva de aprendizagem do App Router e real, mas o investimento e recompensado pela qualidade do produto final.",
    { x: 0.62, y: 1.22, w: 8.8, h: 0.9, fontSize: 13, color: C.TEXT, fontFace: "Calibri" }
  );

  const items = [
    { label: "Pontos fortes",          body: "Flexibilidade, DX, full-stack, performance, SEO" },
    { label: "Pontos a melhorar",      body: "Curva de aprendizagem RSC, dependencia Vercel" },
    { label: "Recomendacao de uso",    body: "Adequado para e-commerce, SaaS, dashboards, portais de conteudo" },
    { label: "Proximos passos",        body: "WebSocket, API oficial, mais indicadores, portfolio de utilizador" },
  ];

  items.forEach((it, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    panel(s, 0.45 + col * 4.85, 2.38 + row * 1.3, 4.65, 1.1);
    blueLabel(s, it.label, 0.62 + col * 4.85, 2.44 + row * 1.3, 4.3);
    s.addText(it.body, { x: 0.62 + col * 4.85, y: 2.78 + row * 1.3, w: 4.3, h: 0.58, fontSize: 11.5, color: C.MUTED, fontFace: "Calibri" });
  });
}

// ── Slide 14: Questoes ──────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C.BG };

  s.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 10, h: 0.08, fill: { color: C.BLUE }, line: { color: C.BLUE } });
  s.addShape(prs.ShapeType.rect, { x: 0, y: 5.42, w: 10, h: 0.04, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  s.addText("Questoes e Debate", {
    x: 0.5, y: 1.6, w: 9.0, h: 0.8,
    fontSize: 34, bold: true, color: C.BLACK, fontFace: "Calibri", align: "center",
  });
  s.addText("Obrigado pela atencao", {
    x: 0.5, y: 2.55, w: 9.0, h: 0.5,
    fontSize: 18, color: C.MUTED, fontFace: "Calibri", align: "center", italic: true,
  });

  s.addShape(prs.ShapeType.rect, { x: 3.0, y: 3.2, w: 4.0, h: 0.03, fill: { color: C.BORDER }, line: { color: C.BORDER } });

  STUDENTS.forEach((st, i) => {
    s.addText(`${st.name}  |  ${st.id}`, {
      x: 0.5, y: 3.42 + i * 0.45, w: 9.0, h: 0.38,
      fontSize: 13, color: C.TEXT, fontFace: "Calibri", align: "center",
    });
  });
  s.addText(`${COURSE}  |  ${YEAR}`, {
    x: 0.5, y: 4.45, w: 9.0, h: 0.3,
    fontSize: 11, color: C.MUTED, fontFace: "Calibri", align: "center",
  });
}

prs.writeFile({ fileName: "docs/apresentacao.pptx" }).then(() => {
  console.log("OK  docs/apresentacao.pptx");
}).catch(e => console.error("ERRO PPTX:", e.message));
