/**
 * Generates report.docx and presentation.pptx for the AWP project.
 * Run:  node generate-docs.mjs
 */

import {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  AlignmentType, PageBreak, TableOfContents,
  BorderStyle, Table, TableRow, TableCell,
  WidthType, ShadingType, Header, Footer,
  PageNumber, NumberFormat, Tab, TabStopType, TabStopLeader,
  LevelFormat, convertInchesToTwip, LineRuleType,
  UnderlineType,
} from "docx";
import pptxgen from "pptxgenjs";
import { writeFileSync } from "fs";
import { mkdirSync } from "fs";

mkdirSync("docs", { recursive: true });

// ─────────────────────────────────────────────────────────────────────────────
// Shared content
// ─────────────────────────────────────────────────────────────────────────────

const TITLE    = "Market Dashboard";
const SUBTITLE = "Plataforma Web para Visualização de Dados Financeiros em Tempo Real";
const COURSE   = "Advanced Web Programming";
const YEAR     = "2025/2026 – 1.º Semestre";
const UNIV     = "[Nome da Universidade]";
const DEPT     = "[Departamento / Escola]";
const PROF     = "[Nome do/a Docente]";

const STUDENTS = [
  { name: "[Nome do Estudante 1]", id: "Nº XXXXXX" },
  { name: "[Nome do Estudante 2]", id: "Nº XXXXXX" },
];

// ─────────────────────────────────────────────────────────────────────────────
// DOCX helpers
// ─────────────────────────────────────────────────────────────────────────────

const FONT = "Times New Roman";
const FONT_SIZE = 24; // half-points → 12pt
const LINE_RULE = { rule: LineRuleType.AUTO, value: 276 }; // 1.15x spacing

function normalPara(text, opts = {}) {
  return new Paragraph({
    spacing: { after: 200, line: 276, lineRule: LineRuleType.AUTO },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, ...opts })],
  });
}

function boldPara(text, opts = {}) {
  return normalPara(text, { bold: true, ...opts });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 200 },
    children: [new TextRun({ text, font: FONT, size: 28, bold: true })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 300, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true })],
  });
}

function heading3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 120 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, bold: true, italics: true })],
  });
}

function bullet(text, level = 0) {
  return new Paragraph({
    bullet: { level },
    spacing: { after: 100, line: 276, lineRule: LineRuleType.AUTO },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE })],
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function emptyLine() {
  return new Paragraph({ spacing: { after: 0 }, children: [new TextRun("")]});
}

function centeredPara(text, opts = {}) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 160 },
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE, ...opts })],
  });
}

function bibEntry(text) {
  return new Paragraph({
    indent: { left: 720, hanging: 720 },
    spacing: { after: 200, line: 276, lineRule: LineRuleType.AUTO },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: FONT_SIZE })],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// DOCX content sections
// ─────────────────────────────────────────────────────────────────────────────

const COVER = [
  emptyLine(), emptyLine(), emptyLine(),
  centeredPara(UNIV, { bold: true, size: 28 }),
  centeredPara(DEPT),
  emptyLine(), emptyLine(),
  centeredPara(TITLE, { bold: true, size: 36 }),
  centeredPara(SUBTITLE, { italics: true, size: 26 }),
  emptyLine(), emptyLine(),
  centeredPara("Trabalho Prático da Unidade Curricular de"),
  centeredPara(COURSE, { bold: true }),
  emptyLine(), emptyLine(),
  ...STUDENTS.flatMap(s => [
    centeredPara(`${s.name}  ·  ${s.id}`),
  ]),
  emptyLine(),
  centeredPara(`Docente: ${PROF}`),
  emptyLine(), emptyLine(),
  centeredPara(YEAR),
  centeredPara("Maio de 2026"),
  pageBreak(),
];

const ABSTRACT = [
  heading1("Resumo"),
  normalPara(
    "O presente relatório descreve o desenvolvimento de uma aplicação web de visualização de dados financeiros em tempo real, denominada Market Dashboard. A plataforma permite a consulta de preços de ativos financeiros — incluindo criptomoedas (Bitcoin e Ethereum) e ações (AAPL, MSFT, TSLA, NVDA) — com recurso a gráficos de velas japonesas (candlestick), seleção de intervalos temporais, paginação de dados históricos por scroll infinito e atualização automática a cada quinze segundos. O projeto foi desenvolvido com Next.js 16, React 19, TypeScript, Tailwind CSS v4 e a biblioteca lightweight-charts v5 da TradingView, integrando a API não oficial do Yahoo Finance como fonte de dados de mercado."
  ),
  emptyLine(),
  boldPara("Palavras-chave: "),
  normalPara("aplicação web, visualização financeira, dados em tempo real, Next.js, candlestick, indicadores técnicos."),
  pageBreak(),
];

// TOC — native Word TOC field
const TOC_SECTION = [
  heading1("Índice"),
  new Paragraph({
    spacing: { after: 200 },
    children: [
      new TableOfContents("Índice", {
        hyperlink: true,
        headingStyleRange: "1-3",
        stylesWithLevels: [
          { styleName: "Heading 1", level: 1 },
          { styleName: "Heading 2", level: 2 },
          { styleName: "Heading 3", level: 3 },
        ],
      }),
    ],
  }),
  pageBreak(),
];

const FIGURES = [
  heading1("Lista de Figuras e Tabelas"),
  emptyLine(),
  boldPara("Figuras"),
  normalPara("Figura 1 – Página principal do Market Dashboard (tabela de ativos) ........... 6"),
  normalPara("Figura 2 – Página de detalhe de um ticker (métricas e gráfico de velas) ..... 7"),
  normalPara("Figura 3 – Gráfico candlestick com indicadores MA 20, MA 50 e BB 20 ......... 8"),
  normalPara("Figura 4 – Diagrama de arquitetura da aplicação .............................. 9"),
  emptyLine(),
  boldPara("Tabelas"),
  normalPara("Tabela 1 – Tecnologias utilizadas e respetivas versões ....................... 5"),
  normalPara("Tabela 2 – Endpoints da API interna .......................................... 7"),
  pageBreak(),
];

const INTRO = [
  heading1("1. Introdução"),
  normalPara(
    "A análise de mercados financeiros tem-se tornado cada vez mais acessível ao público em geral, graças à proliferação de plataformas digitais que disponibilizam dados de cotações em tempo real. No entanto, o desenvolvimento de tais plataformas apresenta desafios técnicos relevantes, nomeadamente ao nível da integração com fontes de dados externas, da atualização contínua da interface sem degradação de desempenho e da apresentação visual de informação financeira complexa de forma intuitiva."
  ),
  normalPara(
    "O presente projeto, inserido na unidade curricular de Advanced Web Programming (AWP), propõe-se a desenvolver uma aplicação web denominada Market Dashboard, que permite a consulta e visualização de dados financeiros de um conjunto de ativos selecionados. A aplicação oferece ao utilizador uma visão consolidada do estado atual do mercado, bem como ferramentas de análise técnica através de gráficos de candlestick com indicadores sobrepostos."
  ),
  heading2("1.1 Objetivos"),
  bullet("Implementar uma dashboard de mercado que exiba preços atuais, variações e volumes de ativos financeiros;"),
  bullet("Integrar a API do Yahoo Finance como fonte de dados em tempo real;"),
  bullet("Desenvolver gráficos de candlestick interativos com seleção de timeframes (1m a 1W);"),
  bullet("Implementar indicadores técnicos configuráveis pelo utilizador (MA 20, MA 50, Bandas de Bollinger);"),
  bullet("Garantir a atualização automática dos dados sem recarga da página (polling a 15 segundos);"),
  bullet("Suportar carregamento de dados históricos adicionais através de scroll infinito."),
  heading2("1.2 Estrutura do Relatório"),
  normalPara(
    "O presente relatório está organizado da seguinte forma: a Secção 2 descreve a metodologia de resolução do problema, incluindo a análise de requisitos e a seleção tecnológica; a Secção 3 detalha a implementação da aplicação, com foco na arquitetura e nos componentes desenvolvidos; a Secção 4 apresenta as conclusões, dificuldades encontradas, limitações e propostas de melhoria; por fim, a Secção 5 lista as referências bibliográficas utilizadas."
  ),
  pageBreak(),
];

const METHOD = [
  heading1("2. Metodologia de Resolução do Problema"),
  heading2("2.1 Análise de Requisitos"),
  normalPara(
    "A fase inicial do projeto consistiu na identificação e especificação dos requisitos funcionais e não funcionais da aplicação. Os requisitos funcionais incluem a visualização de uma lista de ativos com dados em tempo real, a navegação para páginas de detalhe por ticker, a apresentação de gráficos de candlestick com múltiplos timeframes e a sobreposição de indicadores técnicos. Os requisitos não funcionais incidem sobre o desempenho (carregamento inicial rápido, atualizações fluidas), a responsividade da interface e a robustez face a falhas da API externa."
  ),
  heading2("2.2 Seleção Tecnológica"),
  normalPara(
    "A escolha das tecnologias foi orientada pelos requisitos identificados e pelos conhecimentos adquiridos na unidade curricular. A Tabela 1 sintetiza as tecnologias adotadas e as respetivas versões."
  ),
  emptyLine(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({ children: [boldPara("Tecnologia")], width: { size: 30, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [boldPara("Versão")], width: { size: 20, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [boldPara("Papel no Projeto")], width: { size: 50, type: WidthType.PERCENTAGE } }),
        ],
      }),
      ...([
        ["Next.js", "16.2.6", "Framework React com App Router, SSR e API Routes"],
        ["React", "19.2.4", "Biblioteca de interface de utilizador baseada em componentes"],
        ["TypeScript", "5.x", "Tipagem estática sobre JavaScript"],
        ["Tailwind CSS", "4.x", "Framework CSS utilitário para estilização"],
        ["lightweight-charts", "5.2.0", "Biblioteca de gráficos financeiros (TradingView)"],
        ["Yahoo Finance API", "Não oficial", "Fonte de dados de mercado (REST)"],
      ].map(([tech, ver, role]) =>
        new TableRow({
          children: [
            new TableCell({ children: [normalPara(tech)] }),
            new TableCell({ children: [normalPara(ver)] }),
            new TableCell({ children: [normalPara(role)] }),
          ],
        })
      )),
    ],
  }),
  emptyLine(),
  centeredPara("Tabela 1 – Tecnologias utilizadas e respetivas versões", { italics: true, size: 20 }),
  emptyLine(),
  normalPara(
    "O Next.js foi escolhido por combinar renderização do lado do servidor (SSR) com componentes de cliente interativos, permitindo uma boa performance inicial e atualizações em tempo real. A biblioteca lightweight-charts da TradingView foi preferida em detrimento de alternativas como Chart.js por ser especificamente otimizada para gráficos financeiros de alta performance, suportando candlestick nativo e atualizações incrementais eficientes (Vercel, Inc., 2026; TradingView, Inc., 2026)."
  ),
  heading2("2.3 Metodologia de Desenvolvimento"),
  normalPara(
    "O desenvolvimento seguiu uma abordagem iterativa e incremental, organizada em três fases principais:"
  ),
  bullet("Fase 1 – Dashboard base: implementação da página de listagem de ativos com dados em tempo real provenientes da API do Yahoo Finance;"),
  bullet("Fase 2 – Página de detalhe: desenvolvimento das páginas individuais por ticker com métricas de mercado (abertura, máximo, mínimo, volume);"),
  bullet("Fase 3 – Visualização avançada: integração dos gráficos de candlestick com seleção de timeframe, scroll infinito, polling e indicadores técnicos configuráveis."),
  normalPara(
    "Em cada iteração procedeu-se à análise de requisitos, implementação, teste em ambiente de desenvolvimento e refinamento com base no feedback obtido."
  ),
  pageBreak(),
];

const IMPL = [
  heading1("3. Descrição da Implementação"),
  heading2("3.1 Arquitetura da Aplicação"),
  normalPara(
    "A aplicação segue a arquitetura App Router do Next.js, distinguindo componentes de servidor (que executam exclusivamente no servidor e produzem HTML estático ou dinâmico) de componentes de cliente (que hidratam no browser e suportam interatividade). Esta separação permite otimizar o bundle JavaScript enviado ao cliente e garantir um carregamento inicial rápido."
  ),
  normalPara(
    "A estrutura de ficheiros respeita as convenções do Next.js 16:"
  ),
  bullet("app/page.tsx – Redireciona para /market;"),
  bullet("app/market/page.tsx – Dashboard principal com lista de ativos;"),
  bullet("app/market/[ticker]/page.tsx – Página de detalhe com gráfico;"),
  bullet("app/api/stocks/route.ts – Endpoint REST para preços atuais;"),
  bullet("app/api/ohlcv/[ticker]/route.ts – Endpoint REST para dados OHLCV históricos;"),
  bullet("components/filter-input.tsx – Campo de pesquisa e filtro;"),
  bullet("components/candlestick-chart.tsx – Wrapper com ssr: false;"),
  bullet("components/candlestick-chart-inner.tsx – Componente de gráfico (cliente)."),
  heading2("3.2 Camada de API"),
  normalPara(
    "Foram implementados dois endpoints REST internos que servem de proxy entre o frontend e a API do Yahoo Finance, evitando expor diretamente ao cliente as chamadas a serviços externos e permitindo aplicar lógica de validação e tratamento de erros."
  ),
  emptyLine(),
  new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          new TableCell({ children: [boldPara("Endpoint")], width: { size: 35, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [boldPara("Parâmetros")], width: { size: 30, type: WidthType.PERCENTAGE } }),
          new TableCell({ children: [boldPara("Descrição")], width: { size: 35, type: WidthType.PERCENTAGE } }),
        ],
      }),
      ...[
        ["GET /api/stocks", "–", "Retorna preços atuais dos 6 ativos com fallback para dados simulados"],
        ["GET /api/ohlcv/[ticker]", "interval, before", "Retorna ~400 velas OHLCV do Yahoo Finance para o ticker e timeframe indicados"],
      ].map(([ep, params, desc]) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ep, font: "Courier New", size: 20 })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: params, font: "Courier New", size: 20 })] })] }),
            new TableCell({ children: [normalPara(desc)] }),
          ],
        })
      ),
    ],
  }),
  emptyLine(),
  centeredPara("Tabela 2 – Endpoints da API interna", { italics: true, size: 20 }),
  emptyLine(),
  normalPara(
    "O endpoint /api/ohlcv/[ticker] aceita o parâmetro interval (e.g., 1m, 5m, 60m, 1d, 1wk) e um parâmetro opcional before (timestamp Unix) para paginação retroativa. Quando before está ausente, o período de consulta termina no instante atual; quando presente, é utilizado como período2, permitindo carregar dados mais antigos sem sobreposição."
  ),
  heading2("3.3 Gráfico de Candlestick"),
  normalPara(
    "O componente principal de visualização é o CandlestickChartInner, um componente de cliente que utiliza a biblioteca lightweight-charts v5 da TradingView (TradingView, Inc., 2026). A sua arquitetura é baseada em quatro efeitos React independentes:"
  ),
  heading3("3.3.1 Inicialização do Gráfico"),
  normalPara(
    "No primeiro useEffect (dependências vazias), o gráfico é criado imperativemente sobre um elemento div através de createChart(). São configuradas as séries de candlestick e de volume (histograma), bem como um ResizeObserver para ajuste dinâmico da largura ao redimensionar a janela. A abordagem de importação estática combinada com next/dynamic (ssr: false) garante que a biblioteca, que acede ao DOM na inicialização, nunca é carregada durante a renderização do lado do servidor."
  ),
  heading3("3.3.2 Carregamento de Dados por Timeframe"),
  normalPara(
    "O segundo useEffect executa sempre que o timeframe selecionado muda, realizando uma chamada ao endpoint /api/ohlcv e populando as séries com os dados recebidos. Os indicadores técnicos ativos são recriados após cada carregamento."
  ),
  heading3("3.3.3 Paginação por Scroll"),
  normalPara(
    "Um handler subscrito em subscribeVisibleLogicalRangeChange deteta quando o utilizador desloca o gráfico para a esquerda (range.from < 10). Nesse caso, é desencadeada uma chamada ao endpoint com o parâmetro before igual ao timestamp da vela mais antiga em memória, sendo os dados mais antigos fundidos com os existentes após deduplicação e reordenação."
  ),
  heading3("3.3.4 Atualização em Tempo Real (Polling)"),
  normalPara(
    "Um setInterval de 15 segundos executa periodicamente uma chamada ao endpoint. O código filtra todas as velas com timestamp ≥ ao da última vela em memória, atualizando a vela em curso e/ou anexando novas velas. Uma referência pollInFlightRef previne concorrência entre chamadas consecutivas. Os indicadores são re-calculados após cada atualização."
  ),
  heading2("3.4 Indicadores Técnicos"),
  normalPara(
    "O sistema de indicadores é gerido pela função syncIndicators, que cria, atualiza ou remove séries auxiliares (LineSeries) consoante o conjunto de indicadores ativos. Estão disponíveis quatro indicadores:"
  ),
  bullet("Volume (Vol) – histograma de volume, ativo por omissão, pode ser ocultado;"),
  bullet("MA 20 – Média Móvel Simples de 20 períodos (cor ciano), calculada sobre o campo close das velas em memória;"),
  bullet("MA 50 – Média Móvel Simples de 50 períodos (cor âmbar);"),
  bullet("BB 20 – Bandas de Bollinger de 20 períodos com desvio-padrão 2 (cor índigo, linha tracejada), compostas por banda superior e inferior."),
  normalPara(
    "Os cálculos são realizados inteiramente no cliente, sem chamadas adicionais ao servidor, o que minimiza a latência e a carga na API."
  ),
  pageBreak(),
];

const CONCLUSIONS = [
  heading1("4. Conclusões"),
  heading2("4.1 Dificuldades Encontradas"),
  normalPara(
    "Ao longo do desenvolvimento, foram encontradas várias dificuldades de natureza técnica:"
  ),
  bullet("API do Yahoo Finance: A API utilizada é não oficial e não documentada, o que implicou análise empírica do formato de resposta, tratamento de campos nulos em séries OHLCV e implementação de dados de fallback para garantir a robustez da aplicação em caso de falha;"),
  bullet("Next.js 16 e componentes SSR/cliente: A distinção entre componentes de servidor e de cliente introduziu complexidade na gestão do ciclo de vida dos componentes. A biblioteca lightweight-charts acede ao DOM na importação, pelo que foi necessário recorrer a next/dynamic com ssr: false para evitar erros de hidratação;"),
  bullet("Migração de API do lightweight-charts v5: A versão 5 da biblioteca introduziu alterações de compatibilidade, nomeadamente a substituição de métodos como addCandlestickSeries() pela nova assinatura addSeries(CandlestickSeries, options), o que exigiu análise do código-fonte e dos types da biblioteca;"),
  bullet("Gestão de estado assíncrono: A coexistência de polling, paginação por scroll e alterações de timeframe exigiu o uso cuidadoso de refs mutáveis (useRef) para evitar closures obsoletas e condições de corrida, em particular no que respeita à variável pollInFlightRef."),
  heading2("4.2 Limitações"),
  bullet("Fonte de dados não oficial: A dependência de uma API não oficial do Yahoo Finance constitui um risco de disponibilidade, pois a sua estrutura ou disponibilidade pode ser alterada sem aviso;"),
  bullet("Conjunto fixo de ativos: A aplicação está limitada a seis ativos predefinidos, não permitindo ao utilizador pesquisar ou adicionar outros tickers;"),
  bullet("Polling vs. tempo real: A atualização a cada 15 segundos baseia-se em polling HTTP, o que introduz latência comparativamente a soluções baseadas em WebSocket ou Server-Sent Events;"),
  bullet("Ausência de persistência: Não existe qualquer camada de autenticação, perfis de utilizador ou gestão de portfólio;"),
  bullet("Disponibilidade de dados históricos: A API do Yahoo Finance impõe limites nos dados históricos disponíveis por intervalo (e.g., dados de 1 minuto apenas nos últimos 7 dias)."),
  heading2("4.3 Alternativas e Melhorias"),
  bullet("WebSocket / Server-Sent Events: Substituição do polling por um canal de comunicação persistente para reduzir latência e carga no servidor;"),
  bullet("API oficial: Integração com APIs financeiras oficiais (e.g., Alpha Vantage, Polygon.io, Finnhub) para maior fiabilidade e mais dados históricos;"),
  bullet("Pesquisa de tickers: Implementação de um campo de pesquisa que permita ao utilizador consultar qualquer ativo suportado pela fonte de dados;"),
  bullet("Indicadores adicionais: RSI, MACD, Estocástico, Ichimoku Cloud e outros indicadores de análise técnica;"),
  bullet("Gestão de portfólio: Autenticação de utilizadores, watchlists personalizadas, alertas de preço e registo de transações;"),
  bullet("Testes automatizados: Implementação de testes unitários (Jest/Vitest) e de integração (Playwright/Cypress) para garantir a estabilidade da aplicação."),
  heading2("4.4 Síntese"),
  normalPara(
    "O projeto Market Dashboard cumpre os objetivos propostos, demonstrando a viabilidade de construir uma plataforma de visualização financeira moderna com tecnologias web contemporâneas. A integração de dados em tempo real, gráficos interativos e indicadores técnicos num framework fullstack como o Next.js revelou-se uma abordagem eficaz, ainda que com os condicionalismos identificados. As melhorias propostas constituem uma base sólida para evoluções futuras da plataforma."
  ),
  pageBreak(),
];

const BIBLIOGRAPHY = [
  heading1("5. Bibliografia"),
  normalPara("Referências formatadas segundo a norma APA 7.ª edição."),
  emptyLine(),
  bibEntry("Fielding, R. T. (2000). Architectural styles and the design of network-based software architectures [Dissertação de doutoramento, University of California, Irvine]. https://roy.gbiv.com/pubs/dissertation/top.htm"),
  bibEntry("Meta Platforms, Inc. (2026). React: A JavaScript library for building user interfaces. https://react.dev"),
  bibEntry("Microsoft. (2012–2026). TypeScript: JavaScript with syntax for types. https://www.typescriptlang.org"),
  bibEntry("Tailwind Labs Inc. (2026). Tailwind CSS: Rapidly build modern websites without ever leaving your HTML. https://tailwindcss.com"),
  bibEntry("TradingView, Inc. (2026). Lightweight Charts™: Performant financial charts built with HTML5 canvas. https://tradingview.github.io/lightweight-charts/"),
  bibEntry("Tufte, E. R. (1983). The visual display of quantitative information. Graphics Press."),
  bibEntry("Vercel, Inc. (2026). Next.js: The React framework for production. https://nextjs.org"),
  bibEntry("Yahoo Finance. (2026). Yahoo Finance – Dados de mercado financeiro. https://finance.yahoo.com"),
];

// ─────────────────────────────────────────────────────────────────────────────
// Build and save DOCX
// ─────────────────────────────────────────────────────────────────────────────

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullet-list",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } },
      }],
    }],
  },
  styles: {
    default: {
      document: {
        run: { font: FONT, size: FONT_SIZE },
        paragraph: { spacing: { after: 200 } },
      },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 28, bold: true, font: FONT, color: "1a1a1a" },
        paragraph: { spacing: { before: 480, after: 240 } },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: "2d2d2d" },
        paragraph: { spacing: { before: 360, after: 180 } },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: FONT_SIZE, bold: true, italics: true, font: FONT, color: "3d3d3d" },
        paragraph: { spacing: { before: 240, after: 120 } },
      },
    ],
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(1),
          right: convertInchesToTwip(1),
          bottom: convertInchesToTwip(1),
          left: convertInchesToTwip(1.18), // standard Portuguese report margin
        },
      },
    },
    headers: {
      default: new Header({
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "aaaaaa" } },
            children: [new TextRun({ text: `${TITLE} – ${COURSE}`, font: FONT, size: 18, color: "666666" })],
          }),
        ],
      }),
    },
    footers: {
      default: new Footer({
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            border: { top: { style: BorderStyle.SINGLE, size: 6, color: "aaaaaa" } },
            children: [
              new TextRun({ text: "Página ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: "666666" }),
              new TextRun({ text: " de ", font: FONT, size: 18, color: "666666" }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: "666666" }),
            ],
          }),
        ],
      }),
    },
    children: [
      ...COVER,
      ...ABSTRACT,
      ...TOC_SECTION,
      ...FIGURES,
      ...INTRO,
      ...METHOD,
      ...IMPL,
      ...CONCLUSIONS,
      ...BIBLIOGRAPHY,
    ],
  }],
});

Packer.toBuffer(doc).then((buf) => {
  writeFileSync("docs/relatorio.docx", buf);
  console.log("✅  docs/relatorio.docx generated");
});

// ─────────────────────────────────────────────────────────────────────────────
// PPTX
// ─────────────────────────────────────────────────────────────────────────────

const prs = new pptxgen();

prs.layout  = "LAYOUT_WIDE";  // 16:9
prs.author  = STUDENTS.map(s => s.name).join(", ");
prs.subject = COURSE;
prs.title   = TITLE;

// Theme colours
const C_BG    = "0d0d0d";
const C_CARD  = "1a1a1a";
const C_GREEN = "10b981";
const C_ZINC  = "52525b";
const C_WHITE = "f4f4f5";
const C_SUB   = "a1a1aa";

// Slide master settings (manual dark theme per slide)
function addSlide(title, subtitle) {
  const slide = prs.addSlide();
  slide.background = { color: C_BG };
  if (title) {
    slide.addText(title, {
      x: 0.4, y: 0.3, w: "90%", h: 0.6,
      fontSize: 28, bold: true, color: C_WHITE, fontFace: "Calibri",
    });
  }
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.4, y: 0.9, w: "90%", h: 0.4,
      fontSize: 14, color: C_SUB, fontFace: "Calibri",
    });
  }
  return slide;
}

function addBodyText(slide, lines, x, y, w, h, opts = {}) {
  slide.addText(lines, {
    x, y, w, h,
    fontSize: 13, color: C_WHITE, fontFace: "Calibri",
    valign: "top",
    ...opts,
  });
}

function addCard(slide, x, y, w, h, color = C_CARD) {
  slide.addShape(prs.ShapeType.rect, {
    x, y, w, h,
    fill: { color },
    line: { color: "2d2d2d", width: 0.5 },
  });
}

// ── Slide 1: Cover ───────────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C_BG };

  // Green accent bar
  s.addShape(prs.ShapeType.rect, { x: 0, y: 4.8, w: 10, h: 0.06, fill: { color: C_GREEN }, line: { color: C_GREEN } });

  s.addText(TITLE, {
    x: 0.6, y: 1.2, w: 8.8, h: 1.0,
    fontSize: 44, bold: true, color: C_WHITE, fontFace: "Calibri", align: "center",
  });
  s.addText(SUBTITLE, {
    x: 0.6, y: 2.3, w: 8.8, h: 0.6,
    fontSize: 16, color: C_GREEN, fontFace: "Calibri", align: "center",
  });
  s.addText(COURSE + "  ·  " + YEAR, {
    x: 0.6, y: 3.1, w: 8.8, h: 0.4,
    fontSize: 13, color: C_SUB, fontFace: "Calibri", align: "center",
  });
  s.addText(STUDENTS.map(st => `${st.name}  (${st.id})`).join("        "), {
    x: 0.6, y: 3.7, w: 8.8, h: 0.4,
    fontSize: 13, color: C_WHITE, fontFace: "Calibri", align: "center",
  });
  s.addText(UNIV + "  ·  Maio 2026", {
    x: 0.6, y: 4.2, w: 8.8, h: 0.3,
    fontSize: 11, color: C_ZINC, fontFace: "Calibri", align: "center",
  });
}

// ── Slide 2: Agenda ──────────────────────────────────────────────────────────
{
  const s = addSlide("Agenda", "Estrutura da Apresentação");
  const items = [
    ["01", "Contexto e Motivação"],
    ["02", "Objetivos"],
    ["03", "Metodologia e Tecnologias"],
    ["04", "Arquitetura da Aplicação"],
    ["05", "Funcionalidades Implementadas"],
    ["06", "Dificuldades e Limitações"],
    ["07", "Conclusões e Melhorias"],
  ];
  items.forEach(([num, label], i) => {
    const col = i < 4 ? 0 : 1;
    const row = i < 4 ? i : i - 4;
    addCard(s, 0.4 + col * 4.9, 1.4 + row * 0.82, 4.6, 0.65);
    s.addText(num, { x: 0.55 + col * 4.9, y: 1.52 + row * 0.82, w: 0.5, h: 0.4, fontSize: 18, bold: true, color: C_GREEN, fontFace: "Calibri" });
    s.addText(label, { x: 1.1 + col * 4.9, y: 1.52 + row * 0.82, w: 3.8, h: 0.4, fontSize: 14, color: C_WHITE, fontFace: "Calibri" });
  });
}

// ── Slide 3: Contexto e Motivação ───────────────────────────────────────────
{
  const s = addSlide("Contexto e Motivação");
  const points = [
    "Crescente interesse no mercado financeiro por parte do público geral",
    "Necessidade de plataformas acessíveis para visualização de dados em tempo real",
    "Desafios técnicos: integração de APIs externas, performance e UX",
    "Oportunidade de aplicar conhecimentos de AWP num domínio prático e relevante",
  ];
  points.forEach((p, i) => {
    addCard(s, 0.4, 1.4 + i * 0.92, 9.2, 0.72);
    s.addText("›", { x: 0.55, y: 1.52 + i * 0.92, w: 0.3, h: 0.5, fontSize: 20, color: C_GREEN, fontFace: "Calibri" });
    s.addText(p, { x: 0.95, y: 1.52 + i * 0.92, w: 8.4, h: 0.52, fontSize: 14, color: C_WHITE, fontFace: "Calibri" });
  });
}

// ── Slide 4: Objetivos ───────────────────────────────────────────────────────
{
  const s = addSlide("Objetivos do Projeto");
  const objs = [
    ["📊", "Dashboard com preços em tempo real"],
    ["🕯️", "Gráficos candlestick interativos (1m – 1W)"],
    ["📡", "Integração com Yahoo Finance API"],
    ["📈", "Indicadores técnicos (MA20, MA50, BB)"],
    ["⏱️", "Atualização automática a cada 15 segundos"],
    ["↔️", "Scroll infinito para dados históricos"],
  ];
  objs.forEach(([icon, text], i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(s, 0.4 + col * 4.9, 1.4 + row * 1.05, 4.6, 0.85);
    s.addText(icon + "  " + text, {
      x: 0.55 + col * 4.9, y: 1.52 + row * 1.05, w: 4.3, h: 0.65,
      fontSize: 13, color: C_WHITE, fontFace: "Calibri",
    });
  });
}

// ── Slide 5: Tecnologias ─────────────────────────────────────────────────────
{
  const s = addSlide("Stack Tecnológico");
  const techs = [
    { name: "Next.js 16", role: "Framework fullstack (App Router, SSR, API Routes)", color: "ffffff" },
    { name: "React 19", role: "Biblioteca de componentes de UI", color: "61dafb" },
    { name: "TypeScript", role: "Tipagem estática e robustez de código", color: "3178c6" },
    { name: "Tailwind CSS v4", role: "Estilização utilitária e tema dark", color: "38bdf8" },
    { name: "lightweight-charts v5", role: "Gráficos financeiros de alta performance", color: C_GREEN },
    { name: "Yahoo Finance API", role: "Fonte de dados de mercado (não oficial)", color: "7c3aed" },
  ];
  techs.forEach((t, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    addCard(s, 0.4 + col * 4.9, 1.4 + row * 1.05, 4.6, 0.85);
    s.addText(t.name, { x: 0.55 + col * 4.9, y: 1.45 + row * 1.05, w: 4.3, h: 0.38, fontSize: 14, bold: true, color: t.color, fontFace: "Calibri" });
    s.addText(t.role, { x: 0.55 + col * 4.9, y: 1.82 + row * 1.05, w: 4.3, h: 0.38, fontSize: 11, color: C_SUB, fontFace: "Calibri" });
  });
}

// ── Slide 6: Arquitetura ─────────────────────────────────────────────────────
{
  const s = addSlide("Arquitetura da Aplicação");

  // Layers
  const layers = [
    { label: "Browser", items: ["Dashboard", "Ticker Detail", "Candlestick Chart"], color: "1e3a5f" },
    { label: "Next.js Server", items: ["/api/stocks", "/api/ohlcv/[ticker]"], color: "1a3a1a" },
    { label: "Yahoo Finance", items: ["query1.finance.yahoo.com"], color: "3a1a1a" },
  ];
  layers.forEach((layer, i) => {
    addCard(s, 0.3, 1.35 + i * 1.2, 9.4, 1.0, layer.color);
    s.addText(layer.label, { x: 0.45, y: 1.42 + i * 1.2, w: 1.8, h: 0.35, fontSize: 12, bold: true, color: C_WHITE, fontFace: "Calibri" });
    s.addText(layer.items.join("    ·    "), { x: 0.45, y: 1.77 + i * 1.2, w: 9.0, h: 0.35, fontSize: 12, color: C_SUB, fontFace: "Calibri" });

    // Arrow between layers
    if (i < layers.length - 1) {
      s.addText("↕", { x: 4.6, y: 2.32 + i * 1.2, w: 0.4, h: 0.3, fontSize: 16, color: C_GREEN, fontFace: "Calibri", align: "center" });
    }
  });
}

// ── Slide 7: Gráfico de Candlestick ─────────────────────────────────────────
{
  const s = addSlide("Gráfico de Candlestick", "Funcionalidade central da aplicação");
  const features = [
    ["Timeframes", "1m · 5m · 15m · 1H · 1D · 1W"],
    ["Dados iniciais", "~400 velas por timeframe"],
    ["Scroll infinito", "Carrega dados mais antigos ao deslocar para a esquerda"],
    ["Polling", "Atualização automática a cada 15 segundos"],
    ["Indicadores", "MA 20 (ciano) · MA 50 (âmbar) · BB 20 (índigo)"],
    ["Volume", "Histograma na base do gráfico (ativável/desativável)"],
  ];
  features.forEach(([label, value], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    addCard(s, 0.4 + col * 4.9, 1.4 + row * 1.0, 4.6, 0.82);
    s.addText(label, { x: 0.55 + col * 4.9, y: 1.46 + row * 1.0, w: 4.3, h: 0.32, fontSize: 12, bold: true, color: C_GREEN, fontFace: "Calibri" });
    s.addText(value, { x: 0.55 + col * 4.9, y: 1.78 + row * 1.0, w: 4.3, h: 0.32, fontSize: 11, color: C_WHITE, fontFace: "Calibri" });
  });
}

// ── Slide 8: Dificuldades ────────────────────────────────────────────────────
{
  const s = addSlide("Dificuldades Encontradas");
  const diffs = [
    "API não oficial do Yahoo Finance — necessidade de fallback e análise empírica da estrutura de resposta",
    "Gestão de componentes SSR vs. cliente — lightweight-charts acede ao DOM na importação (solução: next/dynamic + ssr: false)",
    "Migração para lightweight-charts v5 — alteração de API (addCandlestickSeries → addSeries)",
    "Condições de corrida — polling, scroll e troca de TF em simultâneo exigiram uso de refs mutáveis",
  ];
  diffs.forEach((d, i) => {
    addCard(s, 0.4, 1.4 + i * 0.9, 9.2, 0.72, "2a1a1a");
    s.addText("⚠", { x: 0.55, y: 1.5 + i * 0.9, w: 0.35, h: 0.5, fontSize: 16, color: "f59e0b", fontFace: "Calibri" });
    s.addText(d, { x: 1.0, y: 1.5 + i * 0.9, w: 8.4, h: 0.5, fontSize: 12, color: C_WHITE, fontFace: "Calibri" });
  });
}

// ── Slide 9: Limitações e Melhorias ─────────────────────────────────────────
{
  const s = addSlide("Limitações e Propostas de Melhoria");

  addCard(s, 0.4, 1.3, 4.4, 3.5, "2a1a1a");
  s.addText("Limitações", { x: 0.6, y: 1.38, w: 4.0, h: 0.4, fontSize: 15, bold: true, color: "f87171", fontFace: "Calibri" });
  const lims = ["API não oficial", "6 tickers fixos", "Polling (não WebSocket)", "Sem autenticação", "Dados históricos limitados por intervalo"];
  lims.forEach((l, i) => s.addText("• " + l, { x: 0.65, y: 1.85 + i * 0.52, w: 3.9, h: 0.45, fontSize: 12, color: C_WHITE, fontFace: "Calibri" }));

  addCard(s, 5.2, 1.3, 4.4, 3.5, "1a2a1a");
  s.addText("Melhorias", { x: 5.4, y: 1.38, w: 4.0, h: 0.4, fontSize: 15, bold: true, color: "4ade80", fontFace: "Calibri" });
  const impr = ["WebSocket / SSE em tempo real", "API oficial (Polygon.io)", "Pesquisa de qualquer ticker", "Mais indicadores (RSI, MACD)", "Portfólio e alertas de preço"];
  impr.forEach((m, i) => s.addText("• " + m, { x: 5.45, y: 1.85 + i * 0.52, w: 3.9, h: 0.45, fontSize: 12, color: C_WHITE, fontFace: "Calibri" }));
}

// ── Slide 10: Conclusões ─────────────────────────────────────────────────────
{
  const s = prs.addSlide();
  s.background = { color: C_BG };
  s.addShape(prs.ShapeType.rect, { x: 0, y: 2.35, w: 10, h: 0.06, fill: { color: C_GREEN }, line: { color: C_GREEN } });

  s.addText("Conclusões", {
    x: 0.6, y: 0.5, w: 8.8, h: 0.7,
    fontSize: 32, bold: true, color: C_WHITE, fontFace: "Calibri", align: "center",
  });
  s.addText(
    "O Market Dashboard demonstra a viabilidade de construir uma plataforma de\nvisualização financeira moderna com Next.js, React e lightweight-charts.\n\nOs objetivos propostos foram cumpridos, com integração de dados em tempo real,\ngráficos interativos e indicadores técnicos configuráveis.",
    {
      x: 0.8, y: 2.6, w: 8.4, h: 2.0,
      fontSize: 14, color: C_WHITE, fontFace: "Calibri", align: "center",
    }
  );
  s.addText("Obrigado pela atenção", {
    x: 0.6, y: 4.7, w: 8.8, h: 0.5,
    fontSize: 18, bold: true, color: C_GREEN, fontFace: "Calibri", align: "center",
  });
  s.addText(STUDENTS.map(st => `${st.name}  (${st.id})`).join("   ·   "), {
    x: 0.6, y: 5.2, w: 8.8, h: 0.35,
    fontSize: 11, color: C_ZINC, fontFace: "Calibri", align: "center",
  });
}

prs.writeFile({ fileName: "docs/apresentacao.pptx" }).then(() => {
  console.log("✅  docs/apresentacao.pptx generated");
});
