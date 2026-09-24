/* =========================================================
   Painel de Triagem | Mudança de Sede — TODOS Empreendimentos
   Leitura REAL da lista SharePoint via Microsoft Graph, com login
   corporativo (MSAL). Nenhum segredo neste arquivo.
   Coleta ANÔNIMA: o painel não lê nem exibe nome, e-mail ou idEnvio.
   ========================================================= */
(function () {
  "use strict";

  /* ---------------------------------------------------------
     1. CONFIGURAÇÃO (campos vazios = CONFIGURAÇÃO PENDENTE)
     --------------------------------------------------------- */
  const CONFIG = {
    tenantId: "",   // CONFIGURAÇÃO PENDENTE — ID do diretório (tenant) no Entra ID
    clientId: "",   // CONFIGURAÇÃO PENDENTE — ID do aplicativo (client) registrado no Entra ID
    siteUrl: "",    // CONFIGURAÇÃO PENDENTE — URL do site "People Analytics - Pessoas e Cultura"
    listName: "Mudança de Sede - Dúvidas",
    listId: "",

    // Nomes INTERNOS das colunas usadas pelo painel
    campos: {
      idEnvio: "IdEnvio",       // usado só para descartar gravação duplicada; nunca exibido
      dataHora: "DataHora",
      numero: "NumeroDuvida",
      duvida: "Duvida",
      status: "Status",
      tema: "Tema",             // se a coluna não existir ou estiver vazia, o tema é identificado automaticamente
      modoTeste: "ModoTeste"
    },

    escopos: ["https://graph.microsoft.com/Sites.Read.All"],
    intervaloAtualizacaoMs: 60000,
    fuso: "America/Sao_Paulo",
    recentes: 5,
    perguntasPorPagina: 30,
    termosNuvem: 30
  };

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const semAcento = (s) => String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const chaveNome = (s) => semAcento(s).toLowerCase().replace(/\s+/g, " ").trim();

  /* ---------------------------------------------------------
     2. STATUS E TEMAS
     --------------------------------------------------------- */
  // Só existem dois status no painel. Vazio ou "Nova/Novo" = Nova; qualquer outro valor = Em tratamento.
  function normalizarStatus(v) {
    const s = chaveNome((v && v.Value) || v);
    return (!s || s === "nova" || s === "novo") ? "Nova" : "Em tratamento";
  }

  // Identificação automática de tema (usada quando a coluna Tema está vazia). Primeira regra que casar vence.
  const TEMAS = [
    ["Cronograma e prazos", /\b(prazo|prazos|data|datas|quando|cronograma|previs|inicio|início|calendario|calendário|etapa|fase)/],
    ["Moradia e custo de vida", /\b(morad|moraria|aluguel|casa|imovel|imóvel|apartamento|custo de vida|hospedagem|hotel)/],
    ["Transporte e deslocamento", /\b(transporte|desloca|onibus|ônibus|fretado|carro|viagem|viagens|estacionamento|distancia|distância|vale.?transporte)/],
    ["Família e dependentes", /\b(famil|filho|filha|esposa|marido|conjuge|cônjuge|dependente|escola|creche)/],
    ["Benefícios e remuneração", /\b(benefic|salari|salário|remunera|auxilio|auxílio|ajuda de custo|plano de saude|plano de saúde|bonus|bônus|reembolso|vale)/],
    ["Trabalho, jornada e flexibilidade", /\b(remoto|home office|hibrid|híbrid|jornada|horario|horário|escala|flexib|presencial|turno)/],
    ["Estrutura e local de trabalho", /\b(estrutura|escritorio|escritório|predio|prédio|sede|infraestrutura|espaço|espaco|sala|equipamento|local de trabalho)/],
    ["Carreira e pessoas", /\b(carreira|cargo|promo|desligamento|demiss|vaga|contrata|equipe|time|lider|líder|gestor)/]
  ];
  function classificarTema(texto) {
    const t = String(texto || "").toLowerCase();
    for (const [nome, re] of TEMAS) if (re.test(t) || re.test(semAcento(t))) return nome;
    return "Outros";
  }

  /* ---------------------------------------------------------
     3. NUVEM DE PALAVRAS (frequência real, sem palavras comuns)
     --------------------------------------------------------- */
  // Termos: frequência de palavras relevantes (sem stopwords), com expressões compostas e plural agrupado
  const STOPWORDS = new Set((
    "a à ao aos as às o os um uma uns umas de do da dos das d em no na nos nas num numa por pelo pela pelos pelas " +
    "para pra pro pras pros com sem sob sobre entre ate até apos após desde contra perante durante mediante " +
    "e ou mas nem que se porque pois porem porém quando onde como qual quais quanto quanta quantos quantas quem cujo " +
    "ja já nao não sim tambem também ainda mais menos muito muita muitos muitas pouco pouca bem mal so só apenas " +
    "eu tu ele ela nos nós vos vós eles elas voce você voces vocês me te lhe lhes se nosso nossa nossos nossas " +
    "meu minha meus minhas seu sua seus suas dele dela deles delas isso isto aquilo esse essa esses essas este esta " +
    "estes estas aquele aquela aqueles aquelas outro outra outros outras mesmo mesma mesmos mesmas cada todo toda todos todas " +
    "algum alguma alguns algumas nenhum nenhuma qualquer quaisquer tal tais " +
    "ser sera será serao serão seria seriam sao são era eram foi foram sendo sido sou somos esta está estao estão estava estavam " +
    "estar estara estará estarao estarão ter tem têm tera terá terao terão teria teriam tinha tinham tido tendo " +
    "haver ha há havera haverá haveria houve fazer faz fara fará farao farão feito ficar fica ficara ficará ficarao ficarão " +
    "vai vao vão ir iremos vamos pode podem podera poderá poderao poderão poderia poderiam deve devem devera deverá deveria " +
    "precisa precisam precisar saber sabe gostaria gostariamos gostaríamos queria queremos quer existe existem " +
    "aqui ali la lá entao então assim agora depois antes sempre nunca tambem caso sobre etc ok " +
    "dúvida duvida dúvidas duvidas pergunta perguntas questao questão gente pessoal time equipe acerca respeito relacao relação " +
    "forma sera algum algo coisa coisas vez vezes parte ficam ficaria ficariam teremos temos tenho tenha tenham " +
    "vamos vou irao irão seremos sejam seja possivel possível sobre atualmente hoje puder puderem " +
    "relacionado relacionada relacionados relacionadas possibilidade referente quanto " +
    "definido definida definidos definir continuara continuar continuarao disponibilizar disponibiliza empresa " +
    "novo novos nova novas mesma mesmo apos atual atuais sera gostaria pessoas"
  ).split(/\s+/).map(semAcento));

  const EXPRESSOES = [
    ["ribeirão preto", "Ribeirão Preto"], ["plano de saúde", "plano de saúde"], ["plano de saude", "plano de saúde"],
    ["trabalho remoto", "trabalho remoto"], ["home office", "home office"], ["ajuda de custo", "ajuda de custo"],
    ["auxílio mudança", "auxílio mudança"], ["auxilio mudança", "auxílio mudança"], ["auxílio moradia", "auxílio moradia"],
    ["vale transporte", "vale-transporte"], ["vale-transporte", "vale-transporte"], ["vale alimentação", "vale-alimentação"],
    ["vale refeição", "vale-refeição"], ["carga horária", "carga horária"], ["nova sede", "nova sede"]
  ];

  function extrairTermos(textos) {
    const contagem = new Map(); // chave → { n, formas: Map(forma → n) }
    const somar = (chave, forma) => {
      let t = contagem.get(chave);
      if (!t) { t = { n: 0, formas: new Map() }; contagem.set(chave, t); }
      t.n += 1;
      t.formas.set(forma, (t.formas.get(forma) || 0) + 1);
    };

    textos.forEach((texto) => {
      let t = " " + String(texto || "").toLowerCase() + " ";
      EXPRESSOES.forEach(([frase, rotulo]) => {
        const re = new RegExp("(^|[^\\p{L}])" + frase.replace(/[-]/g, "[- ]?") + "(?=[^\\p{L}]|$)", "giu");
        t = t.replace(re, (m, pre) => { somar("expr:" + semAcento(rotulo).toLowerCase(), rotulo); return pre + " "; });
      });
      (t.match(/[\p{L}\p{N}]+/gu) || []).forEach((palavra) => {
        const chave = semAcento(palavra);
        if (chave.length < 3 || /^\d+$/.test(chave) || STOPWORDS.has(chave)) return;
        somar(chave, palavra);
      });
    });

    // Agrupa plural simples ("filhos" → "filho") quando a forma singular também aparece
    Array.from(contagem.keys()).forEach((k) => {
      if (k.startsWith("expr:") || !k.endsWith("s")) return;
      const singular = k.slice(0, -1);
      const alvo = contagem.get(singular);
      if (alvo && contagem.has(k)) {
        const origem = contagem.get(k);
        alvo.n += origem.n;
        origem.formas.forEach((n, f) => alvo.formas.set(f, (alvo.formas.get(f) || 0) + n));
        contagem.delete(k);
      }
    });

    return Array.from(contagem.entries())
      .map(([chave, t]) => {
        const forma = Array.from(t.formas.entries()).sort((a, b) => b[1] - a[1])[0][0];
        return { chave, termo: forma, n: t.n };
      })
      .sort((a, b) => b.n - a.n || a.termo.localeCompare(b.termo, "pt-BR"));
  }


  window.PortalCalc = Object.freeze({ extrairTermos, classificarTema, normalizarStatus });

  /* ---------------------------------------------------------
     4. DATAS
     --------------------------------------------------------- */
  const fmt = new Intl.DateTimeFormat("pt-BR", {
    timeZone: CONFIG.fuso, day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit", hour12: false
  });
  function dataHora(d) {
    const p = {};
    fmt.formatToParts(d).forEach((x) => { p[x.type] = x.value; });
    return p.day + "/" + p.month + "/" + p.year + " - " + p.hour + ":" + p.minute;
  }

  /* ---------------------------------------------------------
     5. AUTENTICAÇÃO (MSAL) E MICROSOFT GRAPH
     --------------------------------------------------------- */
  class ErroPortal extends Error {
    constructor(tipo, mensagem) { super(mensagem); this.tipo = tipo; }
  }

  let pca = null;
  let conta = null;

  function configuracaoPendente() {
    const faltando = [];
    if (!CONFIG.tenantId) faltando.push(["tenantId", "ID do diretório (tenant) do Entra ID"]);
    if (!CONFIG.clientId) faltando.push(["clientId", "ID do aplicativo (client) registrado no Entra ID"]);
    if (!CONFIG.siteUrl) faltando.push(["siteUrl", "URL do site SharePoint \"People Analytics - Pessoas e Cultura\""]);
    return faltando;
  }

  const redirectUri = () => location.origin + location.pathname.replace(/index\.html$/, "");

  async function iniciarAuth() {
    if (!window.msal) throw new ErroPortal("config", "Biblioteca de autenticação não carregada (vendor/msal-browser.min.js).");
    pca = new window.msal.PublicClientApplication({
      auth: {
        clientId: CONFIG.clientId,
        authority: "https://login.microsoftonline.com/" + CONFIG.tenantId,
        redirectUri: redirectUri(),
        navigateToLoginRequestUrl: true
      },
      cache: { cacheLocation: "localStorage" }
    });
    await pca.initialize();
    const retorno = await pca.handleRedirectPromise();
    if (retorno && retorno.account) pca.setActiveAccount(retorno.account);
    conta = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
    if (conta) pca.setActiveAccount(conta);
  }

  function entrar() {
    pca.loginRedirect({ scopes: CONFIG.escopos, prompt: "select_account" });
  }

  async function obterToken() {
    if (!conta) throw new ErroPortal("login", "Login necessário.");
    try {
      const r = await pca.acquireTokenSilent({ scopes: CONFIG.escopos, account: conta });
      return r.accessToken;
    } catch (e) {
      if (e instanceof window.msal.InteractionRequiredAuthError) throw new ErroPortal("login", "Sessão expirada.");
      throw new ErroPortal("auth", "Não foi possível obter o acesso (" + (e.errorCode || e.message) + ").");
    }
  }

  async function graphGet(url, tentativa) {
    const token = await obterToken();
    let resp;
    try {
      resp = await fetch(url.startsWith("http") ? url : GRAPH + url, {
        headers: { Authorization: "Bearer " + token, Accept: "application/json" },
        cache: "no-store"
      });
    } catch (e) {
      throw new ErroPortal("rede", "Sem conexão com o Microsoft Graph.");
    }
    if ((resp.status === 429 || resp.status === 503) && (tentativa || 0) < 2) {
      const espera = Math.min(parseInt(resp.headers.get("Retry-After") || "5", 10), 30) * 1000;
      await new Promise((r) => setTimeout(r, espera));
      return graphGet(url, (tentativa || 0) + 1);
    }
    if (resp.status === 401) throw new ErroPortal("login", "Sessão expirada.");
    if (resp.status === 403) throw new ErroPortal("permissao", "Sua conta não tem permissão para ler a lista \"" + CONFIG.listName + "\" no SharePoint.");
    if (resp.status === 404) throw new ErroPortal("naoEncontrado", "Site ou lista não encontrados. Confira siteUrl e listName em portal.js.");
    if (!resp.ok) throw new ErroPortal("http", "O Microsoft Graph respondeu HTTP " + resp.status + ".");
    return resp.json();
  }

  let siteId = null;
  let listId = CONFIG.listId || null;

  async function resolverLista() {
    if (!siteId) {
      const u = new URL(CONFIG.siteUrl);
      const m = u.pathname.match(/^\/(sites|teams)\/[^/]+/i);
      const caminho = m ? m[0] : "";
      const site = await graphGet("/sites/" + u.hostname + ":" + (caminho || "/") + "?$select=id");
      siteId = site.id;
    }
    if (!listId) {
      const alvo = chaveNome(CONFIG.listName);
      let url = "/sites/" + siteId + "/lists?$select=id,displayName&$top=200";
      while (url && !listId) {
        const r = await graphGet(url);
        const achou = (r.value || []).find((l) => chaveNome(l.displayName) === alvo);
        if (achou) listId = achou.id;
        url = r["@odata.nextLink"];
      }
      if (!listId) throw new ErroPortal("naoEncontrado", "Lista \"" + CONFIG.listName + "\" não encontrada no site configurado.");
    }
  }

  const ehTeste = (v) => v === true || /^(yes|sim|true|1)$/i.test(String(v == null ? "" : v).trim());

  async function carregarRegistros() {
    await resolverLista();
    const c = CONFIG.campos;
    // Campos lidos: somente os necessários para a triagem. Nenhum dado pessoal é lido ou guardado.
    const sel = [c.idEnvio, c.dataHora, c.numero, c.duvida, c.status, c.tema, c.modoTeste].filter(Boolean).join(",");
    const base = "/sites/" + siteId + "/lists/" + listId + "/items?$select=id,createdDateTime&$top=500&$expand=fields";
    async function buscar(url) {
      const itens = [];
      while (url) {
        const r = await graphGet(url);
        itens.push.apply(itens, r.value || []);
        url = r["@odata.nextLink"];
      }
      return itens;
    }
    let brutos;
    try {
      brutos = await buscar(base + "($select=" + sel + ")");
    } catch (e) {
      // Alguma coluna configurada não existe na lista (ex.: Tema): lê os campos disponíveis e usa só os necessários
      if (e.tipo !== "http") throw e;
      brutos = await buscar(base);
    }

    const vistos = new Set();
    const registros = [];
    brutos.forEach((item) => {
      const f = item.fields || {};
      const duvida = String(f[c.duvida] || "").trim();
      if (!duvida || ehTeste(f[c.modoTeste])) return;               // válido = dúvida preenchida e não-teste
      const chave = [f[c.idEnvio] || item.id, f[c.numero] || "", duvida].join("|");
      if (vistos.has(chave)) return;                                  // ignora gravação duplicada do mesmo envio
      vistos.add(chave);
      const d = new Date(f[c.dataHora] || item.createdDateTime);
      const temaLista = String((f[c.tema] && f[c.tema].Value) || f[c.tema] || "").trim();
      registros.push({
        id: item.id,
        data: isNaN(d) ? null : d,
        numero: Number(f[c.numero]) || 0,
        duvida,
        status: normalizarStatus(f[c.status]),
        tema: temaLista || classificarTema(duvida)
      });
    });
    // Nº = ordem de chegada (1 = primeira dúvida recebida)
    registros.sort((a, b) => ((a.data ? a.data.getTime() : 0) - (b.data ? b.data.getTime() : 0)) || a.numero - b.numero);
    registros.forEach((r, i) => { r.n = i + 1; });
    return registros;
  }

  /* ---------------------------------------------------------
     6. ELEMENTOS E ESTADO
     --------------------------------------------------------- */
  const $ = (id) => document.getElementById(id);
  const el = {
    lastUpdate: $("last-update"), updateAuto: $("update-auto"),
    stateConfig: $("state-config"), configMissing: $("config-missing"),
    stateLogin: $("state-login"), btnLogin: $("btn-login"),
    stateError: $("state-error"), errorText: $("error-text"), btnRetry: $("btn-retry"),
    app: $("app"),
    tabGeral: $("tab-geral"), tabPerguntas: $("tab-perguntas"),
    panelGeral: $("panel-geral"), panelPerguntas: $("panel-perguntas"),
    kpiTotal: $("kpi-total"), kpiNovas: $("kpi-novas"), kpiTratamento: $("kpi-tratamento"),
    bars: $("bars"), barsEmpty: $("bars-empty"),
    cloud: $("cloud"), cloudEmpty: $("cloud-empty"),
    recentesRows: $("recentes-rows"), recentesEmpty: $("recentes-empty"), btnVerTodas: $("btn-ver-todas"),
    fBusca: $("f-busca"), fTema: $("f-tema"), fStatus: $("f-status"),
    qCount: $("q-count"), qRows: $("q-rows"), qEmpty: $("q-empty"), qMore: $("q-more")
  };
  const estado = { registros: [], carregado: false, limite: CONFIG.perguntasPorPagina };

  function no(tag, cls, texto) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    if (texto != null) n.textContent = texto;
    return n;
  }

  /* ---------------------------------------------------------
     7. VISÃO GERAL
     --------------------------------------------------------- */
  function renderGeral() {
    const regs = estado.registros;
    el.kpiTotal.textContent = regs.length;
    el.kpiNovas.textContent = regs.filter((r) => r.status === "Nova").length;
    el.kpiTratamento.textContent = regs.filter((r) => r.status === "Em tratamento").length;

    // Dúvidas por tema (barras horizontais; "Outros" sempre por último)
    const cont = new Map();
    regs.forEach((r) => cont.set(r.tema, (cont.get(r.tema) || 0) + 1));
    const temas = Array.from(cont.entries()).sort((a, b) =>
      (a[0] === "Outros") - (b[0] === "Outros") || b[1] - a[1] || a[0].localeCompare(b[0], "pt-BR"));
    const max = temas.reduce((m, t) => Math.max(m, t[1]), 0);
    el.bars.textContent = "";
    temas.forEach(([nome, n]) => {
      const li = no("li", "bar");
      li.title = nome + ": " + n + (n === 1 ? " dúvida" : " dúvidas");
      li.appendChild(no("span", "bar-label", nome));
      li.appendChild(no("span", "bar-value", String(n)));
      const track = no("span", "bar-track");
      const fill = no("span", "bar-fill");
      fill.style.width = (max ? (n / max) * 100 : 0).toFixed(1) + "%";
      track.appendChild(fill);
      li.appendChild(track);
      el.bars.appendChild(li);
    });
    el.barsEmpty.hidden = temas.length > 0;

    renderNuvem();

    // Perguntas recentes
    const recentes = regs.slice().sort((a, b) => b.n - a.n).slice(0, CONFIG.recentes);
    preencherTabela(el.recentesRows, recentes, "");
    el.recentesEmpty.hidden = recentes.length > 0;
    el.recentesRows.closest("table").hidden = recentes.length === 0;
  }

  const CORES_NUVEM = ["#D6245A", "#1D2A5B", "#C2185B", "#574797", "#E0487A", "#2E3F86", "#B0447E", "#0A7C64"];

  function renderNuvem() {
    const termos = extrairTermos(estado.registros.map((r) => r.duvida)).slice(0, CONFIG.termosNuvem);
    el.cloud.textContent = "";
    el.cloudEmpty.hidden = termos.length > 0;
    el.cloud.hidden = termos.length === 0;
    if (!termos.length) return;
    const max = termos[0].n;
    const min = termos[termos.length - 1].n;
    const ordem = [];
    termos.forEach((t, i) => { if (i % 2 === 0) ordem.push(t); else ordem.unshift(t); }); // maiores no centro
    ordem.forEach((t) => {
      const rank = termos.indexOf(t);
      const peso = max === min ? 0.35 : (t.n - min) / (max - min);
      const b = no("button", rank < 3 ? "cloud-word w-top" : "cloud-word", t.termo);
      b.type = "button";
      b.style.setProperty("--t", peso.toFixed(3));
      b.style.color = CORES_NUVEM[rank % CORES_NUVEM.length];
      b.style.opacity = String(0.65 + 0.35 * peso);
      b.title = t.termo + ": " + t.n + (t.n === 1 ? " ocorrência" : " ocorrências");
      b.addEventListener("click", () => { el.fBusca.value = t.termo; abrirPerguntas(); });
      el.cloud.appendChild(b);
    });
  }

  /* ---------------------------------------------------------
     8. TABELA DE PERGUNTAS
     --------------------------------------------------------- */
  function textoComDestaque(texto, termo) {
    const frag = document.createDocumentFragment();
    const alvo = semAcento(termo).toLowerCase().trim();
    if (!alvo) { frag.appendChild(document.createTextNode(texto)); return frag; }
    const chars = Array.from(texto);
    const norm = chars.map((c) => (semAcento(c).toLowerCase() || c).charAt(0)).join("");
    let i = 0;
    let pos = norm.indexOf(alvo);
    while (pos !== -1) {
      if (pos > i) frag.appendChild(document.createTextNode(chars.slice(i, pos).join("")));
      frag.appendChild(no("mark", null, chars.slice(pos, pos + alvo.length).join("")));
      i = pos + alvo.length;
      pos = norm.indexOf(alvo, i);
    }
    if (i < chars.length) frag.appendChild(document.createTextNode(chars.slice(i).join("")));
    return frag;
  }

  function preencherTabela(tbody, lista, busca) {
    tbody.textContent = "";
    lista.forEach((r) => {
      const tr = document.createElement("tr");
      tr.appendChild(no("td", "c-num", String(r.n)));
      const tdP = no("td", "c-perg");
      tdP.appendChild(textoComDestaque(r.duvida, busca));
      tr.appendChild(tdP);
      tr.appendChild(no("td", "c-tema", r.tema));
      const tdS = no("td", "c-status");
      tdS.appendChild(no("span", "status " + (r.status === "Nova" ? "status--nova" : "status--tratamento"), r.status));
      tr.appendChild(tdS);
      tbody.appendChild(tr);
    });
  }

  function atualizarOpcoesTema() {
    const atual = el.fTema.value;
    const temas = Array.from(new Set(estado.registros.map((r) => r.tema)))
      .sort((a, b) => (a === "Outros") - (b === "Outros") || a.localeCompare(b, "pt-BR"));
    el.fTema.textContent = "";
    el.fTema.appendChild(new Option("Todos", ""));
    temas.forEach((t) => el.fTema.appendChild(new Option(t, t)));
    el.fTema.value = temas.includes(atual) ? atual : "";
  }

  function renderPerguntas() {
    const buscaTxt = el.fBusca.value;
    const busca = semAcento(buscaTxt).toLowerCase().trim();
    const tema = el.fTema.value;
    const status = el.fStatus.value;
    const total = estado.registros.length;
    const lista = estado.registros
      .filter((r) => (!tema || r.tema === tema) && (!status || r.status === status) &&
        (!busca || semAcento(r.duvida).toLowerCase().includes(busca)))
      .sort((a, b) => b.n - a.n);

    el.qCount.textContent = (busca || tema || status)
      ? lista.length + " de " + total + (total === 1 ? " pergunta" : " perguntas")
      : total + (total === 1 ? " pergunta" : " perguntas");
    preencherTabela(el.qRows, lista.slice(0, estado.limite), buscaTxt);
    el.qRows.closest("table").hidden = lista.length === 0;
    el.qEmpty.hidden = lista.length > 0;
    el.qEmpty.textContent = total === 0 ? "Sem dúvidas registradas." : "Nenhuma pergunta encontrada com os filtros selecionados.";
    const resto = lista.length - estado.limite;
    el.qMore.hidden = resto <= 0;
    el.qMore.textContent = "Mostrar mais (" + Math.max(resto, 0) + ")";
  }

  function renderTudo() {
    renderGeral();
    atualizarOpcoesTema();
    renderPerguntas();
  }

  /* ---------------------------------------------------------
     9. NAVEGAÇÃO E FILTROS
     --------------------------------------------------------- */
  function selecionarAba(qual, foco) {
    const perg = qual === "perguntas";
    el.tabGeral.setAttribute("aria-selected", String(!perg));
    el.tabPerguntas.setAttribute("aria-selected", String(perg));
    el.tabGeral.tabIndex = perg ? -1 : 0;
    el.tabPerguntas.tabIndex = perg ? 0 : -1;
    el.panelGeral.hidden = perg;
    el.panelPerguntas.hidden = !perg;
    const hash = perg ? "#perguntas" : "#visao-geral";
    if (location.hash !== hash) history.replaceState(null, "", hash);
    if (foco) (perg ? el.tabPerguntas : el.tabGeral).focus();
  }

  function abrirPerguntas() {
    estado.limite = CONFIG.perguntasPorPagina;
    if (estado.carregado) renderPerguntas();
    selecionarAba("perguntas");
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  function ligarEventos() {
    el.btnLogin.addEventListener("click", entrar);
    el.btnRetry.addEventListener("click", () => atualizar());
    el.tabGeral.addEventListener("click", () => selecionarAba("geral"));
    el.tabPerguntas.addEventListener("click", () => selecionarAba("perguntas"));
    [el.tabGeral, el.tabPerguntas].forEach((t) => t.addEventListener("keydown", (e) => {
      if (["ArrowRight", "ArrowLeft", "ArrowUp", "ArrowDown"].includes(e.key)) {
        e.preventDefault();
        selecionarAba(el.tabGeral.getAttribute("aria-selected") === "true" ? "perguntas" : "geral", true);
      }
    }));
    window.addEventListener("hashchange", () => selecionarAba(location.hash === "#perguntas" ? "perguntas" : "geral"));
    el.btnVerTodas.addEventListener("click", () => {
      el.fBusca.value = ""; el.fTema.value = ""; el.fStatus.value = "";
      abrirPerguntas();
    });
    let t = null;
    const filtrar = () => { estado.limite = CONFIG.perguntasPorPagina; if (estado.carregado) renderPerguntas(); };
    el.fBusca.addEventListener("input", () => { clearTimeout(t); t = setTimeout(filtrar, 180); });
    el.fTema.addEventListener("change", filtrar);
    el.fStatus.addEventListener("change", filtrar);
    el.qMore.addEventListener("click", () => { estado.limite += CONFIG.perguntasPorPagina; renderPerguntas(); });
  }

  /* ---------------------------------------------------------
     10. ATUALIZAÇÃO AUTOMÁTICA (60 s)
     --------------------------------------------------------- */
  let timer = null;
  let carregando = false;

  function mostrarErro(msg) {
    el.errorText.textContent = msg;
    el.stateError.hidden = false;
    el.updateAuto.classList.add("is-error");
  }

  async function atualizar() {
    if (carregando) return;
    carregando = true;
    clearTimeout(timer);
    el.updateAuto.classList.add("is-loading");
    try {
      estado.registros = await carregarRegistros();
      estado.carregado = true;
      el.stateError.hidden = true;
      el.stateLogin.hidden = true;
      el.updateAuto.classList.remove("is-error");
      el.app.hidden = false;
      el.lastUpdate.textContent = dataHora(new Date());
      renderTudo();
    } catch (e) {
      if (e.tipo === "login") {
        el.app.hidden = !estado.carregado;
        el.stateLogin.hidden = false;
      } else {
        const base = e.message || "Falha ao carregar os dados.";
        mostrarErro(estado.carregado ? base + " Exibindo os últimos dados carregados." : base);
      }
    } finally {
      carregando = false;
      el.updateAuto.classList.remove("is-loading");
      if (!document.hidden) timer = setTimeout(atualizar, CONFIG.intervaloAtualizacaoMs);
    }
  }

  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { clearTimeout(timer); return; }
    if (conta && !carregando) atualizar();
  });

  /* ---------------------------------------------------------
     11. INICIALIZAÇÃO
     --------------------------------------------------------- */
  async function init() {
    ligarEventos();
    selecionarAba(location.hash === "#perguntas" ? "perguntas" : "geral");

    const faltando = configuracaoPendente();
    if (faltando.length) {
      faltando.forEach(([chave, desc]) => {
        const li = document.createElement("li");
        li.appendChild(no("code", null, chave));
        li.appendChild(document.createTextNode(" — " + desc));
        el.configMissing.appendChild(li);
      });
      el.stateConfig.hidden = false;
      el.updateAuto.classList.add("is-error");
      return;
    }
    try {
      await iniciarAuth();
    } catch (e) {
      mostrarErro("Falha na autenticação: " + (e.errorCode || e.message));
      return;
    }
    if (!conta) { el.stateLogin.hidden = false; return; }
    atualizar();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
