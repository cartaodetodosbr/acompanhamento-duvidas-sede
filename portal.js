/* =========================================================
   Painel de Triagem | Mudança de Sede — TODOS Empreendimentos
   Leitura REAL via fluxo Power Automate "Portal Mudança de Sede - Consultar
   Dúvidas" (GET), que lê o SharePoint e devolve só campos anônimos.
   Coleta ANÔNIMA: o painel não recebe nem exibe nome, e-mail ou idEnvio.
   ========================================================= */
(function () {
  "use strict";

  /* ---------------------------------------------------------
     1. CONFIGURAÇÃO (campos vazios = CONFIGURAÇÃO PENDENTE)
     --------------------------------------------------------- */
  const CONFIG = {
    // URL do gatilho "When an HTTP request is received" do fluxo de CONSULTA (método GET).
    // Não é o endpoint do formulário. Nenhuma outra credencial fica neste arquivo.
    endpointConsulta: "https://defaulte93279240f9745ba871f4a124f3343.19.environment.api.powerplatform.com:443/powerautomate/automations/direct/cu/24/workflows/5387eff6ce27471a8ae77937130d510e/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=PQQsq6cfu1OjiZu1cf3cPRk1AaWouLrRQ7grrYD2MGY",
    timeoutMs: 30000,
    intervaloAtualizacaoMs: 60000,
    fuso: "America/Sao_Paulo",
    recentes: 5,
    perguntasPorPagina: 30,
    termosNuvem: 30
  };


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
     5. LEITURA DOS DADOS (fluxo de consulta, GET simples — sem cabeçalhos extras)
     Retorno esperado: { ok: true, atualizadoEm, total, perguntas: [
       { numero, dataHora, pergunta, status, tema } ] }
     --------------------------------------------------------- */
  class ErroPortal extends Error {
    constructor(tipo, mensagem) { super(mensagem); this.tipo = tipo; }
  }

  async function carregarRegistros() {
    const controller = new AbortController();
    const t = setTimeout(() => controller.abort(), CONFIG.timeoutMs);
    let resp;
    try {
      resp = await fetch(CONFIG.endpointConsulta, { method: "GET", cache: "no-store", signal: controller.signal });
    } catch (e) {
      // Bloqueio de CORS também chega aqui; o motivo exato aparece no console (F12)
      throw new ErroPortal("rede", e && e.name === "AbortError"
        ? "O serviço de dados demorou para responder."
        : "Não foi possível conectar ao serviço de dados.");
    } finally {
      clearTimeout(t);
    }
    if (!resp.ok) throw new ErroPortal("http", "O serviço de dados respondeu HTTP " + resp.status + ".");
    let dados;
    try { dados = await resp.json(); } catch (_) { throw new ErroPortal("formato", "Resposta do serviço de dados em formato inválido."); }
    if (!dados || dados.ok !== true || !Array.isArray(dados.perguntas)) {
      throw new ErroPortal("formato", "Resposta do serviço de dados sem a lista de perguntas.");
    }

    const registros = [];
    dados.perguntas.forEach((p, i) => {
      const duvida = String((p && p.pergunta) || "").trim();
      if (!duvida) return;
      const d = new Date(p.dataHora);
      const temaLista = String((p.tema && p.tema.Value) || p.tema || "").trim();
      registros.push({
        ordem: i,
        data: isNaN(d) ? null : d,
        numero: Number(p.numero) || 0,
        duvida,
        status: normalizarStatus(p.status),
        tema: temaLista || classificarTema(duvida)
      });
    });
    // Nº = ordem de chegada (1 = primeira dúvida recebida)
    registros.sort((a, b) => ((a.data ? a.data.getTime() : 0) - (b.data ? b.data.getTime() : 0)) || a.numero - b.numero || a.ordem - b.ordem);
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
      el.updateAuto.classList.remove("is-error");
      el.app.hidden = false;
      el.lastUpdate.textContent = dataHora(new Date());
      renderTudo();
    } catch (e) {
      const base = e.message || "Falha ao carregar os dados.";
      mostrarErro(estado.carregado ? base + " Exibindo os últimos dados carregados." : base);
    } finally {
      carregando = false;
      el.updateAuto.classList.remove("is-loading");
      if (!document.hidden) timer = setTimeout(atualizar, CONFIG.intervaloAtualizacaoMs);
    }
  }

  document.addEventListener("visibilitychange", () => {
    if (document.hidden) { clearTimeout(timer); return; }
    if (CONFIG.endpointConsulta && !carregando) atualizar();
  });

  /* ---------------------------------------------------------
     11. INICIALIZAÇÃO
     --------------------------------------------------------- */
  async function init() {
    ligarEventos();
    selecionarAba(location.hash === "#perguntas" ? "perguntas" : "geral");

    if (!CONFIG.endpointConsulta) {
      const li = document.createElement("li");
      li.appendChild(no("code", null, "endpointConsulta"));
      li.appendChild(document.createTextNode(" — URL do fluxo de consulta do Power Automate"));
      el.configMissing.appendChild(li);
      el.stateConfig.hidden = false;
      el.updateAuto.classList.add("is-error");
      return;
    }
    atualizar();
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", init);
  else init();
})();
