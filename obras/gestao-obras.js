/* =====================================================================
   Gestão e Acompanhamento de Obras — Construtora Vieira & Silva
   ---------------------------------------------------------------------
   Carregado pelo index.html (import dinâmico). Usa o mesmo app Firebase,
   o mesmo login e as mesmas telas/estilos do CRM.

   Firestore (tudo debaixo de obras/<obra>):
     obras/<obra>                cadastro, situação, membros {uid: papel}
     obras/<obra>/itens/<id>     EAP: etapa | subetapa | tarefa
     obras/<obra>/codigos/<cód>  trava de código único por obra
     obras/<obra>/lancamentos    produção lançada (somente inclusão, id fixo)
     obras/<obra>/rdos/<data>    Relatório Diário de Obra (um por dia)
     obras/<obra>/arquivos/<id>  fotos comprimidas e anexos pequenos
     obras/<obra>/historico      auditoria (somente inclusão)
     obras/<obra>/serie          eventos de avanço para a curva de evolução
     obras/<obra>/el, cfg        mapa de andamento (páginas antigas, inalteradas)
   As garantias de servidor estão em firestore.rules (raiz do repositório).
   ===================================================================== */
import {
  collection, doc, getDoc, getDocs, setDoc, updateDoc, onSnapshot, query, where,
  serverTimestamp, writeBatch, runTransaction
} from "https://www.gstatic.com/firebasejs/10.12.0/firebase-firestore.js";

const LIB = {
  xlsx: "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
  jspdf: "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
  autotable: "https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"
};

// ---------------------------------------------------------------------
// Constantes de domínio
// ---------------------------------------------------------------------
export const SITUACOES_OBRA = {
  nao_iniciada: "Não iniciada", em_andamento: "Em andamento", paralisada: "Paralisada", concluida: "Concluída"
};
export const SITUACOES_TAREFA = { nao_iniciada: "Não iniciada", em_andamento: "Em andamento", concluida: "Concluída" };
const PAPEIS = { admin: "Administrador", gestor: "Gestor", colaborador: "Colaborador", consulta: "Consulta" };
const TIPOS = { etapa: "Etapa", subetapa: "Subetapa", tarefa: "Tarefa" };
const UNIDADES = ["m²", "m³", "m", "kg", "t", "un", "vb", "h", "L", "pç", "cj", "pt"];
const CLIMA = ["Bom", "Nublado", "Chuva fraca", "Chuva forte", "Impraticável"];
const CONDICAO = { praticavel: "Praticável", parcial: "Parcialmente praticável", impraticavel: "Impraticável" };
const FUNCOES = ["Mestre de obras", "Encarregado", "Pedreiro", "Servente", "Carpinteiro", "Armador", "Eletricista", "Encanador", "Pintor", "Gesseiro", "Azulejista"];
const STATUS_RDO = { rascunho: "Rascunho", enviado: "Enviado", aprovado: "Aprovado", cancelado: "Cancelado" };
const DIAS_VENCIMENTO = 7;
const LOTE_IMPORT = 10;     // itens por gravação (limite de consultas das regras do Firestore)
const MAX_ATIV_RDO = 100;   // atividades vinculadas por RDO (limite de 500 gravações por transação)

// ---------------------------------------------------------------------
// Utilitários
// ---------------------------------------------------------------------
export const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const pad = n => String(n).padStart(2, "0");
export const hojeISO = (d = new Date()) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
export const somaDias = (iso, n) => { const [y, m, d] = iso.split("-").map(Number); return hojeISO(new Date(y, m - 1, d + n)); };
export const fmtData = iso => { if (!iso) return "—"; const [y, m, d] = String(iso).slice(0, 10).split("-"); return d ? `${d}/${m}/${y}` : "—"; };
const tsData = ts => { const d = ts && (ts.toDate ? ts.toDate() : (ts instanceof Date ? ts : null)); return d ? d : null; };
const fmtDataHora = ts => { const d = tsData(ts) || (typeof ts === "string" ? new Date(ts) : null); return d && !isNaN(d) ? `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}` : "—"; };
const NF = new Intl.NumberFormat("pt-BR", { maximumFractionDigits: 3 });
const NF1 = new Intl.NumberFormat("pt-BR", { minimumFractionDigits: 0, maximumFractionDigits: 1 });
export const fmtNum = n => (n == null || n === "" || isNaN(n)) ? "—" : NF.format(n);
export const fmtPct = n => (n == null || isNaN(n)) ? "—" : NF1.format(n) + "%";
export const arred = x => Math.round(Number(x) * 1e6) / 1e6;
/** "1.234,5" | "1234.5" | "12,5" → número; vazio → null; inválido → NaN */
export function parseNum(v) {
  if (v == null) return null;
  if (typeof v === "number") return v;
  let s = String(v).trim().replace(/\s/g, "");
  if (!s) return null;
  if (s.includes(",")) s = s.replace(/\./g, "").replace(",", ".");
  return /^-?\d*\.?\d+$/.test(s) ? Number(s) : NaN;
}
/** Código EAP normalizado e a chave do documento de trava. */
export const normCodigo = c => String(c ?? "").trim().replace(/\s+/g, " ").toUpperCase();
export function validaCodigo(c) {
  const n = normCodigo(c);
  if (!n) return "Informe o código.";
  if (n.length > 40) return "Código muito longo (máx. 40 caracteres).";
  if (n.includes("/")) return "O código não pode conter “/”.";
  if (/^\.+$/.test(n) || /^__.*__$/.test(n)) return "Código inválido.";
  return "";
}
export const codigoKey = c => normCodigo(c);
/** Compara códigos tipo 1.2.10 de forma natural. */
export const cmpCodigo = (a, b) => String(a).localeCompare(String(b), "pt-BR", { numeric: true, sensitivity: "base" });
const slug = s => String(s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, "-").replace(/^-|-$/g, "").slice(0, 40) || "obra";

// ---------------------------------------------------------------------
// Cálculos (puros — testados separadamente)
// ---------------------------------------------------------------------
/**
 * Avanço de uma tarefa.
 *  - manual: percentual informado (0–100).
 *  - produção: executado ÷ previsto × 100. Previsto 0 → avanço indefinido (sinalizado).
 *    Acima de 100% é sinalizado; o avanço "considerado" para situação e consolidação fica em 100,
 *    mas a quantidade e o percentual reais continuam visíveis.
 */
export function calcAvanco(it) {
  if (!it || it.tipo !== "tarefa") return null;
  if ((it.modo || "manual") === "producao") {
    const prev = Number(it.qtdPrevista) || 0, exe = Number(it.qtdExecutada) || 0;
    if (prev <= 0) return { avanco: exe > 0 ? null : 0, real: null, acima: false, semPrevisto: true, exe, prev };
    const real = exe / prev * 100;
    return { avanco: Math.max(0, Math.min(100, real)), real, acima: exe > prev + 1e-9, semPrevisto: false, exe, prev };
  }
  const p = Math.max(0, Math.min(100, Number(it.percentual) || 0));
  return { avanco: p, real: p, acima: false, semPrevisto: false };
}
export function situacaoPorAvanco(av, exeSemPrevisto) {
  if (av == null) return exeSemPrevisto ? "em_andamento" : "nao_iniciada";
  if (av >= 100 - 1e-9) return "concluida";
  if (av > 0) return "em_andamento";
  return "nao_iniciada";
}
export const atrasada = (fimPrev, av, hoje = hojeISO()) => !!fimPrev && fimPrev < hoje && !(av >= 100 - 1e-9);
/** Consolida um conjunto de tarefas já derivadas. */
export function consolidar(tarefas) {
  const r = { total: tarefas.length, nao_iniciada: 0, em_andamento: 0, concluida: 0, atrasadas: 0, semPeso: 0, semPrevisto: 0,
    pctConcluidas: null, mediaSimples: null, ponderado: null, somaPesos: 0 };
  if (!tarefas.length) return r;
  let soma = 0, sp = 0, spa = 0;
  for (const t of tarefas) {
    r[t.situacao]++;
    if (t.atrasada) r.atrasadas++;
    if (t.calc.semPrevisto) r.semPrevisto++;
    const av = t.calc.avanco ?? 0;
    soma += av;
    const p = Number(t.peso);
    if (t.peso == null || t.peso === "" || !(p > 0)) r.semPeso++; else { sp += p; spa += p * av; }
  }
  r.pctConcluidas = r.concluida / r.total * 100;
  r.mediaSimples = soma / r.total;
  r.somaPesos = sp;
  r.ponderado = r.semPeso === 0 && sp > 0 ? spa / sp : null;
  return r;
}
/** Monta a árvore e os valores derivados. Rápido para milhares de itens (O(n log n)). */
export function derivar(lista, hoje = hojeISO()) {
  const porId = new Map(), filhos = new Map();
  for (const it of lista) { porId.set(it.id, it); }
  for (const it of lista) {
    const k = it.paiId && porId.has(it.paiId) ? it.paiId : "";
    if (!filhos.has(k)) filhos.set(k, []);
    filhos.get(k).push(it);
  }
  for (const arr of filhos.values()) arr.sort((a, b) => (a.ordem ?? 0) - (b.ordem ?? 0) || cmpCodigo(a.codigo, b.codigo));
  const der = new Map();
  const visit = (it, nivel, etapaId) => {
    const ch = filhos.get(it.id) || [];
    const d = { ...it, nivel, etapaId: it.tipo === "etapa" ? it.id : etapaId, filhos: ch.map(c => c.id) };
    der.set(it.id, d);
    if (it.tipo === "tarefa") {
      d.calc = calcAvanco(it);
      d.situacao = situacaoPorAvanco(d.calc.avanco, d.calc.semPrevisto && d.calc.exe > 0);
      d.atrasada = atrasada(it.fimPrev, d.calc.avanco, hoje);
      d.tarefas = [d];
    } else {
      d.tarefas = [];
    }
    for (const c of ch) { const dc = visit(c, nivel + 1, d.etapaId); d.tarefas.push(...dc.tarefas); }
    if (it.tipo !== "tarefa") {
      d.cons = consolidar(d.tarefas);
      const av = d.cons.ponderado ?? d.cons.mediaSimples;
      d.situacao = d.cons.total ? situacaoPorAvanco(av) : "nao_iniciada";
      d.atrasada = d.cons.atrasadas > 0;
      d.inicioCalc = it.inicioPrev || minData(d.tarefas.map(t => t.inicioPrev));
      d.fimCalc = it.fimPrev || maxData(d.tarefas.map(t => t.fimPrev));
    }
    return d;
  };
  const raiz = (filhos.get("") || []).map(it => visit(it, 0, null).id);
  const tarefas = [...der.values()].filter(d => d.tipo === "tarefa");
  return { der, raiz, tarefas, cons: consolidar(tarefas) };
}
const minData = arr => arr.filter(Boolean).sort()[0] || null;
const maxData = arr => arr.filter(Boolean).sort().slice(-1)[0] || null;

/** Papel do usuário numa obra a partir do perfil global e dos membros da obra. */
export function papelNaObra(obra, uid, perfilGlobal) {
  if (perfilGlobal === "admin") return "admin";
  if (!obra) return "";
  if (obra.legado || !obra.membros) return perfilGlobal === "operador" ? "colaborador" : perfilGlobal === "visualizador" ? "consulta" : "";
  const m = obra.membros[uid] || "";
  if (perfilGlobal === "visualizador") return m ? "consulta" : "";
  if (perfilGlobal === "operador") return m;
  return "";
}
export const capacidades = papel => ({
  ver: !!papel, registrar: ["admin", "gestor", "colaborador"].includes(papel),
  gerir: ["admin", "gestor"].includes(papel), admin: papel === "admin"
});

// ---------------------------------------------------------------------
// Estado
// ---------------------------------------------------------------------
let C = null;                 // contexto vindo do index.html
const S = {
  obras: [], obrasJson: [], usuarios: [], unsubObras: null, unsubUsuarios: null,
  obraId: null, unsubItens: null, unsubRdos: null, itens: [], itensProntos: false, rdos: [], modelo: null,
  aba: null, rota: [], sub: [], modoHub: null,
  ui: { busca: "", situacao: "", resp: "", etapa: "", prazo: "", soEtapas: false, ocultarConcluidas: false, abertos: new Set(), limite: 600, inicializado: null },
  rdoEd: null
};
const $ = id => document.getElementById(id);
const toast = m => C.showToast(m);
const uid = () => C.auth.currentUser && C.auth.currentUser.uid;
const nomeUsuario = id => { const u = S.usuarios.find(x => x.id === id); return u ? (u.nome || u.email || "") : ""; };
const meuNome = () => nomeUsuario(uid()) || (C.auth.currentUser && C.auth.currentUser.email) || "";
const obraAtual = () => S.obras.find(o => o.id === S.obraId) || null;
const papelAtual = () => papelNaObra(obraAtual(), uid(), C.perfil());
const pode = () => capacidades(papelAtual());
const R = (...p) => doc(C.db, "obras", S.obraId, ...p);
const COL = nome => collection(C.db, "obras", S.obraId, nome);
const erroMsg = e => e && e.code === "permission-denied" ? "Sem permissão para esta ação." : (e && e.message) || String(e);

// ---------------------------------------------------------------------
// Inicialização e rotas
// ---------------------------------------------------------------------
export function iniciarGestaoObras(ctx) {
  C = ctx;
  injetarCss();
  garantirModal();
  return { rotear, parar };
}

function parar() {
  _assinado = null;
  // ao sair, descarrega o mapa (o iframe guarda a sessão de quem estava logado)
  const corpo = $("obra-corpo"); if (corpo) { corpo.innerHTML = ""; corpo.dataset.mapa = ""; }
  [S.unsubObras, S.unsubUsuarios, S.unsubItens, S.unsubRdos].forEach(u => u && u());
  Object.assign(S, { obras: [], usuarios: [], unsubObras: null, unsubUsuarios: null, obraId: null, unsubItens: null, unsubRdos: null, itens: [], rdos: [], modelo: null });
}

let _assinado = null;
async function assinarGeral() {
  if (_assinado === uid()) return;
  parar();
  _assinado = uid();
  try { const r = await fetch("obras/obras.json", { cache: "no-cache" }); S.obrasJson = r.ok ? await r.json() : []; } catch (e) { S.obrasJson = []; }
  // Obras do obras.json: só aparecem como "legadas" se ainda NÃO foram cadastradas no banco.
  // Se já foram e o usuário não está na equipe, a leitura é negada → a obra some da lista dele.
  if (C.perfil() !== "admin") {
    const chec = await Promise.all(S.obrasJson.map(async j => {
      try { const d = await getDoc(doc(C.db, "obras", j.id)); return (!d.exists() || !d.data().membros) ? j : null; } catch (e) { return null; }
    }));
    S.obrasJson = chec.filter(Boolean);
  }
  S.unsubUsuarios = onSnapshot(collection(C.db, "usuarios"), snap => {
    S.usuarios = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => String(a.nome || a.email).localeCompare(String(b.nome || b.email), "pt-BR"));
  }, e => console.warn("usuarios", e));
  await new Promise(res => {
    const q = C.perfil() === "admin" ? collection(C.db, "obras") : query(collection(C.db, "obras"), where("membrosUids", "array-contains", uid()));
    let primeira = true;
    S.unsubObras = onSnapshot(q, snap => {
      const fs = snap.docs.map(d => ({ id: d.id, ...d.data() })).filter(o => o.nome);
      // Obras do obras.json ainda não cadastradas no banco continuam aparecendo (só o mapa).
      const leg = S.obrasJson.filter(j => !fs.some(o => o.id === j.id)).map(j => ({ ...j, legado: true, situacao: mapSituacaoLegado(j.status) }));
      S.obras = [...fs.map(o => { const j = S.obrasJson.find(x => x.id === o.id); return { ...o, blocos: o.blocos || (j && j.blocos) || [] }; }), ...leg]
        .sort((a, b) => (a.demo ? 1 : 0) - (b.demo ? 1 : 0) || String(a.nome).localeCompare(String(b.nome), "pt-BR"));
      // obra aberta antes de aparecer na lista (recém-criada ou recém-cadastrada): assina agora
      const oa = obraAtual();
      if (S.obraId && oa && !oa.legado && !S.unsubItens) abrirObra(S.obraId);
      if (primeira) { primeira = false; res(); } else rerender();
    }, e => {
      console.error("obras", e);
      S.obras = S.obrasJson.map(j => ({ ...j, legado: true, situacao: mapSituacaoLegado(j.status) }));
      if (primeira) { primeira = false; res(); } else rerender();
    });
  });
}
const mapSituacaoLegado = s => ({ "em andamento": "em_andamento", "concluída": "concluida", "paralisada": "paralisada" }[String(s || "").toLowerCase()] || "em_andamento");

function rerender() { if (S.rota) desenhar(); }

/** partes = segmentos depois de #acompanhamento (ou #rdo). */
async function rotear(partes, modo) {
  S.modoHub = modo || null;
  await assinarGeral();
  const [obraId, aba, ...resto] = partes;
  S.rota = partes;
  if (obraId && obraId !== S.obraId) abrirObra(obraId);
  if (!obraId) fecharObra();
  S.aba = aba === "cronograma" ? "tarefas" : (aba || (modo === "rdo" ? "rdo" : "painel"));
  S.sub = resto;
  desenhar();
}

function abrirObra(id) {
  const aba = S.aba;
  fecharObra();
  S.obraId = id; S.aba = aba;
  const o = obraAtual();
  if (!o || o.legado) return;
  S.itensProntos = false;
  S.unsubItens = onSnapshot(COL("itens"), snap => {
    S.itens = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    S.itensProntos = true;
    S.modelo = derivar(S.itens);
    if (S.ui.inicializado !== id) { S.ui.inicializado = id; S.ui.abertos = new Set(S.modelo.raiz); }
    if (S.aba === "rdo" && S.rdoEd) desenharRdoEditor();
    else if (S.aba === "tarefas" || S.aba === "painel" || document.querySelector("#obra-corpo .go-carregando")) desenhar();
  }, e => { console.error("itens", e); S.itensProntos = true; desenharErro(e); });
  S.unsubRdos = onSnapshot(COL("rdos"), snap => {
    S.rdos = snap.docs.map(d => ({ id: d.id, ...d.data() })).sort((a, b) => b.data.localeCompare(a.data));
    if (S.aba === "painel" || (S.aba === "rdo" && !S.sub.length)) desenhar();
    else if (S.aba === "rdo" && S.sub[0] && !(S.rdoEd && S.rdoEd.dirty)) abrirRdoEditor(S.sub[0]);
  }, e => console.error("rdos", e));
}
function fecharObra() {
  S.unsubItens && S.unsubItens(); S.unsubRdos && S.unsubRdos();
  Object.assign(S, { obraId: null, unsubItens: null, unsubRdos: null, itens: [], rdos: [], modelo: null, itensProntos: false, rdoEd: null });
}

function desenharErro(e) {
  const c = $("obra-corpo");
  if (c) c.innerHTML = `<div class="obra-vazio"><div class="obra-vazio-ic">⚠️</div><h3>Não foi possível ler os dados da obra</h3><p>Motivo: ${esc(e && (e.code || e.message) || e)}. Se for “permission-denied”, confira se você é membro da obra e se as regras do Firestore foram publicadas.</p></div>`;
}

function desenhar() {
  const lista = $("obras-lista"), det = $("obra-detalhe");
  const o = obraAtual();
  if (!S.obraId || !o) {
    det.style.display = "none"; lista.style.display = "block";
    if (S.obraId && !o) { lista.innerHTML = `<div class="obra-vazio"><h3>Obra não encontrada ou sem acesso</h3><p><a href="#acompanhamento">Voltar para a lista de obras</a></p></div>`; return; }
    renderListaObras();
    return;
  }
  lista.style.display = "none"; det.style.display = "block";
  const p = papelAtual();
  $("obra-titulo").innerHTML = `${esc(o.nome)} ${o.demo ? '<span class="go-demo">DEMONSTRAÇÃO · dados fictícios</span>' : ""}`;
  $("obra-sub").innerHTML = [o.endereco && esc(o.endereco), `<span class="go-sit s-${o.situacao}">${esc(SITUACOES_OBRA[o.situacao] || "")}</span>`,
    p && `<span class="go-papel">Seu papel: ${esc(PAPEIS[p])}</span>`].filter(Boolean).join(" · ");
  const base = "#acompanhamento/" + encodeURIComponent(o.id);
  const blocos = (o.blocos || []).filter(b => b.pagina);
  const abas = o.legado
    ? [["mapa", "🗺️ Mapa de andamento"]]
    : [["painel", "📊 Painel"], ["tarefas", "📋 Tarefas (EAP)"], ["rdo", "📝 RDO"], ...(blocos.length ? [["mapa", "🗺️ Mapa"]] : []), ...(pode().admin ? [["cadastro", "⚙️ Cadastro e acesso"]] : [])];
  if (o.legado && S.aba !== "mapa") S.aba = "mapa";
  $("obra-tabs").innerHTML = abas.map(([k, t]) => `<a class="tab ${S.aba === k ? "active" : ""}" href="${base}/${k}${k === "mapa" && blocos[0] ? "/" + (S.sub[0] && S.aba === "mapa" ? S.sub[0] : blocos[0].id) : ""}">${t}</a>`).join("");
  const corpo = $("obra-corpo");
  if (S.aba !== "mapa") corpo.dataset.mapa = "";
  if (S.aba === "mapa") return renderMapa(o, S.sub[0]);
  if (S.aba === "cadastro") return pode().admin ? renderCadastro() : (corpo.innerHTML = `<div class="obra-vazio"><h3>Somente administradores</h3></div>`);
  if (!S.itensProntos) { corpo.innerHTML = `<div class="go-carregando">Carregando…</div>`; return; }
  if (S.aba === "painel") return renderPainel();
  if (S.aba === "tarefas") return renderTarefas();
  if (S.aba === "rdo") return S.sub[0] ? abrirRdoEditor(S.sub[0]) : renderRdoLista();
  if (S.aba === "cadastro") return pode().admin ? renderCadastro() : (corpo.innerHTML = `<div class="obra-vazio"><h3>Somente administradores</h3></div>`);
  corpo.innerHTML = "";
}

// ---------------------------------------------------------------------
// Lista de obras
// ---------------------------------------------------------------------
function renderListaObras() {
  const el = $("obras-lista");
  const admin = C.perfil() === "admin";
  const legadas = S.obras.filter(o => o.legado);
  const titulo = S.modoHub === "rdo" ? "Relatório Diário de Obra" : "Obras";
  const sub = S.modoHub === "rdo" ? "Escolha a obra para registrar ou aprovar o RDO." : "Escolha a obra para ver o painel, as tarefas, os RDOs e o mapa.";
  el.innerHTML = `
    <div class="obras-intro go-intro"><div><h2>${titulo}</h2><p>${sub}</p></div>
      ${admin ? `<div class="go-acoes"><button class="btn-export" onclick="GO.criarDemo()">🧪 Criar obra de demonstração</button><button class="btn-new" onclick="GO.novaObra()">+ Nova obra</button></div>` : ""}
    </div>
    ${admin && legadas.length ? `<div class="migr-banner"><div><b>Cadastrar ${legadas.length === 1 ? "a obra" : "as obras"} do sistema anterior</b>
      <span>${esc(legadas.map(o => o.nome).join(", "))} ainda ${legadas.length === 1 ? "está" : "estão"} só no arquivo obras.json. Ao cadastrar, nada do mapa é apagado — você define quem acessa e libera Painel, Tarefas e RDO.</span></div>
      <button class="btn-migr" onclick="GO.migrarLegadas()">Cadastrar agora</button></div>` : ""}
    ${S.obras.length ? `<div class="obras-grid">${S.obras.map(cardObra).join("")}</div>`
      : `<div class="obra-vazio"><div class="obra-vazio-ic">🏗️</div><h3>Nenhuma obra liberada para você</h3><p>Peça a um administrador para incluir seu usuário na equipe da obra.</p></div>`}`;
}
function cardObra(o) {
  const destino = S.modoHub === "rdo" && !o.legado ? "/rdo" : "";
  const p = papelNaObra(o, uid(), C.perfil());
  return `<a class="obra-card ${o.demo ? "go-card-demo" : ""}" href="#acompanhamento/${encodeURIComponent(o.id)}${destino}">
    <div class="obra-card-top">${o.demo ? '<span class="go-demo">DEMO</span>' : ""}<span class="go-sit s-${o.situacao}">${esc(SITUACOES_OBRA[o.situacao] || "")}</span></div>
    <div class="obra-nome">${esc(o.nome)}</div>
    <div class="obra-end">${esc(o.endereco || "")}</div>
    <dl class="obra-meta">
      <dt>Responsável</dt><dd>${esc(o.responsavelNome || nomeUsuario(o.responsavelUid) || "—")}</dd>
      <dt>Previsto</dt><dd>${o.inicioPrev || o.fimPrev ? `${fmtData(o.inicioPrev)} a ${fmtData(o.fimPrev)}` : "—"}</dd>
      <dt>Seu papel</dt><dd>${esc(PAPEIS[p] || "—")}</dd>
      ${o.legado ? `<dt>Situação</dt><dd><span class="obra-sub">Somente mapa — aguardando cadastro</span></dd>` : ""}
    </dl>
    <div class="obra-go">Abrir obra →</div>
  </a>`;
}

// ---------------------------------------------------------------------
// Modal genérico (usa as classes de modal do CRM)
// ---------------------------------------------------------------------
function garantirModal() {
  if ($("go-modal")) return;
  const d = document.createElement("div");
  d.className = "modal-bg"; d.id = "go-modal";
  d.innerHTML = `<div class="modal modal-lg" id="go-modal-box"></div>`;
  document.body.appendChild(d);
  const p = document.createElement("div");
  p.className = "modal-bg"; p.id = "go-pergunta"; p.style.zIndex = 150;
  p.innerHTML = `<div class="modal" style="max-width:440px" id="go-pergunta-box"></div>`;
  document.body.appendChild(p);
}
function abrirModal({ titulo, corpo, rodape, largura }) {
  const box = $("go-modal-box");
  box.style.maxWidth = largura || "";
  box.innerHTML = `<div class="mh"><div class="mt">${titulo}</div><button class="mc" onclick="GO.fecharModal()">×</button></div>
    <div class="fg go-fg">${corpo}</div>${rodape ? `<div class="mf">${rodape}</div>` : ""}`;
  $("go-modal").classList.add("open");
  box.scrollTop = 0;
}
function fecharModal() { $("go-modal").classList.remove("open"); }
/** Pergunta com campo de texto opcional. Resolve com o texto (ou "" sem campo) ou null se cancelar. */
function perguntar(titulo, texto, { campo = false, obrigatorio = false, botao = "Confirmar", perigo = false, placeholder = "" } = {}) {
  return new Promise(res => {
    const box = $("go-pergunta-box");
    box.innerHTML = `<div class="mh"><div class="mt">${esc(titulo)}</div></div>
      <div class="fg"><div style="font-size:13px;color:#444;line-height:1.5">${texto}</div>
      ${campo ? `<textarea id="go-perg-txt" placeholder="${esc(placeholder)}"></textarea><div class="go-erro" id="go-perg-err"></div>` : ""}</div>
      <div class="mf"><button class="btn-cancel" id="go-perg-nao">Cancelar</button><button class="${perigo ? "btn-danger" : "btn-save"}" id="go-perg-sim">${esc(botao)}</button></div>`;
    $("go-pergunta").classList.add("open");
    const fim = v => { $("go-pergunta").classList.remove("open"); res(v); };
    $("go-perg-nao").onclick = () => fim(null);
    $("go-perg-sim").onclick = () => {
      const v = campo ? $("go-perg-txt").value.trim() : "";
      if (campo && obrigatorio && v.length < 3) { $("go-perg-err").textContent = "Descreva o motivo (mínimo 3 caracteres)."; return; }
      fim(v);
    };
    if (campo) setTimeout(() => $("go-perg-txt").focus(), 50);
  });
}

// ---------------------------------------------------------------------
// Histórico (auditoria) e série de avanço
// ---------------------------------------------------------------------
function regHist(w, obraId, { alvoTipo, alvoId, acao, campos = null, resumo = "" }) {
  const ref = doc(collection(C.db, "obras", obraId, "historico"));
  w.set(ref, { alvoTipo, alvoId, acao, campos, resumo, autorUid: uid(), autorNome: meuNome(), em: serverTimestamp(), emISO: new Date().toISOString() });
}
function regSerie(w, obraId, itemId, antes, depois, dataRef) {
  if (antes === depois) return;
  const ref = doc(collection(C.db, "obras", obraId, "serie"));
  w.set(ref, { itemId, antes: antes ?? 0, depois: depois ?? 0, dataRef: dataRef || hojeISO(), autorUid: uid(), em: serverTimestamp() });
}
/** Diferenças campo a campo para o histórico. */
function diffCampos(antes, depois, campos) {
  const out = {};
  for (const k of campos) {
    const a = antes ? antes[k] ?? null : null, b = depois[k] ?? null;
    if (JSON.stringify(a) !== JSON.stringify(b)) out[k] = { antes: a, depois: b };
  }
  return Object.keys(out).length ? out : null;
}
const NOMES_CAMPO = { codigo: "Código", descricao: "Descrição", paiId: "Vínculo", responsavelUid: "Responsável", inicioPrev: "Início previsto", fimPrev: "Término previsto",
  inicioReal: "Início real", fimReal: "Término real", modo: "Modo", percentual: "Percentual", qtdPrevista: "Qtd. prevista", unidade: "Unidade", qtdExecutada: "Qtd. executada",
  peso: "Peso", observacoes: "Observações", anexos: "Anexos", nome: "Nome", endereco: "Endereço", situacao: "Situação", membros: "Equipe", status: "Situação do RDO", avanco: "Avanço (%)" };
function fmtValorHist(k, v) {
  if (v == null || v === "") return "—";
  if (/inicio|fim/i.test(k) && /^\d{4}-\d\d-\d\d$/.test(v)) return fmtData(v);
  if (k === "responsavelUid") return nomeUsuario(v) || v;
  if (k === "paiId") { const it = S.itens.find(x => x.id === v); return it ? it.codigo : v; }
  if (k === "anexos") return `${(v || []).length} arquivo(s)`;
  if (k === "membros") return `${Object.keys(v || {}).length} pessoa(s)`;
  if (k === "situacao") return SITUACOES_OBRA[v] || v;
  if (k === "status") return STATUS_RDO[v] || v;
  if (typeof v === "number") return fmtNum(v);
  return String(v);
}
async function carregarHistorico(obraId, alvoId) {
  const snap = await getDocs(query(collection(C.db, "obras", obraId, "historico"), where("alvoId", "==", alvoId)));
  return snap.docs.map(d => d.data()).sort((a, b) => String(b.emISO).localeCompare(String(a.emISO)));
}
function htmlHistorico(lista) {
  if (!lista.length) return `<div class="go-muted">Sem alterações registradas.</div>`;
  return `<div class="go-hist">${lista.map(h => `<div class="go-hist-i"><div class="go-hist-h"><b>${esc(h.autorNome || "—")}</b> · ${fmtDataHora(h.em || h.emISO)} · ${esc(h.resumo || h.acao)}</div>
    ${h.campos ? `<ul>${Object.entries(h.campos).map(([k, v]) => `<li>${esc(NOMES_CAMPO[k] || k)}: <s>${esc(fmtValorHist(k, v.antes))}</s> → <b>${esc(fmtValorHist(k, v.depois))}</b></li>`).join("")}</ul>` : ""}</div>`).join("")}</div>`;
}

// ---------------------------------------------------------------------
// Mapa de andamento (páginas existentes em iframe) + importação antiga
// ---------------------------------------------------------------------
function renderMapa(obra, blocoId) {
  const corpo = $("obra-corpo");
  const blocos = (obra.blocos || []).filter(b => b.pagina);
  const bloco = blocos.find(b => b.id === blocoId) || blocos[0];
  if (!bloco) { corpo.innerHTML = '<div class="obra-vazio"><h3>Nenhum pavimento com mapa ainda</h3></div>'; return; }
  const atual = $("obra-frame");
  if (!(atual && atual.dataset.src === bloco.pagina && corpo.dataset.mapa === "1")) {
    corpo.innerHTML = `<div id="obra-migracao"></div><iframe id="obra-frame" class="obra-frame" data-src="${esc(bloco.pagina)}" src="${esc(bloco.pagina)}" title="Mapa de andamento — ${esc(obra.nome)} ${esc(bloco.nome)}"></iframe>`;
    corpo.dataset.mapa = "1";
  }
  verificarMigracaoMapa(obra);
}
async function verificarMigracaoMapa(obra) {
  const box = $("obra-migracao");
  if (!box) return;
  box.innerHTML = "";
  if (!obra.migracao || C.perfil() !== "admin") return;
  try {
    const flag = await getDoc(doc(C.db, "obras", obra.id, "cfg", "migracao"));
    if (flag.exists()) return;
  } catch (e) {
    box.innerHTML = `<div class="migr-banner erro"><div><b>Não consegui ler o banco desta obra</b><span>Motivo: ${esc(e && (e.code || e.message) || e)}.</span></div></div>`;
    return;
  }
  box.innerHTML = `<div class="migr-banner"><div><b>Importar o andamento já marcado</b>
    <span>${obra.migracao.registros} registros da página antiga deste mapa ainda não foram trazidos para o sistema.</span></div>
    <button class="btn-migr" id="btn-migr">Importar agora</button></div>`;
  $("btn-migr").onclick = () => importarMigracaoMapa(obra);
}
async function importarMigracaoMapa(obra) {
  const btn = $("btn-migr");
  if (!confirm(`Importar ${obra.migracao.registros} registros de andamento para ${obra.nome}?\n\nO que já estiver marcado no sistema é mantido; os registros importados são somados por cima.`)) return;
  btn.disabled = true; btn.textContent = "Importando…";
  try {
    const r = await fetch(obra.migracao.arquivo, { cache: "no-cache" });
    const pacote = await r.json();
    const entradas = Object.entries(pacote.docs);
    for (let i = 0; i < entradas.length; i += 400) {
      const b = writeBatch(C.db);
      entradas.slice(i, i + 400).forEach(([id, data]) => b.set(doc(C.db, "obras", obra.id, pacote.colecao, id), data, { merge: true }));
      await b.commit();
    }
    await setDoc(doc(C.db, "obras", obra.id, "cfg", "migracao"), { feito: true, registros: entradas.length, arquivo: obra.migracao.arquivo, por: uid(), em: serverTimestamp() });
    $("obra-migracao").innerHTML = "";
    toast(`${entradas.length} registros importados`);
  } catch (e) {
    console.error(e); btn.disabled = false; btn.textContent = "Tentar de novo"; toast(erroMsg(e));
  }
}

// ---------------------------------------------------------------------
// Cadastro de obras, equipe e acesso (Administrador)
// ---------------------------------------------------------------------
function optsUsuarios(sel, { vazio = "— sem responsável —", lista = S.usuarios } = {}) {
  return `<option value="">${vazio}</option>` + lista.map(u => `<option value="${esc(u.id)}" ${u.id === sel ? "selected" : ""}>${esc(u.nome || u.email)}</option>`).join("");
}
function formObraHtml(o = {}) {
  return `<div class="go-grid2">
    <div class="fgf go-span2"><label class="fl">Nome da obra *</label><input id="ob-nome" value="${esc(o.nome || "")}" maxlength="120"></div>
    <div class="fgf go-span2"><label class="fl">Descrição</label><textarea id="ob-desc">${esc(o.descricao || o.escopo || "")}</textarea></div>
    <div class="fgf go-span2"><label class="fl">Endereço</label><input id="ob-end" value="${esc(o.endereco || "")}"></div>
    <div class="fgf"><label class="fl">Responsável</label><select id="ob-resp">${optsUsuarios(o.responsavelUid)}</select></div>
    <div class="fgf"><label class="fl">Situação</label><select id="ob-sit">${Object.entries(SITUACOES_OBRA).map(([k, v]) => `<option value="${k}" ${k === (o.situacao || "nao_iniciada") ? "selected" : ""}>${v}</option>`).join("")}</select></div>
    <div class="fgf"><label class="fl">Início previsto</label><input type="date" id="ob-ini" value="${esc(o.inicioPrev || "")}"></div>
    <div class="fgf"><label class="fl">Término previsto</label><input type="date" id="ob-fim" value="${esc(o.fimPrev || "")}"></div>
  </div><div class="go-erro" id="ob-err"></div>`;
}
function lerFormObra() {
  const d = {
    nome: $("ob-nome").value.trim(), descricao: $("ob-desc").value.trim(), endereco: $("ob-end").value.trim(),
    responsavelUid: $("ob-resp").value || null, situacao: $("ob-sit").value, inicioPrev: $("ob-ini").value || null, fimPrev: $("ob-fim").value || null
  };
  d.responsavelNome = d.responsavelUid ? nomeUsuario(d.responsavelUid) : "";
  if (!d.nome) return { erro: "Informe o nome da obra." };
  if (d.inicioPrev && d.fimPrev && d.fimPrev < d.inicioPrev) return { erro: "O término previsto é anterior ao início." };
  return { d };
}
function membrosHtml(o, padraoLegado = false) {
  return `<div class="go-membros">${S.usuarios.map(u => {
    const role = u.role || "operador";
    let atual = (o.membros || {})[u.id] || "";
    if (padraoLegado && !o.membros) atual = role === "operador" ? "colaborador" : role === "visualizador" ? "consulta" : "";
    const sel = role === "admin" ? `<span class="go-muted">Administrador (acesso total)</span>`
      : `<select data-membro="${esc(u.id)}"><option value="">Sem acesso</option>${(role === "visualizador" ? ["consulta"] : ["gestor", "colaborador", "consulta"])
        .map(p => `<option value="${p}" ${atual === p ? "selected" : ""}>${PAPEIS[p]}</option>`).join("")}</select>`;
    return `<div class="go-membro"><div><b>${esc(u.nome || u.email)}</b><span>${esc(u.email || "")} · perfil ${esc(role)}</span></div>${sel}</div>`;
  }).join("") || '<div class="go-muted">Nenhum usuário cadastrado.</div>'}</div>
  <div class="go-dica">Gestor: acompanha a obra, edita a EAP e aprova RDOs · Colaborador: registra RDOs e o andamento · Consulta: só visualiza. Usuários com perfil “visualizador” só podem ser Consulta.</div>`;
}
function lerMembros(raiz) {
  const membros = {};
  raiz.querySelectorAll("[data-membro]").forEach(s => { if (s.value) membros[s.dataset.membro] = s.value; });
  return { membros, membrosUids: Object.keys(membros) };
}

async function novaObra() {
  abrirModal({ titulo: "Nova obra", corpo: formObraHtml(), rodape: `<button class="btn-cancel" onclick="GO.fecharModal()">Cancelar</button><button class="btn-save" id="ob-salvar">Cadastrar obra</button>` });
  $("ob-salvar").onclick = async () => {
    const { d, erro } = lerFormObra();
    if (erro) { $("ob-err").textContent = erro; return; }
    $("ob-salvar").disabled = true;
    try {
      let id = slug(d.nome), n = 2;
      while ((await getDoc(doc(C.db, "obras", id))).exists()) id = `${slug(d.nome)}-${n++}`;
      const b = writeBatch(C.db);
      b.set(doc(C.db, "obras", id), { ...d, membros: {}, membrosUids: [], blocos: [], criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp() });
      regHist(b, id, { alvoTipo: "obra", alvoId: id, acao: "criar", resumo: "Obra cadastrada" });
      await b.commit();
      fecharModal(); toast("Obra cadastrada — defina a equipe");
      location.hash = `#acompanhamento/${id}/cadastro`;
    } catch (e) { $("ob-err").textContent = erroMsg(e); $("ob-salvar").disabled = false; }
  };
}

async function migrarLegadas() {
  const legadas = S.obras.filter(o => o.legado);
  abrirModal({
    titulo: "Cadastrar obras do sistema anterior",
    corpo: `<p class="go-dica" style="margin:0">Os dados do mapa (andamento por ambiente) ficam como estão. A equipe abaixo já vem preenchida para que ninguém perca o acesso que tem hoje: operadores como Colaborador e visualizadores como Consulta. Ajuste se quiser.</p>
      ${legadas.map(o => `<div class="go-box" data-obra="${esc(o.id)}"><div class="go-box-t">${esc(o.nome)}</div>${membrosHtml(o, true)}</div>`).join("")}<div class="go-erro" id="mg-err"></div>`,
    rodape: `<button class="btn-cancel" onclick="GO.fecharModal()">Cancelar</button><button class="btn-save" id="mg-ok">Cadastrar</button>`
  });
  $("mg-ok").onclick = async () => {
    $("mg-ok").disabled = true;
    try {
      for (const o of legadas) {
        const raiz = document.querySelector(`[data-obra="${CSS.escape(o.id)}"]`);
        const { membros, membrosUids } = lerMembros(raiz);
        const b = writeBatch(C.db);
        b.set(doc(C.db, "obras", o.id), {
          nome: o.nome, descricao: o.escopo || "", endereco: o.endereco || "", projeto: o.projeto || "", escopo: o.escopo || "",
          situacao: mapSituacaoLegado(o.status), responsavelUid: null, responsavelNome: "", inicioPrev: null, fimPrev: null,
          blocos: o.blocos || [], migracao: o.migracao || null, membros, membrosUids, criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp()
        }, { merge: true });
        regHist(b, o.id, { alvoTipo: "obra", alvoId: o.id, acao: "criar", resumo: "Obra cadastrada a partir do obras.json" });
        await b.commit();
      }
      fecharModal(); toast("Obras cadastradas");
    } catch (e) { $("mg-err").textContent = erroMsg(e); $("mg-ok").disabled = false; }
  };
}

async function renderCadastro() {
  const o = obraAtual();
  const corpo = $("obra-corpo");
  corpo.innerHTML = `<div class="go-cols">
    <section class="go-sec"><h3>Dados da obra</h3>${formObraHtml(o)}<div class="go-row-end"><button class="btn-save" id="ob-salvar">Salvar dados</button></div></section>
    <section class="go-sec"><h3>Equipe e acesso</h3><div id="ob-membros">${membrosHtml(o)}</div><div class="go-erro" id="mb-err"></div><div class="go-row-end"><button class="btn-save" id="mb-salvar">Salvar equipe</button></div></section>
  </div>
  ${o.demo ? `<section class="go-sec go-perigo"><h3>Obra de demonstração</h3><p class="go-dica">Esta obra tem dados fictícios. Apagar remove a obra, as tarefas, os RDOs, os arquivos e o histórico dela.</p><button class="btn-danger" onclick="GO.apagarDemo()">Apagar obra de demonstração</button></section>` : ""}
  <section class="go-sec"><h3>Histórico do cadastro</h3><div id="ob-hist" class="go-muted">Carregando…</div></section>`;
  $("ob-salvar").onclick = async () => {
    const { d, erro } = lerFormObra();
    if (erro) { $("ob-err").textContent = erro; return; }
    try {
      const b = writeBatch(C.db);
      b.update(doc(C.db, "obras", o.id), { ...d, atualizadoEm: serverTimestamp() });
      regHist(b, o.id, { alvoTipo: "obra", alvoId: o.id, acao: "editar", resumo: "Dados da obra alterados", campos: diffCampos(o, d, ["nome", "descricao", "endereco", "responsavelUid", "situacao", "inicioPrev", "fimPrev"]) });
      await b.commit(); toast("Dados salvos");
    } catch (e) { $("ob-err").textContent = erroMsg(e); }
  };
  $("mb-salvar").onclick = async () => {
    const m = lerMembros($("ob-membros"));
    try {
      const b = writeBatch(C.db);
      b.update(doc(C.db, "obras", o.id), { ...m, atualizadoEm: serverTimestamp() });
      regHist(b, o.id, { alvoTipo: "obra", alvoId: o.id, acao: "acesso", resumo: "Equipe e acesso alterados", campos: diffCampos({ membros: o.membros || {} }, { membros: m.membros }, ["membros"]) });
      await b.commit(); toast("Equipe salva");
    } catch (e) { $("mb-err").textContent = erroMsg(e); }
  };
  try { $("ob-hist").innerHTML = htmlHistorico(await carregarHistorico(o.id, o.id)); } catch (e) { $("ob-hist").textContent = erroMsg(e); }
}

// ---------------------------------------------------------------------
// Arquivos (fotos comprimidas e anexos pequenos guardados no Firestore)
// ---------------------------------------------------------------------
const MAX_ANEXO = 700 * 1024;
function lerDataUrl(file) { return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file); }); }
function carregarImagem(src) { return new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; }); }
async function reduzir(img, lado, qualidade) {
  const k = Math.min(1, lado / Math.max(img.width, img.height));
  const cv = document.createElement("canvas");
  cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
  const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, cv.width, cv.height);
  g.drawImage(img, 0, 0, cv.width, cv.height);
  return cv.toDataURL("image/jpeg", qualidade);
}
async function salvarArquivo(file, contexto) {
  let dados, tipo = file.type || "application/octet-stream", thumb = null, nome = file.name;
  if (/^image\//.test(tipo)) {
    const img = await carregarImagem(await lerDataUrl(file));
    let lado = 1600, q = 0.75;
    dados = await reduzir(img, lado, q);
    while (dados.length > 900000 && lado > 600) { lado -= 300; q = Math.max(0.5, q - 0.08); dados = await reduzir(img, lado, q); }
    thumb = await reduzir(img, 240, 0.6);
    tipo = "image/jpeg"; nome = nome.replace(/\.[a-z0-9]+$/i, "") + ".jpg";
  } else {
    if (file.size > MAX_ANEXO) throw new Error(`“${file.name}” tem ${fmtNum(file.size / 1024 / 1024)} MB. O limite de anexo é 700 KB (fotos são comprimidas automaticamente).`);
    dados = await lerDataUrl(file);
  }
  const ref = doc(COL("arquivos"));
  const meta = { id: ref.id, nome, tipo, tamanho: Math.round(dados.length * 0.75) };
  await setDoc(ref, { ...meta, dados, contexto, autorUid: uid(), autorNome: meuNome(), em: serverTimestamp() });
  return thumb ? { ...meta, thumb } : meta;
}
async function baixarArquivo(id) {
  try {
    const d = await getDoc(R("arquivos", id));
    if (!d.exists()) { toast("Arquivo não encontrado"); return; }
    const a = document.createElement("a"); a.href = d.data().dados; a.download = d.data().nome || "arquivo";
    document.body.appendChild(a); a.click(); a.remove();
  } catch (e) { toast(erroMsg(e)); }
}
const htmlAnexos = (lista, { remover = false, onRemover = "" } = {}) => (lista || []).length
  ? `<div class="go-anexos">${lista.map(a => `<div class="go-anexo">${a.thumb ? `<img src="${a.thumb}" alt="">` : `<span class="go-anexo-ic">📎</span>`}
      <div class="go-anexo-n"><a href="javascript:void 0" onclick="GO.baixarArquivo('${esc(a.id)}')">${esc(a.nome)}</a><span>${fmtNum((a.tamanho || 0) / 1024)} KB${a.legenda ? " · " + esc(a.legenda) : ""}</span></div>
      ${remover ? `<button class="card-btn del" title="Remover" onclick="${onRemover}('${esc(a.id)}')">🗑</button>` : ""}</div>`).join("")}</div>`
  : `<div class="go-muted">Nenhum arquivo.</div>`;

// ---------------------------------------------------------------------
// Lista de tarefas (EAP)
// ---------------------------------------------------------------------
const semAcento = s => String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
function filtroTarefaAtivo() { const u = S.ui; return !!(u.busca || u.situacao || u.resp || u.prazo || u.ocultarConcluidas); }
function casaTarefa(d, hoje = hojeISO()) {
  const u = S.ui;
  if (u.etapa && d.etapaId !== u.etapa) return false;
  if (u.busca) { const b = semAcento(u.busca); if (!semAcento(d.codigo).includes(b) && !semAcento(d.descricao).includes(b)) return false; }
  if (u.situacao) { if (u.situacao === "atrasada" ? !d.atrasada : d.situacao !== u.situacao) return false; }
  if (u.resp) { if (u.resp === "_sem" ? !!d.responsavelUid : d.responsavelUid !== u.resp) return false; }
  if (u.ocultarConcluidas && d.situacao === "concluida") return false;
  if (u.prazo) {
    const f = d.fimPrev;
    if (u.prazo === "sem") { if (f) return false; }
    else if (u.prazo === "atrasadas") { if (!d.atrasada) return false; }
    else { const lim = somaDias(hoje, Number(u.prazo)); if (!f || f < hoje || f > lim || d.situacao === "concluida") return false; }
  }
  return true;
}
function linhasVisiveis() {
  const M = S.modelo, u = S.ui, ativo = filtroTarefaAtivo(), rows = [];
  const hoje = hojeISO();
  const cache = new Map();
  const vis = d => {
    if (cache.has(d.id)) return cache.get(d.id);
    let v;
    if (d.tipo === "tarefa") v = casaTarefa(d, hoje);
    else if (u.etapa && d.etapaId !== u.etapa) v = false;
    else if (!ativo) v = true;
    else {
      v = d.tarefas.some(t => casaTarefa(t, hoje));
      if (!v && u.busca && !u.situacao && !u.resp && !u.prazo) { const b = semAcento(u.busca); v = semAcento(d.codigo).includes(b) || semAcento(d.descricao).includes(b); }
      if (v && u.ocultarConcluidas && d.situacao === "concluida" && !d.tarefas.some(t => casaTarefa(t, hoje))) v = false;
    }
    cache.set(d.id, v); return v;
  };
  const walk = ids => {
    for (const id of ids) {
      const d = M.der.get(id);
      if (!d || !vis(d)) continue;
      if (u.soEtapas && d.tipo === "tarefa") continue;
      rows.push(d);
      if (d.filhos.length && (ativo || u.etapa || S.ui.abertos.has(d.id))) walk(d.filhos);
    }
  };
  walk(M.raiz);
  return rows;
}

function renderTarefas() {
  const P = pode(), u = S.ui, M = S.modelo;
  const etapas = M.raiz.map(id => M.der.get(id)).filter(d => d.tipo === "etapa");
  const resps = [...new Set(S.itens.map(i => i.responsavelUid).filter(Boolean))];
  $("obra-corpo").innerHTML = `
  <div class="go-toolbar">
    <input class="go-busca" id="tf-busca" placeholder="🔎 Buscar por código ou descrição" value="${esc(u.busca)}">
    <select id="tf-sit"><option value="">Todas as situações</option>${Object.entries(SITUACOES_TAREFA).map(([k, v]) => `<option value="${k}" ${u.situacao === k ? "selected" : ""}>${v}</option>`).join("")}<option value="atrasada" ${u.situacao === "atrasada" ? "selected" : ""}>Atrasadas</option></select>
    <select id="tf-resp"><option value="">Todos os responsáveis</option><option value="_sem" ${u.resp === "_sem" ? "selected" : ""}>Sem responsável</option>${resps.map(r => `<option value="${esc(r)}" ${u.resp === r ? "selected" : ""}>${esc(nomeUsuario(r) || r)}</option>`).join("")}</select>
    <select id="tf-etapa"><option value="">Todas as etapas</option>${etapas.map(e => `<option value="${esc(e.id)}" ${u.etapa === e.id ? "selected" : ""}>${esc(e.codigo + " · " + e.descricao)}</option>`).join("")}</select>
    <select id="tf-prazo"><option value="">Qualquer prazo</option><option value="atrasadas" ${u.prazo === "atrasadas" ? "selected" : ""}>Prazo vencido</option><option value="7" ${u.prazo === "7" ? "selected" : ""}>Vence em 7 dias</option><option value="30" ${u.prazo === "30" ? "selected" : ""}>Vence em 30 dias</option><option value="sem" ${u.prazo === "sem" ? "selected" : ""}>Sem prazo</option></select>
    <label class="go-chk"><input type="checkbox" id="tf-soetapas" ${u.soEtapas ? "checked" : ""}> Somente etapas</label>
    <label class="go-chk"><input type="checkbox" id="tf-ocultar" ${u.ocultarConcluidas ? "checked" : ""}> Ocultar concluídas</label>
  </div>
  <div class="go-toolbar go-toolbar2">
    <button class="btn-sm" onclick="GO.expandir(true)">⊞ Expandir tudo</button><button class="btn-sm" onclick="GO.expandir(false)">⊟ Recolher tudo</button>
    <button class="btn-sm" onclick="GO.limparFiltros()">Limpar filtros</button>
    <span class="go-flex"></span>
    ${P.gerir ? `<button class="btn-export" onclick="GO.importar()">⬆️ Importar planilha</button>` : ""}
    <button class="btn-export" onclick="GO.exportarXlsx()">⬇️ Excel</button>
    <button class="btn-export" onclick="GO.exportarPdf()">⬇️ PDF</button>
    ${P.gerir ? `<button class="btn-export" onclick="GO.abrirItem(null,'etapa')">+ Etapa</button><button class="btn-new" onclick="GO.abrirItem(null,'tarefa')">+ Tarefa</button>` : ""}
  </div>
  <div class="go-tabela">
    <div class="go-tr go-th"><div>Código</div><div>Descrição</div><div>Responsável</div><div>Previsto</div><div>Real</div><div>Avanço</div><div>Produção</div><div>Situação</div><div></div></div>
    <div id="tf-linhas"></div>
  </div>
  <div class="go-rodape-lista" id="tf-rodape"></div>`;
  const on = (id, ev, fn) => $(id).addEventListener(ev, fn);
  let t = null;
  on("tf-busca", "input", e => { clearTimeout(t); t = setTimeout(() => { u.busca = e.target.value.trim(); u.limite = 600; renderLinhas(); }, 180); });
  on("tf-sit", "change", e => { u.situacao = e.target.value; u.limite = 600; renderLinhas(); });
  on("tf-resp", "change", e => { u.resp = e.target.value; u.limite = 600; renderLinhas(); });
  on("tf-etapa", "change", e => { u.etapa = e.target.value; u.limite = 600; renderLinhas(); });
  on("tf-prazo", "change", e => { u.prazo = e.target.value; u.limite = 600; renderLinhas(); });
  on("tf-soetapas", "change", e => { u.soEtapas = e.target.checked; renderLinhas(); });
  on("tf-ocultar", "change", e => { u.ocultarConcluidas = e.target.checked; renderLinhas(); });
  renderLinhas();
}

function badgeSit(d) {
  const s = `<span class="go-chip c-${d.situacao}">${SITUACOES_TAREFA[d.situacao]}</span>`;
  return d.atrasada ? s + `<span class="go-chip c-atraso" title="Término previsto vencido e avanço abaixo de 100%">${d.tipo === "tarefa" ? "Atrasada" : d.cons.atrasadas + " atras."}</span>` : s;
}
function barra(av, extra = "") {
  const v = av == null ? 0 : av;
  return `<div class="go-bar ${extra}"><span style="width:${Math.max(0, Math.min(100, v))}%"></span></div>`;
}
function celAvanco(d) {
  if (d.tipo !== "tarefa") {
    const c = d.cons, av = c.ponderado ?? c.mediaSimples;
    if (!c.total) return `<span class="go-muted">sem tarefas</span>`;
    return `${barra(av)}<span class="go-pct" title="${c.ponderado != null ? "Avanço ponderado pelos pesos das tarefas" : "Média simples do avanço das tarefas (há tarefas sem peso)"}">${fmtPct(av)}${c.ponderado != null ? "" : "*"}</span>`;
  }
  const k = d.calc;
  if (k.semPrevisto && k.exe > 0) return `<span class="go-alerta" title="Quantidade prevista igual a zero — não é possível calcular o avanço">⚠️ sem previsto</span>`;
  return `${barra(k.avanco, k.acima ? "acima" : "")}<span class="go-pct">${fmtPct(k.acima ? k.real : k.avanco)}</span>${k.acima ? `<span class="go-alerta" title="Produção acima da quantidade prevista">▲</span>` : ""}`;
}
function celProducao(d) {
  if (d.tipo !== "tarefa") return `<span class="go-muted">${d.cons.concluida}/${d.cons.total} concl.</span>`;
  if ((d.modo || "manual") !== "producao") return `<span class="go-muted">% manual</span>`;
  return `<span class="${d.calc.acima ? "go-acima" : ""}">${fmtNum(d.qtdExecutada || 0)} / ${fmtNum(d.qtdPrevista || 0)} ${esc(d.unidade || "")}</span>`;
}
function linhaHtml(d, P) {
  const tem = d.filhos.length, aberto = S.ui.abertos.has(d.id) || filtroTarefaAtivo() || S.ui.etapa;
  const ini = d.tipo === "tarefa" ? d.inicioPrev : d.inicioCalc, fim = d.tipo === "tarefa" ? d.fimPrev : d.fimCalc;
  return `<div class="go-tr t-${d.tipo}" data-id="${esc(d.id)}" onclick="GO.abrirItem('${esc(d.id)}')">
    <div class="go-cod" style="padding-left:${d.nivel * 16}px">${tem ? `<button class="go-caret" onclick="event.stopPropagation();GO.alternar('${esc(d.id)}')">${aberto ? "▾" : "▸"}</button>` : `<span class="go-caret-v"></span>`}<span>${esc(d.codigo)}</span></div>
    <div class="go-desc"><span class="go-tipo">${TIPOS[d.tipo]}</span>${esc(d.descricao)}${(d.anexos || []).length ? ` <span class="go-muted" title="Anexos">📎${d.anexos.length}</span>` : ""}</div>
    <div class="go-resp" data-l="Responsável">${esc(d.responsavelNome || nomeUsuario(d.responsavelUid) || "—")}</div>
    <div class="go-dt" data-l="Previsto">${fmtData(ini)} – ${fmtData(fim)}</div>
    <div class="go-dt" data-l="Real">${d.tipo === "tarefa" ? `${fmtData(d.inicioReal)} – ${fmtData(d.fimReal)}` : ""}</div>
    <div class="go-av" data-l="Avanço">${celAvanco(d)}</div>
    <div class="go-prod" data-l="Produção">${celProducao(d)}</div>
    <div class="go-sitc">${badgeSit(d)}</div>
    <div class="go-acoes-l" onclick="event.stopPropagation()">${P.gerir ? `<button class="card-btn" title="Subir" onclick="GO.mover('${esc(d.id)}',-1)">↑</button><button class="card-btn" title="Descer" onclick="GO.mover('${esc(d.id)}',1)">↓</button>${d.tipo !== "tarefa" ? `<button class="card-btn" title="Incluir dentro" onclick="GO.abrirItem(null,'${d.tipo === "etapa" ? "subetapa" : "tarefa"}','${esc(d.id)}')">＋</button>` : ""}` : ""}</div>
  </div>`;
}
function renderLinhas() {
  const box = $("tf-linhas"); if (!box) return;
  const P = pode(), rows = linhasVisiveis(), lim = S.ui.limite;
  const M = S.modelo;
  box.innerHTML = rows.length ? rows.slice(0, lim).map(d => linhaHtml(d, P)).join("")
    : `<div class="go-vazio-l">${S.itens.length ? "Nenhum item atende aos filtros." : (P.gerir ? "A obra ainda não tem etapas. Use “+ Etapa” ou importe uma planilha." : "A obra ainda não tem etapas cadastradas.")}</div>`;
  const nt = rows.filter(r => r.tipo === "tarefa").length;
  $("tf-rodape").innerHTML = `${rows.length > lim ? `<button class="btn-sm" onclick="GO.maisLinhas()">Mostrar mais ${Math.min(600, rows.length - lim)} de ${rows.length - lim} restantes</button>` : ""}
    <span>${nt} de ${M.tarefas.length} tarefas exibidas · ${S.itens.length - M.tarefas.length} etapas/subetapas · * média simples (sem pesos em todas as tarefas)</span>`;
}

// ---------------------------------------------------------------------
// Cadastro/edição de etapa, subetapa e tarefa
// ---------------------------------------------------------------------
const CAMPOS_COLAB = ["percentual", "inicioReal", "fimReal", "observacoes", "anexos"];
const CAMPOS_ITEM = ["tipo", "codigo", "descricao", "paiId", "responsavelUid", "inicioPrev", "fimPrev", "inicioReal", "fimReal", "modo", "percentual", "qtdPrevista", "unidade", "peso", "observacoes"];
function opcoesPai(tipo, sel, proprioId) {
  if (tipo === "etapa") return `<option value="">— (nível principal)</option>`;
  const M = S.modelo, ok = tipo === "subetapa" ? ["etapa"] : ["etapa", "subetapa"];
  const proibidos = new Set();
  if (proprioId && M.der.get(proprioId)) { const marca = id => { proibidos.add(id); (M.der.get(id)?.filhos || []).forEach(marca); }; marca(proprioId); }
  const lista = []; const walk = ids => ids.forEach(id => { const d = M.der.get(id); if (!d || d.tipo === "tarefa") return; if (ok.includes(d.tipo) && !proibidos.has(id)) lista.push(d); walk(d.filhos); });
  walk(M.raiz);
  return `<option value="">— escolha —</option>` + lista.map(d => `<option value="${esc(d.id)}" ${d.id === sel ? "selected" : ""}>${"  ".repeat(d.nivel)}${esc(d.codigo + " · " + d.descricao)}</option>`).join("");
}
function sugerirCodigo(paiId) {
  const M = S.modelo, irmaos = paiId ? (M.der.get(paiId)?.filhos || []).map(id => M.der.get(id)) : M.raiz.map(id => M.der.get(id));
  const base = paiId ? M.der.get(paiId).codigo + "." : "";
  let n = irmaos.length + 1;
  const usados = new Set(S.itens.map(i => i.codigoKey));
  while (usados.has(codigoKey(base + n))) n++;
  return base + n;
}

function abrirItem(id, tipoNovo = "tarefa", paiNovo = "") {
  const P = pode();
  const it = id ? S.itens.find(i => i.id === id) : null;
  if (id && !it) return;
  if (!id && !P.gerir) return;
  const d = it ? S.modelo.der.get(id) : null;
  const tipo = it ? it.tipo : tipoNovo;
  const v = it || { tipo, paiId: paiNovo || "", modo: "manual", percentual: 0, codigo: sugerirCodigo(paiNovo || null) };
  const ger = P.gerir, reg = P.registrar;
  const dis = campo => (ger || (reg && CAMPOS_COLAB.includes(campo))) ? "" : "disabled";
  const modo = v.modo || "manual";
  const calc = it && tipo === "tarefa" ? calcAvanco(it) : null;
  const corpo = `
    <div class="go-grid2">
      ${!it ? `<div class="fgf"><label class="fl">Tipo</label><select id="it-tipo">${Object.entries(TIPOS).map(([k, t]) => `<option value="${k}" ${k === tipo ? "selected" : ""}>${t}</option>`).join("")}</select></div>` : `<div class="fgf"><label class="fl">Tipo</label><input value="${TIPOS[tipo]}" disabled></div>`}
      <div class="fgf"><label class="fl">Vinculado a ${tipo === "etapa" ? "" : "*"}</label><select id="it-pai" ${dis("paiId")}>${opcoesPai(tipo, v.paiId, id)}</select></div>
      <div class="fgf"><label class="fl">Código *</label><input id="it-cod" value="${esc(v.codigo || "")}" maxlength="40" ${dis("codigo")}></div>
      <div class="fgf"><label class="fl">Responsável</label><select id="it-resp" ${dis("responsavelUid")}>${optsUsuarios(v.responsavelUid)}</select></div>
      <div class="fgf go-span2"><label class="fl">Descrição *</label><input id="it-desc" value="${esc(v.descricao || "")}" maxlength="300" ${dis("descricao")}></div>
      <div class="fgf"><label class="fl">Início previsto</label><input type="date" id="it-ini" value="${esc(v.inicioPrev || "")}" ${dis("inicioPrev")}></div>
      <div class="fgf"><label class="fl">Término previsto</label><input type="date" id="it-fim" value="${esc(v.fimPrev || "")}" ${dis("fimPrev")}></div>
      ${tipo === "tarefa" ? `
      <div class="fgf"><label class="fl">Início real</label><input type="date" id="it-inir" value="${esc(v.inicioReal || "")}" ${dis("inicioReal")}></div>
      <div class="fgf"><label class="fl">Término real</label><input type="date" id="it-fimr" value="${esc(v.fimReal || "")}" ${dis("fimReal")}></div>` : ""}
    </div>
    ${tipo === "tarefa" ? `
    <div class="go-box"><div class="go-box-t">Acompanhamento</div>
      <div class="go-radios"><label><input type="radio" name="it-modo" value="manual" ${modo === "manual" ? "checked" : ""} ${dis("modo")}> Percentual informado</label>
      <label><input type="radio" name="it-modo" value="producao" ${modo === "producao" ? "checked" : ""} ${dis("modo")}> Produção (quantidade)</label></div>
      <div id="it-manual" class="go-grid2" style="${modo === "manual" ? "" : "display:none"}">
        <div class="fgf"><label class="fl">Percentual executado (0 a 100%)</label><input id="it-pct" inputmode="decimal" value="${esc(String(v.percentual ?? 0).replace(".", ","))}" ${dis("percentual")}></div>
      </div>
      <div id="it-prod" style="${modo === "producao" ? "" : "display:none"}">
        <div class="go-grid3">
          <div class="fgf"><label class="fl">Quantidade prevista *</label><input id="it-qprev" inputmode="decimal" value="${v.qtdPrevista != null ? esc(String(v.qtdPrevista).replace(".", ",")) : ""}" ${dis("qtdPrevista")}></div>
          <div class="fgf"><label class="fl">Unidade *</label><input id="it-un" list="go-unidades" value="${esc(v.unidade || "")}" maxlength="10" ${dis("unidade")}></div>
          <div class="fgf"><label class="fl">Executado acumulado</label><input value="${fmtNum(v.qtdExecutada || 0)}" disabled></div>
        </div>
        <datalist id="go-unidades">${UNIDADES.map(u => `<option value="${u}">`).join("")}</datalist>
        <div class="go-dica">A quantidade executada vem dos RDOs aprovados${ger ? " e dos ajustes do gestor" : ""}. Avanço = executado ÷ previsto × 100.</div>
        ${calc && calc.semPrevisto ? `<div class="go-alerta-box">⚠️ Quantidade prevista igual a zero: o avanço não pode ser calculado. Executado até agora: ${fmtNum(calc.exe)} ${esc(v.unidade || "")}.</div>` : ""}
        ${calc && calc.acima ? `<div class="go-alerta-box">▲ Produção acima do previsto: ${fmtNum(calc.exe)} de ${fmtNum(calc.prev)} ${esc(v.unidade || "")} (${fmtPct(calc.real)}). Para situação e consolidação a tarefa conta como 100%.</div>` : ""}
        ${it && ger && modo === "producao" ? `<button class="btn-sm" style="margin-top:6px" onclick="GO.ajustarProducao('${esc(id)}')">± Lançar ou ajustar produção</button>` : ""}
      </div>
      <div class="go-grid2" style="margin-top:8px"><div class="fgf"><label class="fl">Peso para o avanço físico (opcional)</label><input id="it-peso" inputmode="decimal" value="${v.peso != null ? esc(String(v.peso).replace(".", ",")) : ""}" placeholder="ex.: valor orçado ou % da obra" ${dis("peso")}></div>
      ${calc ? `<div class="fgf"><label class="fl">Avanço atual</label><div class="go-av-atual">${barra(calc.avanco, calc.acima ? "acima" : "")} <b>${calc.avanco == null ? "—" : fmtPct(calc.acima ? calc.real : calc.avanco)}</b> · ${SITUACOES_TAREFA[d.situacao]}${d.atrasada ? ' · <span class="go-chip c-atraso">Atrasada</span>' : ""}</div></div>` : ""}</div>
    </div>` : (d ? `<div class="go-box"><div class="go-box-t">Consolidado</div>${resumoCons(d.cons)}</div>` : "")}
    <div class="fgf"><label class="fl">Observações</label><textarea id="it-obs" ${dis("observacoes")}>${esc(v.observacoes || "")}</textarea></div>
    ${it ? `<div class="go-box"><div class="go-box-t">Anexos</div><div id="it-anexos">${htmlAnexos(it.anexos, { remover: reg, onRemover: "GO.removerAnexoItem" })}</div>
      ${reg ? `<label class="btn-sm go-upload">📎 Anexar arquivo<input type="file" multiple onchange="GO.anexarItem('${esc(id)}',this)"></label>` : ""}</div>
      <details class="go-box" ontoggle="GO.carregarHistItem(this,'${esc(id)}')"><summary class="go-box-t">Histórico de alterações${tipo === "tarefa" && modo === "producao" ? " e lançamentos" : ""}</summary><div class="go-muted">Carregando…</div></details>` : ""}
    <div class="go-erro" id="it-err"></div>`;
  const podeSalvar = ger || reg;
  abrirModal({
    titulo: it ? `${TIPOS[tipo]} ${esc(it.codigo)}` : `Nova ${TIPOS[tipo].toLowerCase()}`,
    corpo,
    rodape: `${it && ger ? `<button class="btn-danger" style="margin-right:auto" onclick="GO.excluirItem('${esc(id)}')">Excluir</button>` : ""}
      <button class="btn-cancel" onclick="GO.fecharModal()">${podeSalvar ? "Cancelar" : "Fechar"}</button>
      ${podeSalvar ? `<button class="btn-save" id="it-salvar">Salvar</button>` : ""}`
  });
  if (!it) $("it-tipo").onchange = e => { fecharModal(); abrirItem(null, e.target.value, ""); };
  if (!it && tipo !== "etapa") $("it-pai").onchange = e => { $("it-cod").value = sugerirCodigo(e.target.value || null); };
  document.querySelectorAll('input[name="it-modo"]').forEach(r => r.onchange = () => {
    const m = document.querySelector('input[name="it-modo"]:checked').value;
    $("it-manual").style.display = m === "manual" ? "" : "none"; $("it-prod").style.display = m === "producao" ? "" : "none";
  });
  if (podeSalvar) $("it-salvar").onclick = () => salvarItem(id, tipo);
}
function resumoCons(c) {
  if (!c.total) return `<div class="go-muted">Sem tarefas.</div>`;
  return `<div class="go-resumo"><div><b>${c.total}</b> tarefas</div><div><b>${c.concluida}</b> concluídas (${fmtPct(c.pctConcluidas)})</div><div><b>${c.atrasadas}</b> atrasadas</div>
    <div>${c.ponderado != null ? `Avanço físico ponderado: <b>${fmtPct(c.ponderado)}</b>` : `Média simples do avanço: <b>${fmtPct(c.mediaSimples)}</b> <span class="go-muted">(${c.semPeso} tarefa(s) sem peso)</span>`}</div></div>`;
}

function lerFormItem(tipo) {
  const g = id => $(id) ? $(id).value : undefined;
  const d = {
    tipo, paiId: g("it-pai") || null, codigo: normCodigo(g("it-cod")), descricao: (g("it-desc") || "").trim(),
    responsavelUid: g("it-resp") || null, inicioPrev: g("it-ini") || null, fimPrev: g("it-fim") || null, observacoes: (g("it-obs") || "").trim()
  };
  d.responsavelNome = d.responsavelUid ? nomeUsuario(d.responsavelUid) : "";
  const erros = [];
  const ec = validaCodigo(d.codigo); if (ec) erros.push(ec);
  if (!d.descricao) erros.push("Informe a descrição.");
  if (tipo !== "etapa" && !d.paiId) erros.push(`Escolha a ${tipo === "subetapa" ? "etapa" : "etapa ou subetapa"} vinculada.`);
  if (d.inicioPrev && d.fimPrev && d.fimPrev < d.inicioPrev) erros.push("Término previsto anterior ao início previsto.");
  if (tipo === "tarefa") {
    d.inicioReal = g("it-inir") || null; d.fimReal = g("it-fimr") || null;
    if (d.inicioReal && d.fimReal && d.fimReal < d.inicioReal) erros.push("Término real anterior ao início real.");
    d.modo = document.querySelector('input[name="it-modo"]:checked').value;
    const pct = parseNum(g("it-pct")); d.percentual = pct == null ? 0 : pct;
    if (isNaN(d.percentual) || d.percentual < 0 || d.percentual > 100) erros.push("Percentual deve estar entre 0 e 100.");
    const q = parseNum(g("it-qprev")); d.qtdPrevista = q;
    d.unidade = (g("it-un") || "").trim();
    if (d.modo === "producao") {
      if (q == null || isNaN(q) || q < 0) erros.push("Informe a quantidade prevista (zero ou mais).");
      if (!d.unidade) erros.push("Informe a unidade de medida.");
    } else if (q != null && (isNaN(q) || q < 0)) erros.push("Quantidade prevista inválida.");
    const p = parseNum(g("it-peso")); d.peso = p;
    if (p != null && (isNaN(p) || p < 0)) erros.push("Peso inválido.");
  }
  return { d, erros };
}
/** Preenche início/término reais automaticamente quando o avanço muda. */
function autoDatasReais(atual, novo, dataRef) {
  const ant = calcAvanco({ ...atual, tipo: "tarefa" }), dep = calcAvanco({ ...atual, ...novo, tipo: "tarefa" });
  const a = ant?.avanco ?? 0, b = dep?.avanco ?? (dep?.exe > 0 ? 1 : 0), patch = {};
  const ir = "inicioReal" in novo ? novo.inicioReal : atual.inicioReal, fr = "fimReal" in novo ? novo.fimReal : atual.fimReal;
  if (b > 0 && !ir) patch.inicioReal = dataRef;
  if (b >= 100 && !fr) patch.fimReal = dataRef;
  if (b < 100 && a >= 100 && fr && !("fimReal" in novo && novo.fimReal !== atual.fimReal)) patch.fimReal = null;
  return patch;
}

async function salvarItem(id, tipo) {
  const P = pode(), errBox = $("it-err");
  const { d, erros } = lerFormItem(tipo);
  if (erros.length) { errBox.innerHTML = erros.map(esc).join("<br>"); return; }
  const key = codigoKey(d.codigo);
  const dup = S.itens.find(i => i.codigoKey === key && i.id !== id);
  if (dup) { errBox.textContent = `O código ${d.codigo} já é usado por “${dup.descricao}”.`; return; }
  const pai = d.paiId ? S.itens.find(i => i.id === d.paiId) : null;
  if (tipo === "subetapa" && (!pai || pai.tipo !== "etapa")) { errBox.textContent = "Subetapa precisa estar dentro de uma etapa."; return; }
  if (tipo === "tarefa" && (!pai || pai.tipo === "tarefa")) { errBox.textContent = "Tarefa precisa estar dentro de uma etapa ou subetapa."; return; }
  $("it-salvar").disabled = true;
  try {
    const obraId = S.obraId;
    if (!id) {
      const ref = doc(COL("itens"));
      const irmaos = S.itens.filter(i => (i.paiId || null) === (d.paiId || null));
      const ordem = (irmaos.reduce((m, i) => Math.max(m, i.ordem || 0), 0)) + 1000;
      await runTransaction(C.db, async tx => {
        const cod = await tx.get(R("codigos", key));
        if (cod.exists()) throw new Error(`O código ${d.codigo} já existe nesta obra.`);
        tx.set(R("codigos", key), { itemId: ref.id, codigo: d.codigo });
        const dados = { ...d, codigoKey: key, ordem, anexos: [], qtdExecutada: 0, criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp(), atualizadoPor: uid() };
        if (tipo === "tarefa") Object.assign(dados, autoDatasReais({ qtdExecutada: 0, percentual: 0 }, d, hojeISO()));
        tx.set(ref, dados);
        regHist(tx, obraId, { alvoTipo: "item", alvoId: ref.id, acao: "criar", resumo: `${TIPOS[tipo]} ${d.codigo} criada` });
        if (tipo === "tarefa") { const av = calcAvanco({ ...dados }).avanco ?? 0; regSerie(tx, obraId, ref.id, 0, av, hojeISO()); }
      });
      toast(`${TIPOS[tipo]} criada`);
      if (d.paiId) S.ui.abertos.add(d.paiId);
      fecharModal();
      return;
    }
    await runTransaction(C.db, async tx => {
      const snap = await tx.get(R("itens", id));
      if (!snap.exists()) throw new Error("O item foi excluído por outra pessoa.");
      const atual = snap.data();
      let novo = { ...d };
      if (!P.gerir) novo = Object.fromEntries(CAMPOS_COLAB.filter(k => k in d).map(k => [k, d[k]]));
      if (tipo === "tarefa") Object.assign(novo, autoDatasReais(atual, novo, hojeISO()));
      const patch = {};
      for (const [k, v] of Object.entries(novo)) if (JSON.stringify(atual[k] ?? null) !== JSON.stringify(v ?? null)) patch[k] = v ?? null;
      if (!Object.keys(patch).length) return;
      if (P.gerir && atual.codigoKey !== key) {
        const cod = await tx.get(R("codigos", key));
        if (cod.exists()) throw new Error(`O código ${d.codigo} já existe nesta obra.`);
        tx.set(R("codigos", key), { itemId: id, codigo: d.codigo });
        tx.delete(R("codigos", atual.codigoKey));
        patch.codigoKey = key;
      }
      tx.update(R("itens", id), { ...patch, atualizadoEm: serverTimestamp(), atualizadoPor: uid() });
      const campos = diffCampos(atual, patch, Object.keys(patch).filter(k => k !== "codigoKey" && k !== "responsavelNome"));
      if (tipo === "tarefa") {
        const a = calcAvanco(atual)?.avanco ?? 0, b = calcAvanco({ ...atual, ...patch })?.avanco ?? 0;
        if (a !== b) { if (campos) campos.avanco = { antes: a, depois: b }; regSerie(tx, obraId, id, a, b, hojeISO()); }
      }
      regHist(tx, obraId, { alvoTipo: "item", alvoId: id, acao: "editar", resumo: `${TIPOS[tipo]} ${d.codigo || atual.codigo} alterada`, campos });
    });
    toast("Alterações salvas"); fecharModal();
  } catch (e) {
    console.error(e); errBox.textContent = erroMsg(e); $("it-salvar").disabled = false;
  }
}

async function excluirItem(id) {
  const M = S.modelo, d = M.der.get(id);
  const ids = []; const marca = x => { ids.push(x); (M.der.get(x)?.filhos || []).forEach(marca); }; marca(id);
  const itens = ids.map(x => S.itens.find(i => i.id === x)).filter(Boolean);
  const comProd = itens.filter(i => i.ultimoLanc);
  if (comProd.length) {
    await perguntar("Não é possível excluir", `${comProd.length === 1 ? "A tarefa" : "As tarefas"} ${esc(comProd.slice(0, 5).map(i => i.codigo).join(", "))}${comProd.length > 5 ? "…" : ""} já ${comProd.length === 1 ? "tem" : "têm"} produção lançada. Para preservar o histórico, a exclusão é bloqueada. Você pode ajustar a descrição ou o vínculo.`, { botao: "Entendi" });
    return;
  }
  const ok = await perguntar("Excluir item", `Excluir <b>${esc(d.codigo)} · ${esc(d.descricao)}</b>${ids.length > 1 ? ` e os <b>${ids.length - 1}</b> itens dentro dele` : ""}? Os RDOs que citam essas tarefas continuam guardados.`, { botao: "Excluir", perigo: true });
  if (ok === null) return;
  try {
    for (let i = 0; i < itens.length; i += 100) {
      const b = writeBatch(C.db);
      for (const it of itens.slice(i, i + 100)) {
        b.delete(R("itens", it.id));
        if (it.codigoKey) b.delete(R("codigos", it.codigoKey));
        regHist(b, S.obraId, { alvoTipo: "item", alvoId: it.id, acao: "excluir", resumo: `${TIPOS[it.tipo]} ${it.codigo} excluída`, campos: { descricao: { antes: it.descricao, depois: null } } });
      }
      await b.commit();
    }
    fecharModal(); toast(`${itens.length} item(ns) excluído(s)`);
  } catch (e) { toast(erroMsg(e)); }
}

async function mover(id, dir) {
  const it = S.itens.find(i => i.id === id); if (!it) return;
  const irmaos = S.itens.filter(i => (i.paiId || null) === (it.paiId || null)).sort((a, b) => (a.ordem ?? 0) - (b.ordem ?? 0) || cmpCodigo(a.codigo, b.codigo));
  const k = irmaos.findIndex(i => i.id === id), j = k + dir;
  if (j < 0 || j >= irmaos.length) return;
  [irmaos[k], irmaos[j]] = [irmaos[j], irmaos[k]];
  try {
    const b = writeBatch(C.db); let n = 0;
    irmaos.forEach((x, i) => { const o = (i + 1) * 1000; if (x.ordem !== o) { b.update(R("itens", x.id), { ordem: o }); n++; } });
    if (n) await b.commit();
  } catch (e) { toast(erroMsg(e)); }
}

async function ajustarProducao(id) {
  const it = S.itens.find(i => i.id === id); if (!it) return;
  const box = $("go-pergunta-box");
  box.innerHTML = `<div class="mh"><div class="mt">Lançar ou ajustar produção · ${esc(it.codigo)}</div></div>
    <div class="fg"><div class="go-dica">Acumulado atual: <b>${fmtNum(it.qtdExecutada || 0)} ${esc(it.unidade || "")}</b> de ${fmtNum(it.qtdPrevista || 0)}. Use valor negativo para corrigir para menos. O ajuste fica registrado com o motivo; a produção diária normal deve vir pelo RDO.</div>
      <div class="go-grid2"><div class="fgf"><label class="fl">Quantidade (${esc(it.unidade || "")}) *</label><input id="aj-q" inputmode="decimal" placeholder="ex.: 12,5 ou -3"></div>
      <div class="fgf"><label class="fl">Data de referência</label><input type="date" id="aj-d" value="${hojeISO()}"></div></div>
      <div class="fgf"><label class="fl">Motivo *</label><textarea id="aj-m"></textarea></div><div class="go-erro" id="aj-err"></div></div>
    <div class="mf"><button class="btn-cancel" id="aj-n">Cancelar</button><button class="btn-save" id="aj-s">Lançar</button></div>`;
  $("go-pergunta").classList.add("open");
  $("aj-n").onclick = () => $("go-pergunta").classList.remove("open");
  $("aj-s").onclick = async () => {
    const q = parseNum($("aj-q").value), motivo = $("aj-m").value.trim(), data = $("aj-d").value || hojeISO();
    if (q == null || isNaN(q) || q === 0) { $("aj-err").textContent = "Informe uma quantidade diferente de zero."; return; }
    if (motivo.length < 3) { $("aj-err").textContent = "Descreva o motivo."; return; }
    $("aj-s").disabled = true;
    try {
      await runTransaction(C.db, async tx => {
        const s = await tx.get(R("itens", id));
        const atual = s.data(), antes = Number(atual.qtdExecutada) || 0, depois = arred(antes + q);
        if (depois < 0) throw new Error(`O acumulado ficaria negativo (${fmtNum(depois)}).`);
        const lref = doc(COL("lancamentos"));
        tx.set(lref, { tipo: "ajuste", itemId: id, codigo: atual.codigo, qtd: q, antes, depois, unidade: atual.unidade || "", data, motivo, autorUid: uid(), autorNome: meuNome(), em: serverTimestamp() });
        const patch = { qtdExecutada: depois, ultimoLanc: lref.id, ...autoDatasReais(atual, { qtdExecutada: depois }, data) };
        tx.update(R("itens", id), { ...patch, atualizadoEm: serverTimestamp(), atualizadoPor: uid() });
        const a = calcAvanco(atual)?.avanco ?? 0, b = calcAvanco({ ...atual, ...patch })?.avanco ?? 0;
        regHist(tx, S.obraId, { alvoTipo: "item", alvoId: id, acao: "producao", resumo: `Ajuste de produção: ${q > 0 ? "+" : ""}${fmtNum(q)} ${atual.unidade || ""} — ${motivo}`, campos: { qtdExecutada: { antes, depois }, ...(a !== b ? { avanco: { antes: a, depois: b } } : {}) } });
        regSerie(tx, S.obraId, id, a, b, data);
      });
      $("go-pergunta").classList.remove("open"); fecharModal(); toast("Produção lançada");
    } catch (e) { $("aj-err").textContent = erroMsg(e); $("aj-s").disabled = false; }
  };
}

async function anexarItem(id, input) {
  const files = [...input.files]; input.value = "";
  if (!files.length) return;
  toast("Enviando…");
  try {
    const metas = [];
    for (const f of files) metas.push(await salvarArquivo(f, { tipo: "item", id }));
    const it = S.itens.find(i => i.id === id);
    const anexos = [...(it.anexos || []), ...metas.map(({ thumb, ...m }) => ({ ...m, thumb: thumb || null }))];
    const b = writeBatch(C.db);
    b.update(R("itens", id), { anexos, atualizadoEm: serverTimestamp(), atualizadoPor: uid() });
    regHist(b, S.obraId, { alvoTipo: "item", alvoId: id, acao: "anexo", resumo: `Anexo incluído: ${metas.map(m => m.nome).join(", ")}` });
    await b.commit();
    $("it-anexos").innerHTML = htmlAnexos(anexos, { remover: true, onRemover: "GO.removerAnexoItem" });
    $("it-anexos").dataset.item = id;
    toast("Anexo salvo");
  } catch (e) { toast(erroMsg(e)); }
}
async function removerAnexoItem(arqId) {
  const it = S.itens.find(i => (i.anexos || []).some(a => a.id === arqId)); if (!it) return;
  const a = it.anexos.find(x => x.id === arqId);
  if ((await perguntar("Remover anexo", `Remover “${esc(a.nome)}” desta tarefa?`, { botao: "Remover", perigo: true })) === null) return;
  try {
    const anexos = it.anexos.filter(x => x.id !== arqId);
    const b = writeBatch(C.db);
    b.update(R("itens", it.id), { anexos, atualizadoEm: serverTimestamp(), atualizadoPor: uid() });
    regHist(b, S.obraId, { alvoTipo: "item", alvoId: it.id, acao: "anexo", resumo: `Anexo removido: ${a.nome}` });
    await b.commit();
    $("it-anexos").innerHTML = htmlAnexos(anexos, { remover: true, onRemover: "GO.removerAnexoItem" });
  } catch (e) { toast(erroMsg(e)); }
}
async function carregarHistItem(det, id) {
  if (!det.open || det.dataset.ok) return;
  det.dataset.ok = 1;
  const alvo = det.querySelector(".go-muted");
  try {
    const [hist, lanc] = await Promise.all([carregarHistorico(S.obraId, id), getDocs(query(COL("lancamentos"), where("itemId", "==", id)))]);
    const ls = lanc.docs.map(d => d.data()).sort((a, b) => String(tsData(b.em)?.toISOString() || "").localeCompare(String(tsData(a.em)?.toISOString() || "")));
    alvo.outerHTML = (ls.length ? `<div class="go-sub">Lançamentos de produção</div><div class="go-lanc">${ls.map(l => `<div><span>${fmtData(l.data)}</span><span class="go-chip c-${l.tipo}">${{ rdo: "RDO", estorno: "Estorno", ajuste: "Ajuste" }[l.tipo]}</span>
      <b>${l.qtd > 0 ? "+" : ""}${fmtNum(l.qtd)} ${esc(l.unidade || "")}</b><span class="go-muted">${fmtNum(l.antes)} → ${fmtNum(l.depois)}${l.rdoId ? ` · RDO ${fmtData(l.rdoId)} v${l.versao}` : ""}${l.motivo ? " · " + esc(l.motivo) : ""} · ${esc(l.autorNome || "")}</span></div>`).join("")}</div>` : "")
      + `<div class="go-sub">Alterações</div>` + htmlHistorico(hist);
  } catch (e) { alvo.textContent = erroMsg(e); }
}

// ---------------------------------------------------------------------
// RDO — Relatório Diário de Obra
// ---------------------------------------------------------------------
const chipRdo = s => `<span class="go-chip r-${s}">${STATUS_RDO[s] || s}</span>`;
function renderRdoLista() {
  S.rdoEd = null;
  const P = pode();
  const filtro = S.ui.rdoFiltro || "";
  const lista = S.rdos.filter(r => !filtro || r.status === filtro);
  const pend = S.rdos.filter(r => r.status === "enviado").length;
  $("obra-corpo").innerHTML = `
    <div class="go-toolbar">
      ${P.registrar ? `<div class="go-novo-rdo"><label class="fl">Data do relatório</label><input type="date" id="rdo-nova-data" value="${hojeISO()}" max="${hojeISO()}"><button class="btn-new" onclick="GO.novoRdo()">+ Novo RDO</button></div>` : ""}
      <span class="go-flex"></span>
      <select id="rdo-filtro"><option value="">Todas as situações</option>${Object.entries(STATUS_RDO).map(([k, v]) => `<option value="${k}" ${filtro === k ? "selected" : ""}>${v}</option>`).join("")}</select>
    </div>
    ${pend && P.gerir ? `<div class="migr-banner"><div><b>${pend} relatório(s) aguardando sua aprovação</b><span>Ao aprovar, a produção informada entra nas tarefas vinculadas.</span></div><button class="btn-migr" onclick="document.getElementById('rdo-filtro').value='enviado';GO.filtroRdo('enviado')">Ver pendentes</button></div>` : ""}
    <div class="go-rdo-lista">${lista.length ? lista.map(r => {
      const mo = (r.maoDeObra || []).reduce((s, m) => s + (Number(m.qtd) || 0), 0);
      const nat = (r.atividades || []).length;
      return `<a class="go-rdo-card" href="#acompanhamento/${encodeURIComponent(S.obraId)}/rdo/${r.id}">
        <div class="go-rdo-data">${fmtData(r.data)}<span>${["dom", "seg", "ter", "qua", "qui", "sex", "sáb"][new Date(r.data + "T12:00").getDay()]}</span></div>
        <div class="go-rdo-info"><div>${chipRdo(r.status)}${r.versao > 1 ? ` <span class="go-muted">corrigido (v${r.versao})</span>` : ""}</div>
          <div class="go-muted">${nat} atividade(s) · ${mo} pessoa(s) · ${(r.fotos || []).length} foto(s) · clima: ${esc([r.clima?.manha, r.clima?.tarde].filter(Boolean).join(" / ") || "—")}</div>
          <div class="go-muted">Autor: ${esc(r.autorNome || "—")}${r.aprovadorNome && r.status === "aprovado" ? ` · Aprovado por ${esc(r.aprovadorNome)}` : ""}</div></div>
        <div class="obra-go">Abrir →</div></a>`;
    }).join("") : `<div class="obra-vazio"><div class="obra-vazio-ic">📝</div><h3>Nenhum relatório${filtro ? " nesta situação" : ""}</h3><p>${P.registrar ? "Crie o RDO do dia com o botão acima." : ""}</p></div>`}</div>`;
  $("rdo-filtro").onchange = e => filtroRdo(e.target.value);
}
function filtroRdo(v) { S.ui.rdoFiltro = v; renderRdoLista(); }

async function novoRdo() {
  const data = $("rdo-nova-data").value;
  if (!data) { toast("Escolha a data"); return; }
  if (data > hojeISO()) { toast("Não é possível criar RDO de data futura"); return; }
  const destino = `#acompanhamento/${encodeURIComponent(S.obraId)}/rdo/${data}`;
  if (S.rdos.some(r => r.id === data)) { location.hash = destino; return; }
  const anterior = S.rdos.filter(r => r.data < data && r.status !== "cancelado")[0];
  try {
    await runTransaction(C.db, async tx => {
      const s = await tx.get(R("rdos", data));
      if (s.exists()) return;
      tx.set(R("rdos", data), {
        data, status: "rascunho", versao: 0,
        clima: { manha: "Bom", tarde: "Bom", noite: "", condicao: "praticavel", impacto: "", horasParadas: 0 },
        maoDeObra: anterior ? (anterior.maoDeObra || []).map(m => ({ ...m })) : [],
        equipamentos: anterior ? (anterior.equipamentos || []).map(m => ({ ...m })) : [],
        atividades: [], ocorrencias: "", impedimentos: "", observacoes: "", fotos: [], lancamentos: [],
        autorUid: uid(), autorNome: meuNome(), criadoEm: serverTimestamp(), atualizadoEm: serverTimestamp(),
        eventos: [{ acao: "criar", porUid: uid(), porNome: meuNome(), emISO: new Date().toISOString() }]
      });
      regHist(tx, S.obraId, { alvoTipo: "rdo", alvoId: data, acao: "criar", resumo: `RDO de ${fmtData(data)} criado` });
    });
    location.hash = destino;
  } catch (e) { toast(erroMsg(e)); }
}

function abrirRdoEditor(id) {
  const r = S.rdos.find(x => x.id === id);
  if (!r) { $("obra-corpo").innerHTML = `<div class="obra-vazio"><h3>Relatório de ${fmtData(id)} não encontrado</h3><p><a href="#acompanhamento/${encodeURIComponent(S.obraId)}/rdo">Voltar à lista</a></p></div>`; return; }
  if (!S.rdoEd || S.rdoEd.id !== id || !S.rdoEd.dirty) S.rdoEd = { id, dirty: false, v: copiaSemTs(r) };
  desenharRdoEditor();
}
/** Cópia profunda trocando Timestamps do Firestore por texto ISO (para edição local). */
function copiaSemTs(v) {
  if (v && typeof v.toDate === "function") return v.toDate().toISOString();
  if (Array.isArray(v)) return v.map(copiaSemTs);
  if (v && typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, copiaSemTs(x)]));
  return v;
}
const rdoEditavel = r => pode().registrar && ["rascunho", "enviado"].includes(r.status);

function desenharRdoEditor() {
  const ed = S.rdoEd; if (!ed) return;
  const r = ed.v, P = pode(), E = rdoEditavel(r), dis = E ? "" : "disabled";
  const M = S.modelo;
  const totMo = (r.maoDeObra || []).reduce((s, m) => s + (Number(m.qtd) || 0), 0);
  const opt = (lista, sel) => lista.map(x => `<option ${x === sel ? "selected" : ""}>${esc(x)}</option>`).join("");
  const ativ = (r.atividades || []).map((a, i) => {
    const it = a.itemId ? S.itens.find(x => x.id === a.itemId) : null;
    const prod = it && (it.modo || "manual") === "producao";
    let info = "";
    if (a.itemId && !it) info = `<span class="go-alerta">Tarefa excluída da EAP</span>`;
    else if (prod && r.status === "cancelado") info = "";
    else if (prod) {
      const q = parseNum(a.qtd), prev = Number(it.qtdPrevista) || 0, exe = Number(it.qtdExecutada) || 0;
      const apos = r.status === "aprovado" ? exe : arred(exe + (q > 0 ? q : 0));
      info = `Acumulado ${r.status === "aprovado" ? "" : "aprovado "}${fmtNum(exe)} / ${fmtNum(prev)} ${esc(it.unidade || "")}${r.status !== "aprovado" && q > 0 ? ` → após aprovação ${fmtNum(apos)}` : ""}${prev > 0 && apos > prev + 1e-9 ? ` <span class="go-alerta">▲ acima do previsto (${fmtPct(apos / prev * 100)})</span>` : ""}${prev <= 0 ? ` <span class="go-alerta">⚠️ sem quantidade prevista</span>` : ""}`;
    } else if (it) info = `Avanço atual ${fmtPct(Number(it.percentual) || 0)}`;
    return `<div class="go-ativ">
      <div class="go-ativ-t">${it || a.codigo ? `<span><b>${esc(it ? it.codigo : a.codigo)}</b> ${esc(it ? it.descricao : a.descricao)}</span>` : `<input data-k="atividades.${i}.descricao" value="${esc(a.descricao || "")}" placeholder="Descrição da atividade (sem tarefa vinculada)" ${dis}>`}
        ${E ? `<button class="card-btn del" onclick="GO.rdoRemover('atividades',${i})" title="Remover">🗑</button>` : ""}</div>
      <div class="go-ativ-c">
        ${prod ? `<div class="fgf"><label class="fl">Quantidade executada no dia (${esc(it.unidade || "")})</label><input data-k="atividades.${i}.qtd" inputmode="decimal" value="${esc(a.qtd ?? "")}" ${dis}></div>`
          : it ? `<div class="fgf"><label class="fl">Avanço acumulado após o dia (%)</label><input data-k="atividades.${i}.percentual" inputmode="decimal" value="${esc(a.percentual ?? "")}" placeholder="${esc(String(it.percentual ?? 0))}" ${dis}></div>` : ""}
        <div class="fgf go-flex"><label class="fl">Observação</label><input data-k="atividades.${i}.obs" value="${esc(a.obs || "")}" ${dis}></div>
      </div>
      ${info ? `<div class="go-ativ-i">${info}</div>` : ""}</div>`;
  }).join("");
  const eventos = (r.eventos || []).slice().reverse();
  const base = `#acompanhamento/${encodeURIComponent(S.obraId)}/rdo`;
  $("obra-corpo").innerHTML = `
  <div class="go-rdo-head">
    <a class="obra-voltar" href="${base}">‹ Relatórios</a>
    <div class="go-rdo-titulo">RDO de ${fmtData(r.data)} ${chipRdo(r.status)} ${r.versao ? `<span class="go-muted">versão aprovada ${r.versao}</span>` : ""}</div>
    <span class="go-flex"></span>
    <button class="btn-export" onclick="GO.pdfRdo()">⬇️ PDF</button>
  </div>
  ${r.status === "aprovado" ? `<div class="go-ok-box">Aprovado por ${esc(r.aprovadorNome || "—")} em ${fmtDataHora(r.aprovadoEm)}. A produção deste relatório já está nas tarefas. ${P.gerir ? "Para corrigir, use “Reabrir para correção” — os lançamentos são estornados e o histórico fica guardado." : ""}</div>` : ""}
  ${r.status === "cancelado" ? `<div class="go-alerta-box">Relatório cancelado. Nada dele conta nas tarefas.</div>` : ""}
  <div class="go-cols">
    <section class="go-sec"><h3>☁️ Condições climáticas</h3>
      <div class="go-grid3">
        <div class="fgf"><label class="fl">Manhã</label><select data-k="clima.manha" ${dis}>${opt(["", ...CLIMA], r.clima?.manha)}</select></div>
        <div class="fgf"><label class="fl">Tarde</label><select data-k="clima.tarde" ${dis}>${opt(["", ...CLIMA], r.clima?.tarde)}</select></div>
        <div class="fgf"><label class="fl">Noite</label><select data-k="clima.noite" ${dis}>${opt(["", ...CLIMA], r.clima?.noite)}</select></div>
        <div class="fgf"><label class="fl">Condição de trabalho</label><select data-k="clima.condicao" ${dis}>${Object.entries(CONDICAO).map(([k, v]) => `<option value="${k}" ${r.clima?.condicao === k ? "selected" : ""}>${v}</option>`).join("")}</select></div>
        <div class="fgf"><label class="fl">Horas paralisadas</label><input data-k="clima.horasParadas" inputmode="decimal" value="${esc(r.clima?.horasParadas ?? "")}" ${dis}></div>
      </div>
      <div class="fgf"><label class="fl">Impacto na execução</label><input data-k="clima.impacto" value="${esc(r.clima?.impacto || "")}" placeholder="ex.: chuva às 14h suspendeu a concretagem" ${dis}></div>
    </section>
    <section class="go-sec"><h3>👷 Mão de obra <span class="go-muted">(${totMo} pessoas)</span></h3>
      ${(r.maoDeObra || []).map((m, i) => `<div class="go-linha3"><input data-k="maoDeObra.${i}.funcao" list="go-funcoes" value="${esc(m.funcao || "")}" placeholder="Função" ${dis}>
        <input data-k="maoDeObra.${i}.empresa" value="${esc(m.empresa || "")}" placeholder="Própria / empreiteira" ${dis}>
        <input data-k="maoDeObra.${i}.qtd" inputmode="numeric" value="${esc(m.qtd ?? "")}" placeholder="Qtd" class="go-qtd" ${dis}>
        ${E ? `<button class="card-btn del" onclick="GO.rdoRemover('maoDeObra',${i})">🗑</button>` : ""}</div>`).join("") || '<div class="go-muted">Nenhum registro.</div>'}
      <datalist id="go-funcoes">${FUNCOES.map(f => `<option value="${f}">`).join("")}</datalist>
      ${E ? `<button class="btn-sm" onclick="GO.rdoIncluir('maoDeObra')">+ Função</button>` : ""}
      <h3 style="margin-top:14px">🚜 Equipamentos</h3>
      ${(r.equipamentos || []).map((m, i) => `<div class="go-linha3"><input data-k="equipamentos.${i}.nome" value="${esc(m.nome || "")}" placeholder="Equipamento" ${dis}>
        <select data-k="equipamentos.${i}.situacao" ${dis}>${opt(["Operando", "Parado", "Em manutenção"], m.situacao)}</select>
        <input data-k="equipamentos.${i}.qtd" inputmode="numeric" value="${esc(m.qtd ?? "")}" placeholder="Qtd" class="go-qtd" ${dis}>
        ${E ? `<button class="card-btn del" onclick="GO.rdoRemover('equipamentos',${i})">🗑</button>` : ""}</div>`).join("") || '<div class="go-muted">Nenhum registro.</div>'}
      ${E ? `<button class="btn-sm" onclick="GO.rdoIncluir('equipamentos')">+ Equipamento</button>` : ""}
    </section>
  </div>
  <section class="go-sec"><h3>🔨 Atividades realizadas</h3>
    ${ativ || '<div class="go-muted">Nenhuma atividade registrada.</div>'}
    ${E ? `<div class="go-picker"><input id="rdo-busca-tarefa" placeholder="🔎 Vincular tarefa: digite o código ou a descrição" autocomplete="off"><div id="rdo-sug" class="go-sug"></div></div>
      <button class="btn-sm" onclick="GO.rdoIncluir('atividades')">+ Atividade sem tarefa vinculada</button>` : ""}
  </section>
  <div class="go-cols">
    <section class="go-sec"><h3>⚠️ Ocorrências e impedimentos</h3>
      <div class="fgf"><label class="fl">Ocorrências</label><textarea data-k="ocorrencias" ${dis}>${esc(r.ocorrencias || "")}</textarea></div>
      <div class="fgf"><label class="fl">Impedimentos</label><textarea data-k="impedimentos" ${dis}>${esc(r.impedimentos || "")}</textarea></div>
      <div class="fgf"><label class="fl">Observações gerais</label><textarea data-k="observacoes" ${dis}>${esc(r.observacoes || "")}</textarea></div>
    </section>
    <section class="go-sec"><h3>📷 Fotos e anexos <span class="go-muted">(${(r.fotos || []).length})</span></h3>
      <div class="go-fotos">${(r.fotos || []).map((f, i) => `<div class="go-foto">${f.thumb ? `<img src="${f.thumb}" alt="" onclick="GO.baixarArquivo('${esc(f.id)}')">` : `<div class="go-foto-arq" onclick="GO.baixarArquivo('${esc(f.id)}')">📎<br>${esc(f.nome)}</div>`}
        <input data-k="fotos.${i}.legenda" value="${esc(f.legenda || "")}" placeholder="Legenda" ${dis}>${E ? `<button class="card-btn del" onclick="GO.rdoRemoverFoto(${i})">🗑</button>` : ""}</div>`).join("") || '<div class="go-muted">Nenhuma foto.</div>'}</div>
      ${E ? `<label class="btn-sm go-upload">📷 Incluir fotos ou anexos<input type="file" multiple accept="image/*,application/pdf,.doc,.docx,.xls,.xlsx" onchange="GO.rdoFotos(this)"></label><div class="go-dica">Fotos são reduzidas para até 1600 px. Anexos até 700 KB.</div>` : ""}
    </section>
  </div>
  <section class="go-sec"><h3>✍️ Autoria e aprovação</h3>
    <div class="go-resumo"><div>Autor: <b>${esc(r.autorNome || "—")}</b></div><div>Responsável pela aprovação: <b>${esc(r.aprovadorNome || "—")}</b>${r.aprovadoEm && fmtDataHora(r.aprovadoEm) !== "—" ? ` em ${fmtDataHora(r.aprovadoEm)}` : ""}</div></div>
    <div class="go-hist">${eventos.map(e => `<div class="go-hist-i"><div class="go-hist-h"><b>${esc(e.porNome || "—")}</b> · ${fmtDataHora(e.emISO)} · ${esc({ criar: "criou", salvar: "salvou", enviar: "enviou para aprovação", rascunho: "voltou para rascunho", devolver: "devolveu para correção", aprovar: "aprovou", reabrir: "reabriu para correção", cancelar: "cancelou" }[e.acao] || e.acao)}${e.versao ? ` (versão ${e.versao})` : ""}</div>${e.motivo ? `<div class="go-muted">Motivo: ${esc(e.motivo)}</div>` : ""}</div>`).join("")}</div>
  </section>
  <div class="go-erro" id="rdo-err"></div>
  <div class="go-barra-acoes">${acoesRdo(r, P)}</div>`;
  if (E) {
    $("obra-corpo").querySelectorAll("[data-k]").forEach(el => el.addEventListener(el.tagName === "SELECT" ? "change" : "input", () => rdoSet(el.dataset.k, el.value)));
    const busca = $("rdo-busca-tarefa");
    busca.addEventListener("input", () => sugerirTarefas(busca.value));
    busca.addEventListener("focus", () => sugerirTarefas(busca.value));
  }
}
function acoesRdo(r, P) {
  const b = [];
  if (rdoEditavel(r)) {
    if (r.status === "rascunho" && !(r.versao > 0) && (P.gerir || r.autorUid === uid())) b.push(`<button class="btn-cancel" style="margin-right:auto" onclick="GO.excluirRdo()">Excluir rascunho</button>`);
    if (P.gerir && r.status !== "cancelado") b.push(`<button class="btn-cancel" onclick="GO.cancelarRdo()">Cancelar RDO</button>`);
    b.push(`<button class="btn-cancel" onclick="GO.salvarRdo()">Salvar</button>`);
    if (r.status === "rascunho") b.push(`<button class="btn-save" onclick="GO.salvarRdo('enviado')">Enviar para aprovação</button>`);
    if (r.status === "enviado") {
      b.push(`<button class="btn-cancel" onclick="GO.salvarRdo('rascunho')">Voltar para rascunho</button>`);
      if (P.gerir) b.push(`<button class="btn-save" onclick="GO.aprovarRdo()">✔ Aprovar e lançar produção</button>`);
    }
  } else if (r.status === "aprovado" && P.gerir) {
    b.push(`<button class="btn-cancel" onclick="GO.cancelarRdo()">Cancelar RDO</button>`);
    b.push(`<button class="btn-save" onclick="GO.reabrirRdo()">Reabrir para correção</button>`);
  }
  return b.join("");
}
function rdoSet(path, valor) {
  const ks = path.split("."); let o = S.rdoEd.v;
  for (let i = 0; i < ks.length - 1; i++) { o[ks[i]] = o[ks[i]] ?? (/^\d+$/.test(ks[i + 1]) ? [] : {}); o = o[ks[i]]; }
  o[ks[ks.length - 1]] = valor;
  S.rdoEd.dirty = true;
}
function rdoIncluir(lista) {
  const v = S.rdoEd.v;
  v[lista] = v[lista] || [];
  v[lista].push(lista === "maoDeObra" ? { funcao: "", empresa: "Própria", qtd: "" } : lista === "equipamentos" ? { nome: "", situacao: "Operando", qtd: 1 } : { itemId: null, descricao: "", obs: "" });
  S.rdoEd.dirty = true; desenharRdoEditor();
}
function rdoRemover(lista, i) { S.rdoEd.v[lista].splice(i, 1); S.rdoEd.dirty = true; desenharRdoEditor(); }
function sugerirTarefas(txt) {
  const box = $("rdo-sug"); const b = semAcento(txt.trim());
  const ja = new Set((S.rdoEd.v.atividades || []).map(a => a.itemId).filter(Boolean));
  const lista = S.modelo.tarefas.filter(t => !ja.has(t.id) && (!b || semAcento(t.codigo).includes(b) || semAcento(t.descricao).includes(b)))
    .sort((x, y) => (x.situacao === "concluida") - (y.situacao === "concluida") || (x.ordem ?? 0) - (y.ordem ?? 0)).slice(0, 15);
  box.innerHTML = lista.map(t => `<div class="go-sug-i" onmousedown="GO.rdoVincular('${esc(t.id)}')"><b>${esc(t.codigo)}</b> ${esc(t.descricao)} <span class="go-muted">${t.modo === "producao" ? `${fmtNum(t.qtdExecutada || 0)}/${fmtNum(t.qtdPrevista || 0)} ${esc(t.unidade || "")}` : fmtPct(t.calc.avanco)}</span></div>`).join("")
    || `<div class="go-sug-i go-muted">Nenhuma tarefa encontrada.</div>`;
  box.style.display = "block";
  $("rdo-busca-tarefa").onblur = () => setTimeout(() => box.style.display = "none", 150);
}
function rdoVincular(itemId) {
  const t = S.itens.find(i => i.id === itemId); if (!t) return;
  const v = S.rdoEd.v; v.atividades = v.atividades || [];
  if (v.atividades.filter(a => a.itemId).length >= MAX_ATIV_RDO) { toast(`Limite de ${MAX_ATIV_RDO} tarefas por RDO`); return; }
  v.atividades.push({ itemId, codigo: t.codigo, descricao: t.descricao, unidade: t.unidade || "", qtd: "", percentual: "", obs: "" });
  S.rdoEd.dirty = true; desenharRdoEditor();
  setTimeout(() => { const ins = $("obra-corpo").querySelectorAll(`[data-k^="atividades.${v.atividades.length - 1}."]`); ins[0] && ins[0].focus(); }, 30);
}
async function rdoFotos(input) {
  const files = [...input.files]; input.value = "";
  const v = S.rdoEd.v; v.fotos = v.fotos || [];
  if (v.fotos.length + files.length > 30) { toast("Limite de 30 fotos/anexos por RDO"); return; }
  toast("Enviando…");
  try {
    for (const f of files) v.fotos.push({ ...(await salvarArquivo(f, { tipo: "rdo", id: S.rdoEd.id })), legenda: "" });
    await updateDoc(R("rdos", S.rdoEd.id), { fotos: v.fotos, atualizadoEm: serverTimestamp() });
    desenharRdoEditor(); toast("Arquivos incluídos");
  } catch (e) { toast(erroMsg(e)); desenharRdoEditor(); }
}
async function rdoRemoverFoto(i) {
  const v = S.rdoEd.v, f = v.fotos[i];
  if ((await perguntar("Remover arquivo", `Remover “${esc(f.nome)}” deste RDO?`, { botao: "Remover", perigo: true })) === null) return;
  v.fotos.splice(i, 1);
  try {
    await updateDoc(R("rdos", S.rdoEd.id), { fotos: v.fotos, atualizadoEm: serverTimestamp() });
    const b = writeBatch(C.db); b.delete(R("arquivos", f.id)); b.commit().catch(() => {});
    desenharRdoEditor();
  } catch (e) { toast(erroMsg(e)); }
}

/** Normaliza e valida o conteúdo editável do RDO. */
function normalizarRdo(v) {
  const erros = [];
  const num = (x, nome, { inteiro = false, min = 0, max = Infinity } = {}) => {
    if (x === "" || x == null) return null;
    const n = parseNum(x);
    if (isNaN(n) || n < min || n > max || (inteiro && !Number.isInteger(n))) { erros.push(`${nome}: valor inválido (${x}).`); return null; }
    return n;
  };
  const out = {
    clima: { manha: v.clima?.manha || "", tarde: v.clima?.tarde || "", noite: v.clima?.noite || "", condicao: v.clima?.condicao || "praticavel",
      impacto: (v.clima?.impacto || "").trim(), horasParadas: num(v.clima?.horasParadas, "Horas paralisadas", { max: 24 }) ?? 0 },
    maoDeObra: (v.maoDeObra || []).filter(m => (m.funcao || "").trim() || m.qtd).map((m, i) => ({ funcao: (m.funcao || "").trim(), empresa: (m.empresa || "").trim(), qtd: num(m.qtd, `Mão de obra, linha ${i + 1}`, { inteiro: true }) ?? 0 })),
    equipamentos: (v.equipamentos || []).filter(m => (m.nome || "").trim()).map((m, i) => ({ nome: m.nome.trim(), situacao: m.situacao || "Operando", qtd: num(m.qtd, `Equipamento, linha ${i + 1}`, { inteiro: true }) ?? 1 })),
    atividades: (v.atividades || []).filter(a => a.itemId || (a.descricao || "").trim()).map(a => {
      const rot = a.codigo ? `Atividade ${a.codigo}` : "Atividade";
      return { itemId: a.itemId || null, codigo: a.codigo || "", descricao: (a.descricao || "").trim(), unidade: a.unidade || "",
        qtd: num(a.qtd, `${rot} — quantidade do dia`), percentual: num(a.percentual, `${rot} — avanço`, { max: 100 }), obs: (a.obs || "").trim() };
    }),
    ocorrencias: (v.ocorrencias || "").trim(), impedimentos: (v.impedimentos || "").trim(), observacoes: (v.observacoes || "").trim(),
    fotos: (v.fotos || []).map(f => ({ id: f.id, nome: f.nome, tipo: f.tipo || "", tamanho: f.tamanho || 0, thumb: f.thumb || null, legenda: (f.legenda || "").trim() }))
  };
  if (out.atividades.filter(a => a.itemId).length > MAX_ATIV_RDO) erros.push(`Máximo de ${MAX_ATIV_RDO} tarefas vinculadas por RDO.`);
  return { out, erros };
}
const evento = (acao, extra = {}) => ({ acao, porUid: uid(), porNome: meuNome(), emISO: new Date().toISOString(), ...extra });

async function salvarRdo(novoStatus) {
  const ed = S.rdoEd, r = ed.v, err = $("rdo-err");
  const { out, erros } = normalizarRdo(r);
  if (erros.length) { err.innerHTML = erros.map(esc).join("<br>"); return false; }
  if (novoStatus === "enviado" && !out.atividades.length && !out.ocorrencias && !out.observacoes) { err.textContent = "Inclua ao menos uma atividade, ocorrência ou observação antes de enviar."; return false; }
  const status = novoStatus || r.status;
  try {
    const b = writeBatch(C.db);
    const eventos = [...(r.eventos || [])];
    eventos.push(evento(novoStatus === "enviado" ? "enviar" : novoStatus === "rascunho" ? "rascunho" : "salvar"));
    b.update(R("rdos", ed.id), { ...out, status, eventos, atualizadoEm: serverTimestamp(), atualizadoPor: uid() });
    if (novoStatus && novoStatus !== r.status) regHist(b, S.obraId, { alvoTipo: "rdo", alvoId: ed.id, acao: novoStatus, resumo: `RDO de ${fmtData(r.data)}: ${STATUS_RDO[r.status]} → ${STATUS_RDO[novoStatus]}`, campos: { status: { antes: r.status, depois: novoStatus } } });
    await b.commit();
    ed.dirty = false;
    toast(novoStatus === "enviado" ? "RDO enviado para aprovação" : "RDO salvo");
    return true;
  } catch (e) { err.textContent = erroMsg(e); return false; }
}

/** Aprovação: lança a produção nas tarefas numa única transação, com id fixo por RDO/versão/tarefa. */
export function agregarAtividades(atividades) {
  const ag = new Map();
  for (const a of atividades || []) {
    if (!a.itemId) continue;
    const g = ag.get(a.itemId) || { itemId: a.itemId, qtd: 0, temQtd: false, percentual: null };
    if (a.qtd != null && a.qtd !== "") { g.qtd = arred(g.qtd + Number(a.qtd)); g.temQtd = true; }
    if (a.percentual != null && a.percentual !== "") g.percentual = Math.max(g.percentual ?? 0, Number(a.percentual));
    ag.set(a.itemId, g);
  }
  return [...ag.values()];
}
async function aprovarRdo() {
  const ed = S.rdoEd, err = $("rdo-err");
  if (ed.dirty && !(await salvarRdo())) return;
  const r = S.rdos.find(x => x.id === ed.id) || ed.v;
  const ag = agregarAtividades(r.atividades);
  const faltando = ag.filter(g => !S.itens.some(i => i.id === g.itemId));
  if (faltando.length) { err.textContent = "Há atividades vinculadas a tarefas que foram excluídas. Remova-as antes de aprovar."; return; }
  const acima = ag.filter(g => { const it = S.itens.find(i => i.id === g.itemId); return it.modo === "producao" && it.qtdPrevista > 0 && (it.qtdExecutada || 0) + g.qtd > it.qtdPrevista + 1e-9; });
  const ok = await perguntar("Aprovar RDO", `Aprovar o relatório de <b>${fmtData(r.data)}</b> e lançar a produção em <b>${ag.length}</b> tarefa(s)?${acima.length ? `<br><br>▲ ${acima.length} tarefa(s) ficarão acima da quantidade prevista: ${esc(acima.map(g => S.itens.find(i => i.id === g.itemId).codigo).join(", "))}.` : ""}`, { botao: "Aprovar" });
  if (ok === null) return;
  try {
    await aprovarRdoTx(S.obraId, ed.id);
    toast("RDO aprovado — produção lançada");
  } catch (e) { console.error(e); err.textContent = erroMsg(e); }
}
export async function aprovarRdoTx(obraId, rdoId, ctx = { db: C.db, uid: uid(), nome: meuNome() }) {
  const Rf = (...p) => doc(ctx.db, "obras", obraId, ...p);
  await runTransaction(ctx.db, async tx => {
    const rs = await tx.get(Rf("rdos", rdoId));
    if (!rs.exists()) throw new Error("RDO não encontrado.");
    const r = rs.data();
    if (r.status !== "enviado") throw new Error(`Este RDO não está aguardando aprovação (situação atual: ${STATUS_RDO[r.status] || r.status}). Nada foi lançado.`);
    const versao = (r.versao || 0) + 1;
    const ag = agregarAtividades(r.atividades);
    const itens = {}, lrefs = {};
    for (const g of ag) {
      const s = await tx.get(Rf("itens", g.itemId));
      if (!s.exists()) throw new Error(`A tarefa ${g.itemId} não existe mais.`);
      itens[g.itemId] = s.data();
    }
    for (const g of ag) {
      lrefs[g.itemId] = Rf("lancamentos", `${rdoId}__v${versao}__${g.itemId}`);
      if ((await tx.get(lrefs[g.itemId])).exists()) throw new Error(`A produção da versão ${versao} deste RDO já foi lançada. Nada foi somado de novo.`);
    }
    const ids = [];
    for (const g of ag) {
      const it = itens[g.itemId], prod = (it.modo || "manual") === "producao";
      const antes = Number(it.qtdExecutada) || 0, q = prod && g.temQtd ? g.qtd : 0, depois = arred(antes + q);
      const pA = Number(it.percentual) || 0, pD = !prod && g.percentual != null ? g.percentual : pA;
      const lref = lrefs[g.itemId];
      tx.set(lref, { tipo: "rdo", rdoId, versao, itemId: g.itemId, codigo: it.codigo, modo: it.modo || "manual", qtd: q, antes, depois, unidade: it.unidade || "",
        percentualAntes: pA, percentualDepois: pD, data: r.data, autorUid: ctx.uid, autorNome: ctx.nome, em: serverTimestamp() });
      ids.push(lref.id);
      const patch = { ultimoLanc: lref.id };
      if (q !== 0) patch.qtdExecutada = depois;
      if (pD !== pA) patch.percentual = pD;
      Object.assign(patch, autoDatasReais(it, patch, r.data));
      tx.update(Rf("itens", g.itemId), { ...patch, atualizadoEm: serverTimestamp(), atualizadoPor: ctx.uid });
      const a = calcAvanco(it)?.avanco ?? 0, b = calcAvanco({ ...it, ...patch })?.avanco ?? 0;
      if (q !== 0 || pD !== pA) {
        regHistCtx(tx, ctx, obraId, { alvoTipo: "item", alvoId: g.itemId, acao: "producao", resumo: `RDO ${fmtData(r.data)} aprovado (v${versao})`,
          campos: { ...(q !== 0 ? { qtdExecutada: { antes, depois } } : {}), ...(pD !== pA ? { percentual: { antes: pA, depois: pD } } : {}), ...(a !== b ? { avanco: { antes: a, depois: b } } : {}) } });
        regSerieCtx(tx, ctx, obraId, g.itemId, a, b, r.data);
      }
    }
    tx.update(Rf("rdos", rdoId), { status: "aprovado", versao, aprovadorUid: ctx.uid, aprovadorNome: ctx.nome, aprovadoEm: serverTimestamp(), lancamentos: ids,
      eventos: [...(r.eventos || []), { acao: "aprovar", versao, porUid: ctx.uid, porNome: ctx.nome, emISO: new Date().toISOString() }], atualizadoEm: serverTimestamp() });
    regHistCtx(tx, ctx, obraId, { alvoTipo: "rdo", alvoId: rdoId, acao: "aprovar", resumo: `RDO de ${fmtData(r.data)} aprovado (versão ${versao}, ${ids.length} lançamento(s))`, campos: { status: { antes: "enviado", depois: "aprovado" } } });
  });
}
function regHistCtx(w, ctx, obraId, h) {
  w.set(doc(collection(ctx.db, "obras", obraId, "historico")), { campos: null, ...h, autorUid: ctx.uid, autorNome: ctx.nome, em: serverTimestamp(), emISO: new Date().toISOString() });
}
function regSerieCtx(w, ctx, obraId, itemId, antes, depois, dataRef) {
  if (antes === depois) return;
  w.set(doc(collection(ctx.db, "obras", obraId, "serie")), { itemId, antes, depois, dataRef, autorUid: ctx.uid, em: serverTimestamp() });
}

/** Reabrir (→ rascunho) ou cancelar (→ cancelado) um RDO aprovado: estorna cada lançamento uma única vez. */
export async function estornarRdoTx(obraId, rdoId, novoStatus, motivo, ctx = { db: C.db, uid: uid(), nome: meuNome() }) {
  const Rf = (...p) => doc(ctx.db, "obras", obraId, ...p);
  await runTransaction(ctx.db, async tx => {
    const rs = await tx.get(Rf("rdos", rdoId));
    const r = rs.data();
    if (r.status !== "aprovado") throw new Error("Este RDO não está aprovado.");
    const lancs = [];
    for (const id of r.lancamentos || []) { const s = await tx.get(Rf("lancamentos", id)); if (s.exists()) lancs.push({ id, ...s.data() }); }
    const itens = {};
    for (const l of lancs) if (!itens[l.itemId]) { const s = await tx.get(Rf("itens", l.itemId)); itens[l.itemId] = s.exists() ? s.data() : null; }
    for (const l of lancs) if ((await tx.get(Rf("lancamentos", l.id + "__estorno"))).exists()) throw new Error("Este lançamento já foi estornado. Nada foi alterado.");
    for (const l of lancs) {
      const it = itens[l.itemId]; if (!it) continue;
      const antes = Number(it.qtdExecutada) || 0, depois = arred(antes - (Number(l.qtd) || 0));
      if (depois < -1e-9) throw new Error(`A tarefa ${it.codigo} ficaria com acumulado negativo (${fmtNum(depois)}). Revise os ajustes feitos depois da aprovação.`);
      const pAtual = Number(it.percentual) || 0;
      let pNovo = pAtual, revertido = true;
      if (l.percentualDepois !== l.percentualAntes) { if (pAtual === l.percentualDepois) pNovo = l.percentualAntes; else revertido = false; }
      const eref = Rf("lancamentos", l.id + "__estorno");
      tx.set(eref, { tipo: "estorno", origem: l.id, rdoId, versao: l.versao, itemId: l.itemId, codigo: it.codigo, modo: it.modo || "manual", qtd: -(Number(l.qtd) || 0) || 0, antes, depois,
        unidade: it.unidade || "", percentualAntes: pAtual, percentualDepois: pNovo, percentualRevertido: revertido, motivo, data: r.data, autorUid: ctx.uid, autorNome: ctx.nome, em: serverTimestamp() });
      const patch = { ultimoLanc: eref.id };
      if (l.qtd) patch.qtdExecutada = depois;
      if (pNovo !== pAtual) patch.percentual = pNovo;
      Object.assign(patch, autoDatasReais(it, patch, hojeISO()));
      tx.update(Rf("itens", l.itemId), { ...patch, atualizadoEm: serverTimestamp(), atualizadoPor: ctx.uid });
      const a = calcAvanco(it)?.avanco ?? 0, b = calcAvanco({ ...it, ...patch })?.avanco ?? 0;
      itens[l.itemId] = { ...it, ...patch };
      regHistCtx(tx, ctx, obraId, { alvoTipo: "item", alvoId: l.itemId, acao: "estorno", resumo: `Estorno do RDO ${fmtData(r.data)} (v${l.versao}) — ${motivo}${revertido ? "" : " · percentual não revertido: foi alterado depois da aprovação"}`,
        campos: { ...(l.qtd ? { qtdExecutada: { antes, depois } } : {}), ...(pNovo !== pAtual ? { percentual: { antes: pAtual, depois: pNovo } } : {}), ...(a !== b ? { avanco: { antes: a, depois: b } } : {}) } });
      regSerieCtx(tx, ctx, obraId, l.itemId, a, b, hojeISO());
    }
    tx.update(Rf("rdos", rdoId), { status: novoStatus, lancamentos: [], atualizadoEm: serverTimestamp(),
      eventos: [...(r.eventos || []), { acao: novoStatus === "cancelado" ? "cancelar" : "reabrir", motivo, versao: r.versao, porUid: ctx.uid, porNome: ctx.nome, emISO: new Date().toISOString() }] });
    regHistCtx(tx, ctx, obraId, { alvoTipo: "rdo", alvoId: rdoId, acao: novoStatus === "cancelado" ? "cancelar" : "reabrir",
      resumo: `RDO de ${fmtData(r.data)} ${novoStatus === "cancelado" ? "cancelado" : "reaberto para correção"} — ${lancs.length} lançamento(s) estornado(s). Motivo: ${motivo}`, campos: { status: { antes: "aprovado", depois: novoStatus } } });
  });
}
async function reabrirRdo() {
  const m = await perguntar("Reabrir para correção", "A produção lançada por este relatório será <b>estornada</b> das tarefas (o histórico fica guardado). Depois de corrigido e aprovado de novo, a produção entra outra vez com uma nova versão.", { campo: true, obrigatorio: true, botao: "Reabrir", placeholder: "Motivo da correção" });
  if (m === null) return;
  try { S.rdoEd.dirty = false; await estornarRdoTx(S.obraId, S.rdoEd.id, "rascunho", m); toast("RDO reaberto — produção estornada"); }
  catch (e) { $("rdo-err").textContent = erroMsg(e); }
}
async function cancelarRdo() {
  const r = S.rdoEd.v;
  const m = await perguntar("Cancelar RDO", r.status === "aprovado" ? "A produção lançada por este relatório será <b>estornada</b> das tarefas e o relatório ficará cancelado (não é apagado)." : "O relatório ficará cancelado (não é apagado).", { campo: true, obrigatorio: true, botao: "Cancelar RDO", perigo: true, placeholder: "Motivo do cancelamento" });
  if (m === null) return;
  try {
    S.rdoEd.dirty = false;
    if (r.status === "aprovado") await estornarRdoTx(S.obraId, S.rdoEd.id, "cancelado", m);
    else {
      const b = writeBatch(C.db);
      b.update(R("rdos", S.rdoEd.id), { status: "cancelado", eventos: [...(r.eventos || []), evento("cancelar", { motivo: m })], atualizadoEm: serverTimestamp() });
      regHist(b, S.obraId, { alvoTipo: "rdo", alvoId: S.rdoEd.id, acao: "cancelar", resumo: `RDO de ${fmtData(r.data)} cancelado — ${m}`, campos: { status: { antes: r.status, depois: "cancelado" } } });
      await b.commit();
    }
    toast("RDO cancelado");
  } catch (e) { $("rdo-err").textContent = erroMsg(e); }
}
async function excluirRdo() {
  const r = S.rdoEd.v;
  if ((await perguntar("Excluir rascunho", `Excluir o rascunho de ${fmtData(r.data)}? Esta ação não pode ser desfeita.`, { botao: "Excluir", perigo: true })) === null) return;
  try {
    const b = writeBatch(C.db);
    b.delete(R("rdos", r.id));
    regHist(b, S.obraId, { alvoTipo: "rdo", alvoId: r.id, acao: "excluir", resumo: `Rascunho do RDO de ${fmtData(r.data)} excluído` });
    await b.commit();
    for (const f of r.fotos || []) { const b2 = writeBatch(C.db); b2.delete(R("arquivos", f.id)); b2.commit().catch(() => {}); }
    S.rdoEd = null;
    location.hash = `#acompanhamento/${encodeURIComponent(S.obraId)}/rdo`;
  } catch (e) { $("rdo-err").textContent = erroMsg(e); }
}

// ---------------------------------------------------------------------
// Painel de acompanhamento
// ---------------------------------------------------------------------
function criterioAvanco(c) {
  return c.ponderado != null
    ? { rotulo: "Avanço físico ponderado", valor: c.ponderado, criterio: `Critério: soma de (peso × avanço) ÷ soma dos pesos das ${c.total} tarefas. Em tarefas por produção, avanço = executado ÷ previsto (limitado a 100%).` }
    : { rotulo: "Média simples do avanço das tarefas", valor: c.mediaSimples, criterio: `Não é avanço físico ponderado: ${c.semPeso} de ${c.total} tarefa(s) sem peso. Cada tarefa conta igual, independentemente do tamanho. Defina pesos (ex.: valor orçado) em todas as tarefas para ver o avanço físico.` };
}
export function serieEvolucao(tarefas, eventos, hoje = hojeISO()) {
  const ids = new Set(tarefas.map(t => t.id)), N = tarefas.length;
  if (!N) return [];
  const pesos = new Map(tarefas.map(t => [t.id, Number(t.peso) || 0]));
  const ponderado = tarefas.every(t => Number(t.peso) > 0);
  const somaP = [...pesos.values()].reduce((a, b) => a + b, 0);
  const ev = eventos.filter(e => ids.has(e.itemId) && e.dataRef <= hoje).sort((a, b) => a.dataRef.localeCompare(b.dataRef) || (a.ord || 0) - (b.ord || 0));
  const fins = tarefas.map(t => t.fimPrev).filter(Boolean).sort();
  const datas = [...new Set([...ev.map(e => e.dataRef), ...fins, hoje])].sort();
  const inicio = minData([...tarefas.map(t => t.inicioPrev), ev[0]?.dataRef, datas[0]]);
  const est = new Map(); let k = 0;
  const pts = [];
  const ponto = (d, atual) => {
    let conc = 0, s = 0, sp = 0;
    for (const t of tarefas) { const av = atual ? (t.calc.avanco ?? 0) : (est.get(t.id) || 0); if (av >= 100 - 1e-9) conc++; s += av; sp += (pesos.get(t.id) || 0) * av; }
    const prev = fins.filter(f => f <= d).length / N * 100;
    return { data: d, previsto: prev, concluidas: conc / N * 100, avanco: ponderado ? sp / somaP : s / N };
  };
  if (inicio) pts.push({ data: somaDias(inicio, -1), previsto: 0, concluidas: 0, avanco: 0 });
  for (const d of datas) {
    while (k < ev.length && ev[k].dataRef <= d) { est.set(ev[k].itemId, ev[k].depois); k++; }
    pts.push(ponto(d, d === hoje));
  }
  return pts.filter(p => p.data <= maxData([hoje, ...fins]));
}
function graficoEvolucao(pts, rotuloAvanco) {
  if (pts.length < 2) return `<div class="go-muted">Ainda não há registros de avanço suficientes para a curva.</div>`;
  const W = 760, H = 260, L = 40, Rr = 120, T = 14, B = 28;
  const t0 = new Date(pts[0].data + "T12:00").getTime(), t1 = new Date(pts[pts.length - 1].data + "T12:00").getTime();
  const x = d => L + (new Date(d + "T12:00").getTime() - t0) / Math.max(1, t1 - t0) * (W - L - Rr);
  const y = v => T + (100 - v) / 100 * (H - T - B);
  const SER = [["previsto", "Previsto (término)", "#3F4A5C", "5 4"], ["concluidas", "Tarefas concluídas", "#1D9E75", ""], ["avanco", rotuloAvanco, "#0E6FCC", ""]];
  const path = k => pts.map((p, i) => `${i ? "L" : "M"}${x(p.data).toFixed(1)},${y(p[k]).toFixed(1)}`).join("");
  const ult = pts[pts.length - 1];
  const hojeX = x(hojeISO());
  const meses = []; { const d = new Date(t0); d.setDate(1); d.setMonth(d.getMonth() + 1); while (d.getTime() <= t1) { meses.push(hojeISO(d)); d.setMonth(d.getMonth() + 1); } }
  const labels = SER.map(([k, n]) => ({ k, n, y: y(ult[k]) })).sort((a, b) => a.y - b.y);
  for (let i = 1; i < labels.length; i++) if (labels[i].y - labels[i - 1].y < 14) labels[i].y = labels[i - 1].y + 14;
  return `<div class="go-graf" data-pts='${esc(JSON.stringify(pts))}'>
    <div class="go-leg">${SER.map(([, n, c, d]) => `<span><svg width="22" height="8"><line x1="0" y1="4" x2="22" y2="4" stroke="${c}" stroke-width="2" ${d ? `stroke-dasharray="${d}"` : ""}/></svg>${esc(n)}</span>`).join("")}</div>
    <svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Evolução da execução ao longo do tempo">
      ${[0, 25, 50, 75, 100].map(v => `<line x1="${L}" x2="${W - Rr}" y1="${y(v)}" y2="${y(v)}" stroke="#E6EBF2"/><text x="${L - 6}" y="${y(v) + 4}" text-anchor="end" class="go-ax">${v}%</text>`).join("")}
      ${meses.map(m => `<text x="${x(m)}" y="${H - 8}" text-anchor="middle" class="go-ax">${m.slice(5, 7)}/${m.slice(2, 4)}</text>`).join("")}
      ${hojeX <= W - Rr ? `<line x1="${hojeX}" x2="${hojeX}" y1="${T}" y2="${H - B}" stroke="#C8D2E0" stroke-dasharray="2 3"/><text x="${hojeX}" y="${T - 3}" text-anchor="middle" class="go-ax">hoje</text>` : ""}
      ${SER.map(([k, , c, d]) => `<path d="${path(k)}" fill="none" stroke="${c}" stroke-width="2" stroke-linejoin="round" ${d ? `stroke-dasharray="${d}"` : ""}/>`).join("")}
      ${labels.map(l => `<text x="${W - Rr + 6}" y="${l.y + 4}" class="go-lab">${fmtPct(ult[l.k])} ${esc(l.n.split(" ")[0].replace("(", ""))}</text>`).join("")}
      <line class="go-cross" x1="0" x2="0" y1="${T}" y2="${H - B}" stroke="#555" stroke-width="1" style="display:none"/>
      <rect x="${L}" y="${T}" width="${W - L - Rr}" height="${H - T - B}" fill="transparent" class="go-hit"/>
    </svg><div class="go-tip" style="display:none"></div></div>`;
}
function ligarGrafico(el) {
  if (!el) return;
  const pts = JSON.parse(el.dataset.pts), svg = el.querySelector("svg"), cross = el.querySelector(".go-cross"), tip = el.querySelector(".go-tip");
  const W = 760, L = 40, Rr = 120;
  const t0 = new Date(pts[0].data + "T12:00").getTime(), t1 = new Date(pts[pts.length - 1].data + "T12:00").getTime();
  el.querySelector(".go-hit").addEventListener("mousemove", ev => {
    const r = svg.getBoundingClientRect(), vx = (ev.clientX - r.left) / r.width * W;
    const t = t0 + (vx - L) / (W - L - Rr) * (t1 - t0);
    let p = pts[0]; for (const q of pts) if (new Date(q.data + "T12:00").getTime() <= t) p = q;
    const px = L + (new Date(p.data + "T12:00").getTime() - t0) / Math.max(1, t1 - t0) * (W - L - Rr);
    cross.setAttribute("x1", px); cross.setAttribute("x2", px); cross.style.display = "";
    tip.style.display = "block"; tip.style.left = Math.min(r.width - 170, Math.max(0, px / W * r.width + 10)) + "px";
    tip.innerHTML = `<b>${fmtData(p.data)}</b><br>Previsto: ${fmtPct(p.previsto)}<br>Concluídas: ${fmtPct(p.concluidas)}<br>Avanço: ${fmtPct(p.avanco)}`;
  });
  el.querySelector(".go-hit").addEventListener("mouseleave", () => { cross.style.display = "none"; tip.style.display = "none"; });
}

function renderPainel() {
  const M = S.modelo, c = M.cons, P = pode(), hoje = hojeISO();
  const crit = criterioAvanco(c);
  const etapas = M.raiz.map(id => M.der.get(id));
  const venc = M.tarefas.filter(t => t.fimPrev && t.fimPrev >= hoje && t.fimPrev <= somaDias(hoje, DIAS_VENCIMENTO) && t.situacao !== "concluida").sort((a, b) => a.fimPrev.localeCompare(b.fimPrev));
  const atras = M.tarefas.filter(t => t.atrasada).sort((a, b) => a.fimPrev.localeCompare(b.fimPrev));
  const pend = S.rdos.filter(r => r.status === "enviado");
  const recentes = S.rdos.filter(r => r.status !== "cancelado").slice(0, 6);
  const porUn = new Map();
  for (const t of M.tarefas) if (t.modo === "producao") {
    const u = t.unidade || "(sem unidade)"; const g = porUn.get(u) || { prev: 0, exe: 0, n: 0, acima: 0 };
    g.prev += Number(t.qtdPrevista) || 0; g.exe += Number(t.qtdExecutada) || 0; g.n++; if (t.calc.acima) g.acima++; porUn.set(u, g);
  }
  const base = `#acompanhamento/${encodeURIComponent(S.obraId)}`;
  const listaT = arr => arr.length ? `<div class="go-mini">${arr.slice(0, 8).map(t => `<div onclick="GO.abrirItem('${esc(t.id)}')"><span><b>${esc(t.codigo)}</b> ${esc(t.descricao)}</span><span>${fmtData(t.fimPrev)} · ${fmtPct(t.calc.avanco)}</span></div>`).join("")}${arr.length > 8 ? `<div class="go-muted">+ ${arr.length - 8} outras</div>` : ""}</div>` : `<div class="go-muted">Nenhuma.</div>`;
  $("obra-corpo").innerHTML = !c.total ? `<div class="obra-vazio"><div class="obra-vazio-ic">📋</div><h3>Sem tarefas cadastradas</h3><p>Cadastre a EAP em <a href="${base}/tarefas">Tarefas</a> (manualmente ou importando uma planilha) para ver os indicadores.</p></div>` : `
  <div class="go-kpis">
    <div class="go-kpi"><span>Total de tarefas</span><b>${c.total}</b></div>
    <div class="go-kpi"><span>Não iniciadas</span><b>${c.nao_iniciada}</b></div>
    <div class="go-kpi"><span>Em andamento</span><b>${c.em_andamento}</b></div>
    <div class="go-kpi"><span>Concluídas</span><b>${c.concluida}</b></div>
    <div class="go-kpi go-kpi-alerta"><span>Atrasadas</span><b>${c.atrasadas}</b></div>
  </div>
  <div class="go-cols">
    <section class="go-sec go-indic"><div class="go-indic-v">${fmtPct(c.pctConcluidas)}</div><div><b>Tarefas concluídas</b><div class="go-dica">${c.concluida} de ${c.total} tarefas com 100%. Conta quantidade de tarefas, não o tamanho delas.</div></div></section>
    <section class="go-sec go-indic"><div class="go-indic-v ${c.ponderado == null ? "go-indic-simples" : ""}">${fmtPct(crit.valor)}</div><div><b>${crit.rotulo}</b><div class="go-dica">${esc(crit.criterio)}</div></div></section>
  </div>
  ${c.semPrevisto ? `<div class="go-alerta-box">⚠️ ${c.semPrevisto} tarefa(s) por produção sem quantidade prevista — o avanço delas não pode ser calculado e entra como 0%.</div>` : ""}
  <section class="go-sec"><h3>Evolução da execução</h3><div id="pn-graf"><div class="go-muted">Carregando…</div></div>
    <div class="go-dica">Previsto = % de tarefas com término previsto até a data. A curva de avanço usa ${c.ponderado != null ? "os pesos atuais das tarefas" : "a média simples (sem pesos)"} e é reconstruída a partir dos lançamentos registrados no sistema.</div></section>
  <div class="go-cols">
    <section class="go-sec"><h3>Avanço por etapa</h3>${etapas.map(e => { const ce = e.cons, cr = criterioAvanco(ce); return `<div class="go-etapa" onclick="location.hash='${base}/tarefas';GO.filtrarEtapa('${esc(e.id)}')">
      <div class="go-etapa-t"><b>${esc(e.codigo)}</b> ${esc(e.descricao)}</div>
      <div class="go-etapa-b">${barra(ce.total ? cr.valor : 0)}<span class="go-pct">${ce.total ? fmtPct(cr.valor) + (ce.ponderado == null ? "*" : "") : "—"}</span></div>
      <div class="go-muted">${ce.concluida}/${ce.total} concluídas${ce.atrasadas ? ` · <span class="go-chip c-atraso">${ce.atrasadas} atrasada(s)</span>` : ""}</div></div>`; }).join("")}
      ${etapas.some(e => e.cons.total && e.cons.ponderado == null) ? `<div class="go-dica">* média simples (há tarefas sem peso na etapa).</div>` : ""}</section>
    <section class="go-sec"><h3>Produção por unidade</h3>${porUn.size ? `<table class="go-tab"><thead><tr><th>Unidade</th><th>Tarefas</th><th>Executado</th><th>Previsto</th><th>%</th></tr></thead><tbody>${[...porUn].map(([u, g]) => `<tr><td>${esc(u)}</td><td>${g.n}</td><td>${fmtNum(g.exe)}</td><td>${fmtNum(g.prev)}</td><td>${g.prev > 0 ? fmtPct(g.exe / g.prev * 100) : "—"}${g.acima ? ` <span class="go-alerta" title="${g.acima} tarefa(s) acima do previsto">▲${g.acima}</span>` : ""}</td></tr>`).join("")}</tbody></table>
      <div class="go-dica">Quantidades de unidades diferentes (m², m³, kg…) não são somadas entre si.</div>` : `<div class="go-muted">Nenhuma tarefa acompanhada por produção.</div>`}</section>
  </div>
  <div class="go-cols">
    <section class="go-sec"><h3>Vencem nos próximos ${DIAS_VENCIMENTO} dias <span class="go-muted">(${venc.length})</span></h3>${listaT(venc)}</section>
    <section class="go-sec"><h3>Atrasadas <span class="go-muted">(${atras.length})</span></h3>${listaT(atras)}</section>
  </div>
  <div class="go-cols">
    <section class="go-sec"><h3>RDOs aguardando aprovação <span class="go-muted">(${pend.length})</span></h3>${pend.length ? `<div class="go-mini">${pend.map(r => `<div onclick="location.hash='${base}/rdo/${r.id}'"><span>${fmtData(r.data)} · ${esc(r.autorNome || "")}</span><span>${(r.atividades || []).length} atividade(s)</span></div>`).join("")}</div>` : `<div class="go-muted">Nenhum.</div>`}</section>
    <section class="go-sec"><h3>Relatórios recentes</h3>${recentes.length ? `<div class="go-mini">${recentes.map(r => `<div onclick="location.hash='${base}/rdo/${r.id}'"><span>${fmtData(r.data)} ${chipRdo(r.status)}</span><span>${esc(r.autorNome || "")}</span></div>`).join("")}</div>` : `<div class="go-muted">Nenhum relatório ainda.</div>`}</section>
  </div>
  ${P.gerir ? `<section class="go-sec"><h3>Conferência</h3><div class="go-dica">Recalcula a quantidade executada de cada tarefa a partir dos lançamentos (RDOs, estornos e ajustes) e aponta diferenças.</div><button class="btn-sm" onclick="GO.conferir(this)">Conferir acumulados</button><div id="pn-conf"></div></section>` : ""}`;
  if (c.total) carregarSerie(crit.rotulo);
}
async function carregarSerie(rotulo) {
  const alvo = $("pn-graf");
  try {
    const snap = await getDocs(COL("serie"));
    const ev = snap.docs.map((d, i) => ({ ...d.data(), ord: tsData(d.data().em)?.getTime() || i }));
    if (!$("pn-graf")) return;
    alvo.innerHTML = graficoEvolucao(serieEvolucao(S.modelo.tarefas, ev), rotulo);
    ligarGrafico(alvo.querySelector(".go-graf"));
  } catch (e) { alvo.textContent = erroMsg(e); }
}
async function conferir(btn) {
  btn.disabled = true;
  try {
    const snap = await getDocs(COL("lancamentos"));
    const soma = new Map();
    snap.docs.forEach(d => { const l = d.data(); soma.set(l.itemId, arred((soma.get(l.itemId) || 0) + (Number(l.qtd) || 0))); });
    const dif = S.itens.filter(i => i.tipo === "tarefa").filter(i => Math.abs((soma.get(i.id) || 0) - (Number(i.qtdExecutada) || 0)) > 1e-6);
    $("pn-conf").innerHTML = dif.length ? `<div class="go-alerta-box">${dif.length} tarefa(s) com diferença: ${dif.map(i => `${esc(i.codigo)} (tarefa ${fmtNum(i.qtdExecutada || 0)} × lançamentos ${fmtNum(soma.get(i.id) || 0)})`).join("; ")}</div>`
      : `<div class="go-ok-box">✔ ${snap.size} lançamento(s) conferido(s). Todos os acumulados batem.</div>`;
  } catch (e) { $("pn-conf").textContent = erroMsg(e); }
  btn.disabled = false;
}

// ---------------------------------------------------------------------
// Bibliotecas externas (carregadas só quando usadas)
// ---------------------------------------------------------------------
const _libs = {};
function carregarScript(url) {
  return _libs[url] || (_libs[url] = new Promise((res, rej) => {
    const s = document.createElement("script"); s.src = url; s.onload = res;
    s.onerror = () => { delete _libs[url]; rej(new Error("Não foi possível carregar a biblioteca (" + url.split("/")[2] + "). Verifique a internet.")); };
    document.head.appendChild(s);
  }));
}
async function libXlsx() { if (!window.XLSX) await carregarScript(LIB.xlsx); return window.XLSX; }
async function libPdf() { if (!window.jspdf) await carregarScript(LIB.jspdf); if (!window.jspdf.jsPDF.API.autoTable) await carregarScript(LIB.autotable); return window.jspdf.jsPDF; }
const serialExcel = iso => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5; };
const isoDeSerial = n => { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(n) * 864e5); return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`; };
export function parseDataPlanilha(v) {
  if (v == null || v === "") return { iso: null };
  if (typeof v === "number" && v > 20000 && v < 80000) return { iso: isoDeSerial(v) };
  if (v instanceof Date && !isNaN(v)) return { iso: hojeISO(v) };
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})$/);
  if (m) { const y = m[3].length === 2 ? 2000 + +m[3] : +m[3]; const iso = `${y}-${pad(m[2])}-${pad(m[1])}`; return validaIso(iso) ? { iso } : { erro: `data inválida “${s}”` }; }
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m && validaIso(m[0])) return { iso: m[0] };
  return { erro: `data inválida “${s}” (use dd/mm/aaaa)` };
}
const validaIso = iso => { const [y, m, d] = iso.split("-").map(Number); const dt = new Date(y, m - 1, d); return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d; };

// ---------------------------------------------------------------------
// Exportação (Excel e PDF) — respeita os filtros atuais
// ---------------------------------------------------------------------
function linhasExport() {
  const ativo = filtroTarefaAtivo() || S.ui.etapa || S.ui.soEtapas;
  if (!ativo) { const out = []; const w = ids => ids.forEach(id => { const d = S.modelo.der.get(id); out.push(d); w(d.filhos); }); w(S.modelo.raiz); return out; }
  const salvo = S.ui.abertos; S.ui.abertos = new Set(S.itens.map(i => i.id));
  const r = linhasVisiveis(); S.ui.abertos = salvo; return r;
}
function descricaoFiltros() {
  const u = S.ui, p = [];
  if (u.busca) p.push(`busca “${u.busca}”`);
  if (u.situacao) p.push(u.situacao === "atrasada" ? "atrasadas" : SITUACOES_TAREFA[u.situacao]);
  if (u.resp) p.push(u.resp === "_sem" ? "sem responsável" : "resp. " + nomeUsuario(u.resp));
  if (u.etapa) p.push("etapa " + (S.modelo.der.get(u.etapa)?.codigo || ""));
  if (u.prazo) p.push({ atrasadas: "prazo vencido", 7: "vence em 7 dias", 30: "vence em 30 dias", sem: "sem prazo" }[u.prazo]);
  if (u.soEtapas) p.push("somente etapas");
  if (u.ocultarConcluidas) p.push("sem concluídas");
  return p.length ? "Filtros: " + p.join(", ") : "Todos os itens";
}
const avancoExport = d => d.tipo === "tarefa" ? (d.calc.avanco == null ? null : (d.calc.acima ? d.calc.real : d.calc.avanco)) : (d.cons.total ? (d.cons.ponderado ?? d.cons.mediaSimples) : null);
async function exportarXlsx() {
  try {
    const X = await libXlsx(), o = obraAtual(), rows = linhasExport();
    const cab = ["Código", "Tipo", "Descrição", "Responsável", "Início previsto", "Término previsto", "Início real", "Término real", "Modo", "Qtd. prevista", "Unidade", "Qtd. executada", "Avanço (%)", "Situação", "Atrasada", "Peso", "Observações"];
    const aoa = [[`${o.nome}${o.demo ? " (DEMONSTRAÇÃO — dados fictícios)" : ""} — Lista de tarefas`], [`Emitido em ${fmtDataHora(new Date())} · ${descricaoFiltros()}`], [], cab];
    for (const d of rows) {
      const t = d.tipo === "tarefa";
      aoa.push([d.codigo, TIPOS[d.tipo], "  ".repeat(d.nivel) + d.descricao, d.responsavelNome || nomeUsuario(d.responsavelUid) || "",
        t ? d.inicioPrev : d.inicioCalc, t ? d.fimPrev : d.fimCalc, t ? d.inicioReal : null, t ? d.fimReal : null,
        t ? (d.modo === "producao" ? "Produção" : "Percentual") : "", t && d.modo === "producao" ? d.qtdPrevista ?? null : null, t ? d.unidade || "" : "",
        t && d.modo === "producao" ? d.qtdExecutada || 0 : null, avancoExport(d) == null ? null : Math.round(avancoExport(d) * 10) / 10,
        SITUACOES_TAREFA[d.situacao], d.atrasada ? "Sim" : "Não", t ? d.peso ?? null : null, d.observacoes || ""]);
    }
    const ws = X.utils.aoa_to_sheet(aoa);
    for (let r = 4; r < aoa.length; r++) for (const c of [4, 5, 6, 7]) {
      const v = aoa[r][c], ref = X.utils.encode_cell({ r, c });
      if (v) ws[ref] = { t: "n", v: serialExcel(v), z: "dd/mm/yyyy" }; else delete ws[ref];
    }
    ws["!cols"] = [10, 9, 46, 18, 12, 12, 12, 12, 10, 12, 8, 12, 10, 13, 9, 8, 30].map(w => ({ wch: w }));
    const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Tarefas");
    X.writeFile(wb, `tarefas-${o.id}-${hojeISO()}.xlsx`);
  } catch (e) { toast(erroMsg(e)); }
}
async function exportarPdf() {
  try {
    const jsPDF = await libPdf(), o = obraAtual(), rows = linhasExport(), c = S.modelo.cons, cr = criterioAvanco(c);
    const pdf = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
    pdf.setFontSize(13); pdf.text(`${o.nome}${o.demo ? " — DEMONSTRAÇÃO (dados fictícios)" : ""}`, 12, 13);
    pdf.setFontSize(8.5); pdf.setTextColor(90);
    pdf.text(`Lista de tarefas · emitido em ${fmtDataHora(new Date())} · ${descricaoFiltros()}`, 12, 18.5);
    pdf.text(`Tarefas: ${c.total} · concluídas ${c.concluida} (${fmtPct(c.pctConcluidas)}) · em andamento ${c.em_andamento} · não iniciadas ${c.nao_iniciada} · atrasadas ${c.atrasadas} · ${cr.rotulo}: ${fmtPct(cr.valor)}`, 12, 23);
    pdf.autoTable({
      startY: 27, margin: { left: 10, right: 10 }, styles: { fontSize: 7.2, cellPadding: 1.2 }, headStyles: { fillColor: [13, 43, 110] },
      head: [["Código", "Descrição", "Responsável", "Previsto", "Real", "Produção", "Avanço", "Situação"]],
      body: rows.map(d => { const t = d.tipo === "tarefa"; return [d.codigo, "   ".repeat(d.nivel) + d.descricao, d.responsavelNome || nomeUsuario(d.responsavelUid) || "",
        `${fmtData(t ? d.inicioPrev : d.inicioCalc)} – ${fmtData(t ? d.fimPrev : d.fimCalc)}`, t ? `${fmtData(d.inicioReal)} – ${fmtData(d.fimReal)}` : "",
        t ? (d.modo === "producao" ? `${fmtNum(d.qtdExecutada || 0)} / ${fmtNum(d.qtdPrevista || 0)} ${d.unidade || ""}` : "manual") : `${d.cons.concluida}/${d.cons.total} concl.`,
        fmtPct(avancoExport(d)), SITUACOES_TAREFA[d.situacao] + (d.atrasada ? " · ATRASADA" : "")]; }),
      columnStyles: { 0: { cellWidth: 18 }, 1: { cellWidth: 82 }, 2: { cellWidth: 28 }, 3: { cellWidth: 34 }, 4: { cellWidth: 34 }, 5: { cellWidth: 30 }, 6: { cellWidth: 15, halign: "right" } },
      didParseCell: h => { if (h.section === "body" && rows[h.row.index].tipo !== "tarefa") { h.cell.styles.fontStyle = "bold"; h.cell.styles.fillColor = rows[h.row.index].tipo === "etapa" ? [230, 238, 250] : [244, 247, 252]; } },
      didDrawPage: () => { pdf.setFontSize(7); pdf.setTextColor(140); pdf.text(`Construtora Vieira & Silva · página ${pdf.internal.getNumberOfPages()}`, 287, 205, { align: "right" }); }
    });
    if (c.ponderado == null) { pdf.setFontSize(7); pdf.setTextColor(120); pdf.text(cr.criterio, 10, Math.min(200, pdf.lastAutoTable.finalY + 5), { maxWidth: 270 }); }
    pdf.save(`tarefas-${o.id}-${hojeISO()}.pdf`);
  } catch (e) { toast(erroMsg(e)); }
}
async function pdfRdo() {
  try {
    const jsPDF = await libPdf(), o = obraAtual(), r = S.rdoEd.v;
    const pdf = new jsPDF({ unit: "mm", format: "a4" });
    pdf.setFontSize(13); pdf.text(`Relatório Diário de Obra — ${fmtData(r.data)}`, 12, 14);
    pdf.setFontSize(9); pdf.setTextColor(80);
    pdf.text(`${o.nome}${o.demo ? " (DEMONSTRAÇÃO — dados fictícios)" : ""} · Situação: ${STATUS_RDO[r.status]}${r.versao ? " (v" + r.versao + ")" : ""}`, 12, 20);
    pdf.text(`Autor: ${r.autorNome || "—"} · Aprovação: ${r.aprovadorNome || "—"}${r.aprovadoEm ? " em " + fmtDataHora(r.aprovadoEm) : ""}`, 12, 25);
    const cl = r.clima || {};
    const tab = (y, head, body, opts = {}) => { pdf.autoTable({ startY: y, head: [head], body, margin: { left: 12, right: 12 }, styles: { fontSize: 8 }, headStyles: { fillColor: [13, 43, 110] }, ...opts }); return pdf.lastAutoTable.finalY + 4; };
    let y = tab(30, ["Manhã", "Tarde", "Noite", "Condição", "Horas paralisadas", "Impacto"], [[cl.manha || "—", cl.tarde || "—", cl.noite || "—", CONDICAO[cl.condicao] || "—", fmtNum(cl.horasParadas || 0), cl.impacto || ""]]);
    y = tab(y, ["Função", "Empresa", "Qtd."], (r.maoDeObra || []).map(m => [m.funcao, m.empresa, m.qtd]));
    y = tab(y, ["Equipamento", "Situação", "Qtd."], (r.equipamentos || []).map(m => [m.nome, m.situacao, m.qtd]));
    y = tab(y, ["Código", "Atividade", "Qtd. do dia", "Avanço após (%)", "Observação"], (r.atividades || []).map(a => [a.codigo || "", a.descricao, a.qtd ? `${fmtNum(parseNum(a.qtd))} ${a.unidade || ""}` : "", a.percentual !== "" && a.percentual != null ? fmtPct(parseNum(a.percentual)) : "", a.obs || ""]));
    for (const [t, v] of [["Ocorrências", r.ocorrencias], ["Impedimentos", r.impedimentos], ["Observações", r.observacoes]]) if (v) { y = tab(y, [t], [[v]]); }
    const fotos = (r.fotos || []).filter(f => f.thumb);
    if (fotos.length) {
      if (y > 230) { pdf.addPage(); y = 15; }
      pdf.setFontSize(9); pdf.setTextColor(60); pdf.text("Fotos (miniaturas)", 12, y); y += 3;
      fotos.forEach((f, i) => { const col = i % 4, x = 12 + col * 47; if (col === 0 && i) y += 40; if (y > 255) { pdf.addPage(); y = 15; } try { pdf.addImage(f.thumb, "JPEG", x, y, 44, 33); } catch (e) {} pdf.setFontSize(6.5); pdf.text((f.legenda || f.nome || "").slice(0, 40), x, y + 36); });
    }
    pdf.save(`rdo-${o.id}-${r.data}.pdf`);
  } catch (e) { toast(erroMsg(e)); }
}

// ---------------------------------------------------------------------
// Importação por planilha (prévia, validação e relatório de erros)
// ---------------------------------------------------------------------
const COLS_IMPORT = {
  codigo: ["codigo", "cod", "item", "codigo eap", "eap"], tipo: ["tipo", "nivel"], descricao: ["descricao", "servico", "atividade", "tarefa", "nome", "discriminacao"],
  pai: ["codigo pai", "pai", "vinculo", "etapa pai"], responsavel: ["responsavel", "resp"], inicio: ["inicio previsto", "inicio", "data inicio", "data de inicio"],
  fim: ["termino previsto", "termino", "fim", "data termino", "data fim", "data de termino"], modo: ["modo", "acompanhamento", "tipo de acompanhamento"],
  qtd: ["qtd prevista", "quantidade prevista", "quantidade", "qtd", "qtde", "quant"], unidade: ["unidade", "un", "und", "unid"], peso: ["peso", "peso (%)"], obs: ["observacoes", "observacao", "obs"]
};
const normCab = s => semAcento(s).replace(/[.:]/g, "").replace(/\s+/g, " ").trim();
/** Valida as linhas da planilha contra a obra. Puro — testado separadamente. */
export function validarImportacao(linhasBrutas, { existentes = [], usuarios = [] } = {}) {
  const cabIdx = linhasBrutas.slice(0, 15).findIndex(l => { const n = l.map(normCab); return n.some(x => COLS_IMPORT.codigo.includes(x)) && n.some(x => COLS_IMPORT.descricao.includes(x)); });
  if (cabIdx < 0) return { erroGeral: "Não encontrei o cabeçalho. A planilha precisa ter ao menos as colunas “Código” e “Descrição”.", linhas: [] };
  const cab = linhasBrutas[cabIdx].map(normCab), col = {};
  for (const [k, nomes] of Object.entries(COLS_IMPORT)) { const i = cab.findIndex(c => nomes.includes(c)); if (i >= 0) col[k] = i; }
  const get = (l, k) => col[k] == null ? "" : l[col[k]];
  const exist = new Map(existentes.map(i => [i.codigoKey || codigoKey(i.codigo), i]));
  const linhas = [];
  linhasBrutas.slice(cabIdx + 1).forEach((l, j) => {
    if (!l.some(v => String(v ?? "").trim() !== "")) return;
    linhas.push({ n: cabIdx + j + 2, erros: [], avisos: [], cod: normCodigo(get(l, "codigo")), tipoInf: semAcento(get(l, "tipo")), descricao: String(get(l, "descricao") ?? "").trim(),
      paiCod: normCodigo(get(l, "pai")), resp: String(get(l, "responsavel") ?? "").trim(), ini: get(l, "inicio"), fim: get(l, "fim"), modoInf: semAcento(get(l, "modo")),
      qtd: get(l, "qtd"), unidade: String(get(l, "unidade") ?? "").trim(), peso: get(l, "peso"), obs: String(get(l, "obs") ?? "").trim() });
  });
  const vistos = new Map();
  for (const L of linhas) {
    const ec = validaCodigo(L.cod); if (ec) L.erros.push(ec);
    if (!L.descricao) L.erros.push("Descrição vazia.");
    if (L.cod && vistos.has(codigoKey(L.cod))) L.erros.push(`Código repetido na planilha (linha ${vistos.get(codigoKey(L.cod)).n}).`);
    else if (L.cod) vistos.set(codigoKey(L.cod), L);
    if (L.cod && exist.has(codigoKey(L.cod))) L.erros.push("Código já existe nesta obra.");
    if (!L.paiCod && L.cod.includes(".")) L.paiCod = L.cod.slice(0, L.cod.lastIndexOf("."));
    const di = parseDataPlanilha(L.ini), df = parseDataPlanilha(L.fim);
    if (di.erro) L.erros.push("Início: " + di.erro); if (df.erro) L.erros.push("Término: " + df.erro);
    L.inicioPrev = di.iso || null; L.fimPrev = df.iso || null;
    if (L.inicioPrev && L.fimPrev && L.fimPrev < L.inicioPrev) L.erros.push("Término anterior ao início.");
    const q = L.qtd === "" || L.qtd == null ? null : parseNum(L.qtd);
    if (q != null && (isNaN(q) || q < 0)) L.erros.push(`Quantidade inválida (“${L.qtd}”).`);
    L.qtdPrevista = q != null && !isNaN(q) ? q : null;
    const p = L.peso === "" || L.peso == null ? null : parseNum(L.peso);
    if (p != null && (isNaN(p) || p < 0)) L.erros.push(`Peso inválido (“${L.peso}”).`);
    L.pesoN = p != null && !isNaN(p) ? p : null;
    if (L.resp) {
      const r = semAcento(L.resp), u = usuarios.find(x => semAcento(x.email) === r || semAcento(x.nome) === r);
      if (u) L.respUid = u.id; else L.avisos.push(`Responsável “${L.resp}” não encontrado — ficará em branco.`);
    }
  }
  // Tipos e vínculos
  const temFilhos = new Set(linhas.map(L => codigoKey(L.paiCod)).filter(Boolean));
  const porCod = new Map(linhas.filter(L => L.cod).map(L => [codigoKey(L.cod), L]));
  const tipoDe = c => { const L = porCod.get(c); if (L) return L.tipo; const e = exist.get(c); return e ? e.tipo : null; };
  const ordenadas = [...linhas].sort((a, b) => a.cod.split(".").length - b.cod.split(".").length);
  for (const L of ordenadas) {
    const ti = L.tipoInf;
    const paiTipo = L.paiCod ? tipoDe(codigoKey(L.paiCod)) : null;
    if (ti.startsWith("etapa")) L.tipo = "etapa"; else if (ti.startsWith("sub")) L.tipo = "subetapa"; else if (ti.startsWith("tarefa") || ti.startsWith("servico") || ti.startsWith("atividade")) L.tipo = "tarefa";
    else if (ti) { L.erros.push(`Tipo “${L.tipoInf}” desconhecido (use Etapa, Subetapa ou Tarefa).`); L.tipo = "tarefa"; }
    else L.tipo = !L.paiCod ? "etapa" : (temFilhos.has(codigoKey(L.cod)) ? (paiTipo === "etapa" ? "subetapa" : "etapa") : "tarefa");
    if (L.tipo === "etapa" && L.paiCod && !ti) { L.erros.push("Item com filhos dentro de uma subetapa: o limite é Etapa › Subetapa › Tarefa."); }
    if (L.tipo === "etapa" && L.paiCod && ti) L.erros.push("Etapa não pode ter vínculo (código pai).");
    if (L.tipo !== "etapa") {
      if (!L.paiCod) L.erros.push(`${TIPOS[L.tipo]} precisa estar dentro de uma etapa (informe o código pai ou use código 1.1, 1.2…).`);
      else if (!paiTipo) L.erros.push(`Código pai “${L.paiCod}” não existe na planilha nem na obra.`);
      else if (L.tipo === "subetapa" && paiTipo !== "etapa") L.erros.push("Subetapa precisa estar dentro de uma etapa.");
      else if (L.tipo === "tarefa" && paiTipo === "tarefa") L.erros.push(`O pai “${L.paiCod}” é uma tarefa; tarefas não têm itens dentro.`);
    }
    if (L.tipo === "tarefa") {
      const m = L.modoInf;
      L.modo = m.startsWith("prod") || m.startsWith("quant") ? "producao" : m.startsWith("perc") || m.startsWith("manual") ? "manual" : (L.qtdPrevista != null || L.unidade ? "producao" : "manual");
      if (L.modo === "producao" && !L.unidade) L.erros.push("Tarefa por produção sem unidade de medida.");
      if (L.modo === "producao" && L.qtdPrevista == null) L.erros.push("Tarefa por produção sem quantidade prevista.");
      if (L.modo === "producao" && L.qtdPrevista === 0) L.avisos.push("Quantidade prevista igual a zero: o avanço não poderá ser calculado.");
    } else if (L.qtdPrevista != null || L.unidade) L.avisos.push("Quantidade/unidade ignoradas: só tarefas têm produção.");
  }
  // Itens cujo pai tem erro também não entram
  let mudou = true;
  while (mudou) {
    mudou = false;
    for (const L of linhas) {
      if (L.erros.length || !L.paiCod) continue;
      const P = porCod.get(codigoKey(L.paiCod));
      if (P && P !== L && P.erros.length) { L.erros.push(`O item pai “${L.paiCod}” tem erro.`); mudou = true; }
    }
  }
  return { linhas, colunas: Object.keys(col) };
}

async function importar() {
  abrirModal({ titulo: "Importar tarefas de planilha", largura: "980px",
    corpo: `<div class="go-dica" style="margin:0">Aceita .xlsx, .xls ou .csv. Colunas reconhecidas: <b>Código</b> e <b>Descrição</b> (obrigatórias), Tipo, Código pai, Responsável (nome ou e-mail), Início previsto, Término previsto, Modo (Percentual/Produção), Qtd. prevista, Unidade, Peso, Observações.
      Se o tipo ou o pai não forem informados, o sistema deduz pelo código (1 = etapa, 1.1 = subetapa ou tarefa, 1.1.1 = tarefa). Códigos que já existem na obra não são duplicados.</div>
      <div class="go-row"><label class="btn-new go-upload">Escolher planilha<input type="file" id="imp-arq" accept=".xlsx,.xls,.csv"></label><button class="btn-sm" onclick="GO.modeloPlanilha()">⬇️ Baixar modelo</button></div>
      <div id="imp-prev"></div>`,
    rodape: `<button class="btn-cancel" onclick="GO.fecharModal()">Fechar</button>` });
  $("imp-arq").onchange = async e => {
    const f = e.target.files[0]; if (!f) return;
    $("imp-prev").innerHTML = `<div class="go-muted">Lendo…</div>`;
    try {
      const X = await libXlsx();
      const wb = X.read(await f.arrayBuffer(), { type: "array", cellDates: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const brutas = X.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });
      mostrarPrevia(validarImportacao(brutas, { existentes: S.itens, usuarios: S.usuarios }), f.name);
    } catch (err) { $("imp-prev").innerHTML = `<div class="go-erro">${esc(erroMsg(err))}</div>`; }
  };
}
function mostrarPrevia(res, nome) {
  const box = $("imp-prev");
  if (res.erroGeral) { box.innerHTML = `<div class="go-erro">${esc(res.erroGeral)}</div>`; return; }
  const ok = res.linhas.filter(L => !L.erros.length), ruins = res.linhas.filter(L => L.erros.length);
  S._import = res;
  box.innerHTML = `<div class="go-resumo"><div>Arquivo: <b>${esc(nome)}</b></div><div><b>${res.linhas.length}</b> linhas</div><div style="color:var(--vs-green)"><b>${ok.length}</b> válidas</div><div style="color:var(--vs-red)"><b>${ruins.length}</b> com erro</div>
      <div>${["etapa", "subetapa", "tarefa"].map(t => `${ok.filter(L => L.tipo === t).length} ${TIPOS[t].toLowerCase()}(s)`).join(" · ")}</div></div>
    <div class="go-prev"><table class="go-tab"><thead><tr><th>Linha</th><th>Código</th><th>Tipo</th><th>Descrição</th><th>Pai</th><th>Prazo</th><th>Produção</th><th>Resultado</th></tr></thead><tbody>
    ${res.linhas.slice(0, 1500).map(L => `<tr class="${L.erros.length ? "go-tr-erro" : ""}"><td>${L.n}</td><td>${esc(L.cod)}</td><td>${TIPOS[L.tipo] || ""}</td><td>${esc(L.descricao)}</td><td>${esc(L.paiCod)}</td>
      <td>${fmtData(L.inicioPrev)} – ${fmtData(L.fimPrev)}</td><td>${L.tipo === "tarefa" ? (L.modo === "producao" ? `${fmtNum(L.qtdPrevista)} ${esc(L.unidade)}` : "%") : ""}</td>
      <td>${L.erros.length ? `✗ ${esc(L.erros.join(" "))}` : "✓"}${L.avisos.length ? `<div class="go-muted">⚠ ${esc(L.avisos.join(" "))}</div>` : ""}</td></tr>`).join("")}</tbody></table></div>
    ${res.linhas.length > 1500 ? `<div class="go-muted">Mostrando as primeiras 1.500 linhas na prévia.</div>` : ""}
    <div class="go-row">${ruins.length ? `<button class="btn-sm" onclick="GO.relatorioErros()">⬇️ Relatório de erros (${ruins.length})</button>` : ""}<span class="go-flex"></span>
      ${ok.length ? `<button class="btn-save" id="imp-ok">Importar ${ok.length} linha(s) válida(s)</button>` : ""}</div><div id="imp-prog"></div>`;
  if (ok.length) $("imp-ok").onclick = () => executarImportacao(ok);
}
async function relatorioErros() {
  const X = await libXlsx(), ruins = S._import.linhas.filter(L => L.erros.length || L.avisos.length);
  const ws = X.utils.aoa_to_sheet([["Linha", "Código", "Descrição", "Erros", "Avisos"], ...ruins.map(L => [L.n, L.cod, L.descricao, L.erros.join(" "), L.avisos.join(" ")])]);
  ws["!cols"] = [{ wch: 7 }, { wch: 12 }, { wch: 40 }, { wch: 60 }, { wch: 40 }];
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Erros"); X.writeFile(wb, `erros-importacao-${hojeISO()}.xlsx`);
}
async function modeloPlanilha() {
  const X = await libXlsx();
  const ws = X.utils.aoa_to_sheet([
    ["Código", "Tipo", "Descrição", "Código pai", "Responsável", "Início previsto", "Término previsto", "Modo", "Qtd. prevista", "Unidade", "Peso", "Observações"],
    ["1", "Etapa", "EXEMPLO — Fundações", "", "", "", "", "", "", "", "", "Linhas de exemplo: apague antes de importar"],
    ["1.1", "Tarefa", "EXEMPLO — Escavação de sapatas", "1", "", "02/03/2026", "13/03/2026", "Produção", "48,5", "m³", "3500", ""],
    ["1.2", "Tarefa", "EXEMPLO — Concreto das sapatas", "1", "", "16/03/2026", "27/03/2026", "Produção", "22", "m³", "15400", ""],
    ["2", "Etapa", "EXEMPLO — Instalações", "", "", "", "", "", "", "", "", ""],
    ["2.1", "Subetapa", "EXEMPLO — Elétrica", "2", "", "", "", "", "", "", "", ""],
    ["2.1.1", "Tarefa", "EXEMPLO — Quadro de distribuição", "2.1", "", "01/06/2026", "05/06/2026", "Percentual", "", "", "2200", ""]
  ]);
  ws["!cols"] = [8, 9, 36, 10, 16, 14, 14, 11, 12, 8, 8, 36].map(w => ({ wch: w }));
  const wb = X.utils.book_new(); X.utils.book_append_sheet(wb, ws, "Tarefas"); X.writeFile(wb, "modelo-importacao-tarefas.xlsx");
}
async function executarImportacao(validas) {
  const prog = $("imp-prog"), btn = $("imp-ok"); btn.disabled = true;
  const obraId = S.obraId;
  const ids = new Map(S.itens.map(i => [i.codigoKey || codigoKey(i.codigo), i.id]));
  const ordenadas = [...validas].sort((a, b) => ["etapa", "subetapa", "tarefa"].indexOf(a.tipo) - ["etapa", "subetapa", "tarefa"].indexOf(b.tipo) || a.n - b.n);
  for (const L of ordenadas) ids.set(codigoKey(L.cod), doc(COL("itens")).id);
  const ordemBase = new Map();
  const proxOrdem = paiId => { const k = paiId || ""; if (!ordemBase.has(k)) ordemBase.set(k, S.itens.filter(i => (i.paiId || "") === k).reduce((m, i) => Math.max(m, i.ordem || 0), 0)); const v = ordemBase.get(k) + 1000; ordemBase.set(k, v); return v; };
  const linhasPorN = [...validas].sort((a, b) => a.n - b.n);
  const ordemDe = new Map(linhasPorN.map(L => [L.n, proxOrdem(L.paiCod ? ids.get(codigoKey(L.paiCod)) : null)]));
  let feitos = 0; const falhas = [];
  for (let i = 0; i < ordenadas.length; i += LOTE_IMPORT) {
    const lote = ordenadas.slice(i, i + LOTE_IMPORT);
    const b = writeBatch(C.db);
    for (const L of lote) {
      const id = ids.get(codigoKey(L.cod)), key = codigoKey(L.cod);
      b.set(R("codigos", key), { itemId: id, codigo: L.cod });
      const d = { tipo: L.tipo, codigo: L.cod, codigoKey: key, descricao: L.descricao, paiId: L.paiCod ? ids.get(codigoKey(L.paiCod)) : null, ordem: ordemDe.get(L.n),
        responsavelUid: L.respUid || null, responsavelNome: L.respUid ? nomeUsuario(L.respUid) : "", inicioPrev: L.inicioPrev, fimPrev: L.fimPrev, observacoes: L.obs, anexos: [],
        qtdExecutada: 0, criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp(), atualizadoPor: uid(), importado: true };
      if (L.tipo === "tarefa") Object.assign(d, { modo: L.modo, percentual: 0, qtdPrevista: L.modo === "producao" ? L.qtdPrevista : null, unidade: L.modo === "producao" ? L.unidade : "", peso: L.pesoN, inicioReal: null, fimReal: null });
      b.set(R("itens", id), d);
    }
    try { await b.commit(); feitos += lote.length; }
    catch (e) { falhas.push(...lote.map(L => `${L.cod}: ${erroMsg(e)}`)); }
    prog.innerHTML = `<div class="go-bar"><span style="width:${(i + lote.length) / ordenadas.length * 100}%"></span></div><div class="go-muted">${feitos} de ${ordenadas.length} gravados…</div>`;
  }
  const b = writeBatch(C.db);
  regHist(b, obraId, { alvoTipo: "obra", alvoId: obraId, acao: "importar", resumo: `Importação de planilha: ${feitos} item(ns) incluído(s)${falhas.length ? `, ${falhas.length} falha(s)` : ""}` });
  await b.commit().catch(() => {});
  prog.innerHTML = falhas.length ? `<div class="go-erro">${feitos} gravados; ${falhas.length} falharam (o código pode ter sido criado por outra pessoa durante a importação):<br>${esc(falhas.slice(0, 20).join(" · "))}</div>`
    : `<div class="go-ok-box">✔ ${feitos} item(ns) importado(s).</div>`;
  toast(`${feitos} item(ns) importado(s)`);
}

// ---------------------------------------------------------------------
// Obra de demonstração (dados fictícios, identificados)
// ---------------------------------------------------------------------
export function gerarEapDemo(hoje = hojeISO()) {
  const ini = somaDias(hoje, -75);
  const E = [
    ["1", "Serviços preliminares", [["Mobilização e canteiro de obras", null, null, 5, 0, 8000], ["Locação da obra", 420, "m²", 4, 5, 2100]]],
    ["2", "Fundações", [["Escavação de sapatas", 64, "m³", 8, 9, 5800], ["Armação das sapatas", 2350, "kg", 7, 14, 21000], ["Concreto das sapatas", 31, "m³", 6, 19, 18500]]],
    ["3", "Estrutura", null, [
      ["3.1", "Pavimento térreo", [["Formas de pilares e vigas", 310, "m²", 10, 26, 24000], ["Armação de pilares e vigas", 3900, "kg", 10, 28, 35000], ["Concretagem da laje", 46, "m³", 3, 37, 27500]]],
      ["3.2", "Pavimento superior", [["Formas de pilares e vigas", 290, "m²", 10, 41, 22500], ["Armação de pilares e vigas", 3600, "kg", 10, 43, 32000], ["Concretagem da laje", 42, "m³", 3, 53, 25000]]]]],
    ["4", "Alvenaria", [["Alvenaria de vedação — térreo", 560, "m²", 14, 40, 30000], ["Alvenaria de vedação — superior", 520, "m²", 14, 56, 28000]]],
    ["5", "Instalações", null, [
      ["5.1", "Elétrica", [["Eletrodutos embutidos", 1450, "m", 12, 50, 9500], ["Quadros de distribuição", null, null, 4, 85, 6200]]],
      ["5.2", "Hidrossanitária", [["Tubulação de água fria", 380, "m", 10, 52, 8800], ["Rede de esgoto", 210, "m", 10, 58, 9200]]]]],
    ["6", "Revestimentos e acabamentos", [["Chapisco interno", 1900, "m²", 8, 66, 7600], ["Reboco interno", 1900, "m²", 16, 72, 38000], ["Contrapiso", 720, "m²", 8, 82, 21600], ["Porcelanato", 640, "m²", 14, 92, 57600], ["Pintura interna", 2100, "m²", 14, 104, 31500]]],
    ["7", "Cobertura", [["Estrutura metálica da cobertura", 4200, "kg", 8, 60, 46000], ["Telhamento", 380, "m²", 5, 69, 19000]]],
    ["8", "Limpeza e entrega", [["Limpeza final e entrega", null, null, 5, 118, 4500]]]
  ];
  const itens = [];
  const tarefa = (cod, pai, [desc, q, un, dur, off, peso]) => itens.push({ tipo: "tarefa", codigo: cod, descricao: desc, pai, modo: q ? "producao" : "manual", qtdPrevista: q, unidade: un || "",
    inicioPrev: somaDias(ini, off), fimPrev: somaDias(ini, off + dur - 1), peso });
  for (const [c, d, tarefas, subs] of E) {
    itens.push({ tipo: "etapa", codigo: c, descricao: d, pai: null });
    (tarefas || []).forEach((t, i) => tarefa(`${c}.${i + 1}`, c, t));
    (subs || []).forEach(([sc, sd, st]) => { itens.push({ tipo: "subetapa", codigo: sc, descricao: sd, pai: c }); st.forEach((t, i) => tarefa(`${sc}.${i + 1}`, sc, t)); });
  }
  return itens;
}
/** Distribui uma produção fictícia em RDOs semanais. Ritmo menor em algumas tarefas gera atrasos; uma tarefa passa do previsto. */
export function gerarRdosDemo(tarefas, hoje = hojeISO()) {
  const datas = []; for (let d = -70; d <= -2; d += 4) datas.push(somaDias(hoje, d));
  const ritmo = cod => ({ "3.2.2": 0.55, "4.2": 0.5, "5.1.1": 0.7, "6.1": 0.6 }[cod] || (cod === "2.1" ? 1.08 : 1));
  const feito = new Map();
  const rdos = datas.map(data => ({ data, atividades: [] }));
  for (const t of tarefas) {
    for (const r of rdos) {
      if (r.data < t.inicioPrev) continue;
      const dur = (new Date(t.fimPrev) - new Date(t.inicioPrev)) / 864e5 + 1;
      const frac = Math.min(1, ((new Date(r.data) - new Date(t.inicioPrev)) / 864e5 + 1) / dur) * ritmo(t.codigo);
      if (t.modo === "producao") {
        const alvo = arred(Math.round(t.qtdPrevista * Math.min(frac, ritmo(t.codigo) > 1 ? frac : 1) * 10) / 10);
        const q = arred(alvo - (feito.get(t.codigo) || 0));
        if (q > 0) { r.atividades.push({ codigo: t.codigo, qtd: q }); feito.set(t.codigo, alvo); }
      } else {
        const p = Math.round(Math.min(100, frac * 100));
        if (p > (feito.get(t.codigo) || 0)) { r.atividades.push({ codigo: t.codigo, percentual: p }); feito.set(t.codigo, p); }
      }
    }
  }
  return rdos.filter(r => r.atividades.length);
}
async function criarDemo() {
  if (S.obras.some(o => o.id === "demo-residencial-exemplo")) { location.hash = "#acompanhamento/demo-residencial-exemplo"; return; }
  if ((await perguntar("Criar obra de demonstração", "Cria a obra <b>[DEMO] Residencial Exemplo</b> com etapas, tarefas, RDOs aprovados, um RDO pendente e um rascunho — tudo com <b>dados fictícios</b> e identificado como demonstração. Só administradores veem a obra até alguém ser incluído na equipe. Leva cerca de um minuto.", { botao: "Criar" })) === null) return;
  const id = "demo-residencial-exemplo", ctx = { db: C.db, uid: uid(), nome: meuNome() }, hoje = hojeISO();
  const Rf = (...p) => doc(C.db, "obras", id, ...p);
  toast("Criando obra de demonstração…");
  try {
    let b = writeBatch(C.db);
    b.set(Rf(), { nome: "[DEMO] Residencial Exemplo", descricao: "Obra fictícia para demonstração do módulo. Nenhum dado é real.", endereco: "Rua Fictícia, 100 · Bairro Exemplo · Recife/PE",
      responsavelUid: uid(), responsavelNome: meuNome(), inicioPrev: somaDias(hoje, -75), fimPrev: somaDias(hoje, 50), situacao: "em_andamento", demo: true,
      membros: {}, membrosUids: [], blocos: [], criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp() });
    await b.commit();
    const eap = gerarEapDemo(hoje), ids = new Map(eap.map(e => [e.codigo, doc(collection(C.db, "obras", id, "itens")).id]));
    const ordem = new Map();
    for (let i = 0; i < eap.length; i += LOTE_IMPORT) {
      b = writeBatch(C.db);
      for (const e of eap.slice(i, i + LOTE_IMPORT)) {
        const k = e.pai || ""; ordem.set(k, (ordem.get(k) || 0) + 1000);
        b.set(Rf("codigos", codigoKey(e.codigo)), { itemId: ids.get(e.codigo), codigo: e.codigo });
        const d = { tipo: e.tipo, codigo: e.codigo, codigoKey: codigoKey(e.codigo), descricao: e.descricao, paiId: e.pai ? ids.get(e.pai) : null, ordem: ordem.get(k),
          responsavelUid: null, responsavelNome: "", inicioPrev: e.inicioPrev || null, fimPrev: e.fimPrev || null, observacoes: "", anexos: [], qtdExecutada: 0,
          criadoEm: serverTimestamp(), criadoPor: uid(), atualizadoEm: serverTimestamp(), atualizadoPor: uid(), demo: true };
        if (e.tipo === "tarefa") Object.assign(d, { modo: e.modo, percentual: 0, qtdPrevista: e.qtdPrevista, unidade: e.unidade, peso: e.peso, inicioReal: null, fimReal: null });
        b.set(Rf("itens", ids.get(e.codigo)), d);
      }
      await b.commit();
    }
    const rdos = gerarRdosDemo(eap.filter(e => e.tipo === "tarefa"), hoje);
    const mo = [{ funcao: "Mestre de obras", empresa: "Própria", qtd: 1 }, { funcao: "Pedreiro", empresa: "Própria", qtd: 6 }, { funcao: "Servente", empresa: "Própria", qtd: 8 }, { funcao: "Armador", empresa: "Empreiteira Exemplo", qtd: 3 }];
    const eq = [{ nome: "Betoneira 400 L", situacao: "Operando", qtd: 1 }, { nome: "Vibrador de imersão", situacao: "Operando", qtd: 2 }];
    for (const [i, r] of rdos.entries()) {
      const pendente = i === rdos.length - 1;
      b = writeBatch(C.db);
      b.set(Rf("rdos", r.data), { data: r.data, status: "enviado", versao: 0, clima: { manha: i % 5 === 3 ? "Chuva fraca" : "Bom", tarde: i % 5 === 3 ? "Chuva forte" : "Nublado", noite: "", condicao: i % 5 === 3 ? "parcial" : "praticavel", impacto: i % 5 === 3 ? "Chuva à tarde suspendeu serviços externos (fictício)" : "", horasParadas: i % 5 === 3 ? 3 : 0 },
        maoDeObra: mo, equipamentos: eq, atividades: r.atividades.map(a => { const e = eap.find(x => x.codigo === a.codigo); return { itemId: ids.get(a.codigo), codigo: a.codigo, descricao: e.descricao, unidade: e.unidade || "", qtd: a.qtd ?? null, percentual: a.percentual ?? null, obs: "" }; }),
        ocorrencias: i === 2 ? "Entrega de aço com 1 dia de atraso (fictício)." : "", impedimentos: "", observacoes: "Registro fictício de demonstração.", fotos: [], lancamentos: [],
        autorUid: uid(), autorNome: meuNome(), criadoEm: serverTimestamp(), atualizadoEm: serverTimestamp(), eventos: [evento("criar"), evento("enviar")], demo: true });
      await b.commit();
      if (!pendente) await aprovarRdoTx(id, r.data, ctx);
    }
    b = writeBatch(C.db);
    b.set(Rf("rdos", hoje), { data: hoje, status: "rascunho", versao: 0, clima: { manha: "Bom", tarde: "", noite: "", condicao: "praticavel", impacto: "", horasParadas: 0 }, maoDeObra: mo, equipamentos: eq,
      atividades: [], ocorrencias: "", impedimentos: "", observacoes: "Rascunho fictício de demonstração.", fotos: [], lancamentos: [], autorUid: uid(), autorNome: meuNome(),
      criadoEm: serverTimestamp(), atualizadoEm: serverTimestamp(), eventos: [evento("criar")], demo: true });
    regHist(b, id, { alvoTipo: "obra", alvoId: id, acao: "criar", resumo: "Obra de demonstração criada (dados fictícios)" });
    await b.commit();
    toast("Obra de demonstração criada");
    location.hash = `#acompanhamento/${id}`;
  } catch (e) { console.error(e); toast("Falha ao criar a demonstração: " + erroMsg(e)); }
}
async function apagarDemo() {
  const o = obraAtual(); if (!o || !o.demo) return;
  if ((await perguntar("Apagar obra de demonstração", `Apagar <b>${esc(o.nome)}</b> e todos os dados fictícios dela?`, { botao: "Apagar", perigo: true })) === null) return;
  const id = o.id;
  try {
    for (const sub of ["itens", "codigos", "lancamentos", "rdos", "arquivos", "historico", "serie"]) {
      const snap = await getDocs(collection(C.db, "obras", id, sub));
      for (let i = 0; i < snap.docs.length; i += 300) { const b = writeBatch(C.db); snap.docs.slice(i, i + 300).forEach(d => b.delete(d.ref)); await b.commit(); }
    }
    const b = writeBatch(C.db); b.delete(doc(C.db, "obras", id)); await b.commit();
    toast("Demonstração apagada"); location.hash = "#acompanhamento";
  } catch (e) { toast(erroMsg(e)); }
}

// ---------------------------------------------------------------------
// Funções chamadas pelos botões (onclick)
// ---------------------------------------------------------------------
window.GO = {
  fecharModal, novaObra, migrarLegadas, criarDemo, apagarDemo,
  abrirItem, excluirItem, mover, ajustarProducao, anexarItem, removerAnexoItem, carregarHistItem, baixarArquivo,
  alternar: id => { S.ui.abertos.has(id) ? S.ui.abertos.delete(id) : S.ui.abertos.add(id); renderLinhas(); },
  expandir: tudo => { S.ui.abertos = tudo ? new Set(S.itens.filter(i => i.tipo !== "tarefa").map(i => i.id)) : new Set(); renderLinhas(); },
  limparFiltros: () => { Object.assign(S.ui, { busca: "", situacao: "", resp: "", etapa: "", prazo: "", soEtapas: false, ocultarConcluidas: false, limite: 600 }); renderTarefas(); },
  filtrarEtapa: id => { S.ui.etapa = id; },
  maisLinhas: () => { S.ui.limite += 600; renderLinhas(); },
  importar, modeloPlanilha, relatorioErros, exportarXlsx, exportarPdf, conferir,
  novoRdo, filtroRdo, salvarRdo, aprovarRdo, reabrirRdo, cancelarRdo, excluirRdo, pdfRdo,
  rdoIncluir, rdoRemover, rdoVincular, rdoFotos, rdoRemoverFoto
};

// ---------------------------------------------------------------------
// Estilos (seguem as cores e componentes do CRM)
// ---------------------------------------------------------------------
function injetarCss() {
  if ($("go-css")) return;
  const s = document.createElement("style"); s.id = "go-css";
  s.textContent = `
#obras-app .go-intro{display:flex;align-items:flex-end;justify-content:space-between;gap:12px;flex-wrap:wrap}
.go-acoes{display:flex;gap:8px;flex-wrap:wrap}
.go-demo{font-size:10.5px;font-weight:800;letter-spacing:.04em;padding:3px 8px;border-radius:20px;background:#FCE8F3;color:#9D174D;margin-right:6px;vertical-align:middle}
.go-card-demo{border-top-color:#D4537E}
.go-sit{font-size:11px;font-weight:700;padding:3px 10px;border-radius:20px;background:#EEF2F7;color:#475569}
.go-sit.s-em_andamento{background:#FFF4DB;color:#8A5A00}.go-sit.s-concluida{background:#E1F5EE;color:#11684D}.go-sit.s-paralisada{background:#FDE8E8;color:#9B1C1C}
.go-papel{font-size:11.5px;color:#64748B}
.go-carregando{padding:40px;text-align:center;color:#888}
.go-muted{color:#8A94A6;font-size:12px}
.go-dica{font-size:11.5px;color:#7A8597;line-height:1.45;margin-top:6px}
.go-erro{color:#B42318;font-size:12.5px;line-height:1.5;min-height:1px}
.go-flex{flex:1}
.go-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin:10px 0}
.go-row-end{display:flex;justify-content:flex-end;margin-top:10px}
.go-grid2{display:grid;grid-template-columns:1fr 1fr;gap:10px 12px}.go-grid3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px 12px}.go-span2{grid-column:span 2}
.go-fg{gap:12px}
.go-box{border:1px solid #E3E8F0;border-radius:10px;padding:10px 12px;background:#FAFBFD}
.go-box-t{font-size:12px;font-weight:800;color:var(--vs-blue);text-transform:uppercase;letter-spacing:.04em;margin-bottom:8px;cursor:default}
details.go-box summary{cursor:pointer}
.go-sec{background:#fff;border-radius:12px;padding:14px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:12px;min-width:0}
.go-sec h3{font-size:13.5px;font-weight:800;color:var(--vs-dark);margin-bottom:10px}
.go-sec .fg input,.go-sec input,.go-sec select,.go-sec textarea,.go-toolbar input,.go-toolbar select{font-family:inherit;font-size:13px;border-radius:8px;border:1px solid #ddd;padding:7px 9px;background:#fff;color:#1a1a1a;outline:none;width:100%}
.go-sec textarea{min-height:64px;resize:vertical}
.go-sec input:disabled,.go-sec select:disabled,.go-sec textarea:disabled,.go-fg input:disabled,.go-fg select:disabled,.go-fg textarea:disabled{background:#F4F6F9;color:#555}
.go-cols{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.go-perigo{border:1px solid #FECACA}
.go-membros{display:flex;flex-direction:column;gap:6px;max-height:380px;overflow:auto}
.go-membro{display:flex;justify-content:space-between;align-items:center;gap:10px;padding:7px 9px;border:1px solid #EEF1F5;border-radius:8px}
.go-membro b{display:block;font-size:13px}.go-membro span{font-size:11.5px;color:#8A94A6}.go-membro select{width:150px!important;flex-shrink:0}
.go-toolbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;background:#fff;border-radius:12px;padding:10px 12px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:8px}
.go-toolbar select{width:auto;max-width:220px}.go-busca{flex:1 1 240px;width:auto!important}
.go-toolbar2{padding:8px 12px}
.go-chk{display:flex;align-items:center;gap:5px;font-size:12.5px;color:#444;white-space:nowrap}.go-chk input{width:auto!important}
.go-tabela{background:#fff;border-radius:12px;box-shadow:0 1px 4px rgba(0,0,0,.08);overflow:hidden}
.go-tr{display:grid;grid-template-columns:minmax(110px,.9fr) minmax(220px,2.6fr) 1fr 150px 150px 130px 150px 150px 92px;gap:8px;align-items:center;padding:7px 12px;border-bottom:1px solid #F0F2F6;font-size:12.5px;cursor:pointer}
.go-tr:hover{background:#F7FAFF}
.go-th{background:#F4F7FB;font-size:11px;font-weight:700;color:#64748B;text-transform:uppercase;letter-spacing:.03em;cursor:default;position:sticky;top:0}
.go-tr.t-etapa{background:#F0F5FD;font-weight:700}.go-tr.t-subetapa{background:#F8FAFE;font-weight:600}
.go-cod{display:flex;align-items:center;gap:4px;font-variant-numeric:tabular-nums;white-space:nowrap}
.go-caret{border:0;background:none;cursor:pointer;width:18px;color:var(--vs-blue);font-size:12px;padding:0}.go-caret-v{display:inline-block;width:18px}
.go-tipo{font-size:9.5px;font-weight:700;color:#94A3B8;text-transform:uppercase;margin-right:6px;letter-spacing:.03em}
.go-desc{min-width:0;overflow:hidden;text-overflow:ellipsis}
.go-dt{font-size:11.5px;color:#555;white-space:nowrap;font-variant-numeric:tabular-nums}
.go-av{display:flex;align-items:center;gap:6px}
.go-bar{flex:1;height:7px;background:#E8EDF4;border-radius:4px;overflow:hidden;min-width:40px}
.go-bar span{display:block;height:100%;background:var(--vs-blue-cyan);border-radius:4px}
.go-bar.acima span{background:#B45309}
.go-pct{font-size:11.5px;font-weight:700;min-width:44px;text-align:right;font-variant-numeric:tabular-nums}
.go-prod{font-size:11.5px;font-variant-numeric:tabular-nums}
.go-acima{color:#B45309;font-weight:700}
.go-alerta{color:#B45309;font-weight:700;font-size:11.5px}
.go-alerta-box{background:#FFF7E6;border:1px solid #FCD9A0;color:#7A4B00;border-radius:8px;padding:8px 10px;font-size:12.5px;margin:8px 0;line-height:1.45}
.go-ok-box{background:#E9F7F1;border:1px solid #A7DFC8;color:#11684D;border-radius:8px;padding:8px 10px;font-size:12.5px;margin:8px 0;line-height:1.45}
.go-chip{display:inline-block;font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:20px;margin:1px 3px 1px 0;white-space:nowrap}
.c-nao_iniciada{background:#EEF2F7;color:#475569}.c-em_andamento{background:#E6F0FB;color:#0B5AA6}.c-concluida{background:#E1F5EE;color:#11684D}
.c-atraso{background:#FDE8E8;color:#9B1C1C}.c-rdo{background:#E6F0FB;color:#0B5AA6}.c-estorno{background:#FDE8E8;color:#9B1C1C}.c-ajuste{background:#FFF4DB;color:#8A5A00}
.r-rascunho{background:#EEF2F7;color:#475569}.r-enviado{background:#FFF4DB;color:#8A5A00}.r-aprovado{background:#E1F5EE;color:#11684D}.r-cancelado{background:#FDE8E8;color:#9B1C1C}
.go-acoes-l{display:flex;justify-content:flex-end;gap:0}
.go-vazio-l{padding:28px;text-align:center;color:#888;font-size:13px}
.go-rodape-lista{display:flex;gap:12px;align-items:center;justify-content:space-between;flex-wrap:wrap;font-size:11.5px;color:#8A94A6;padding:8px 4px}
.go-radios{display:flex;gap:16px;flex-wrap:wrap;font-size:13px;margin-bottom:8px}.go-radios input{width:auto!important;margin-right:4px}
.go-av-atual{display:flex;align-items:center;gap:8px;font-size:12.5px;padding-top:6px;flex-wrap:wrap}.go-av-atual .go-bar{max-width:120px}
.go-resumo{display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12.5px}
.go-upload{position:relative;overflow:hidden;display:inline-flex;margin-top:6px;cursor:pointer}.go-upload input{position:absolute;inset:0;opacity:0;cursor:pointer}
.go-anexos{display:flex;flex-direction:column;gap:6px}
.go-anexo{display:flex;align-items:center;gap:8px;border:1px solid #EEF1F5;border-radius:8px;padding:5px 8px;background:#fff}
.go-anexo img{width:40px;height:32px;object-fit:cover;border-radius:4px}.go-anexo-ic{font-size:18px;width:40px;text-align:center}
.go-anexo-n{flex:1;min-width:0;font-size:12.5px}.go-anexo-n a{color:var(--vs-blue-cyan);text-decoration:none;word-break:break-all}.go-anexo-n span{display:block;font-size:11px;color:#8A94A6}
.go-hist{display:flex;flex-direction:column;gap:6px;max-height:320px;overflow:auto}
.go-hist-i{border-left:3px solid var(--vs-blue-pale);padding:3px 8px;font-size:12px}
.go-hist-h{color:#334155}.go-hist ul{margin:3px 0 0 16px;color:#555}.go-hist s{color:#999}
.go-sub{font-size:11px;font-weight:800;color:#64748B;text-transform:uppercase;margin:8px 0 4px}
.go-lanc{display:flex;flex-direction:column;gap:4px;font-size:12px}.go-lanc>div{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.go-kpis{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:12px}
.go-kpi{background:#fff;border-radius:12px;padding:12px 14px;box-shadow:0 1px 4px rgba(0,0,0,.08);border-top:3px solid var(--vs-blue-cyan)}
.go-kpi span{display:block;font-size:11.5px;color:#64748B;font-weight:600}.go-kpi b{font-size:26px;font-weight:800;color:var(--vs-dark);font-variant-numeric:tabular-nums}
.go-kpi-alerta{border-top-color:var(--vs-red)}
.go-indic{display:flex;gap:14px;align-items:center}
.go-indic-v{font-size:30px;font-weight:800;color:var(--vs-blue);min-width:112px;text-align:center;font-variant-numeric:tabular-nums}
.go-indic-simples{color:#64748B}
.go-etapa{padding:7px 4px;border-bottom:1px solid #F0F2F6;cursor:pointer}.go-etapa:hover{background:#F7FAFF}
.go-etapa-t{font-size:12.5px}.go-etapa-b{display:flex;align-items:center;gap:8px;margin:3px 0}
.go-tab{width:100%;border-collapse:collapse;font-size:12px}.go-tab th{text-align:left;background:#F4F7FB;color:#64748B;font-weight:700;padding:5px 7px}.go-prev .go-tab th{position:sticky;top:0;z-index:1}.go-tab td{padding:5px 7px;border-bottom:1px solid #F0F2F6;vertical-align:top}
.go-tr-erro td{background:#FFF5F5}
.go-prev{max-height:46vh;overflow:auto;border:1px solid #EEF1F5;border-radius:8px;margin-top:8px}
.go-mini{display:flex;flex-direction:column}.go-mini>div{display:flex;justify-content:space-between;gap:10px;padding:6px 4px;border-bottom:1px solid #F0F2F6;font-size:12.5px;cursor:pointer}.go-mini>div:hover{background:#F7FAFF}
.go-mini>div>span:last-child{color:#64748B;white-space:nowrap}
.go-graf{position:relative;max-width:900px}.go-graf svg{width:100%;height:auto;display:block}
.go-ax{font-size:10px;fill:#8A94A6}.go-lab{font-size:10.5px;fill:#334155;font-weight:700}
.go-leg{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:#334155;margin-bottom:4px}.go-leg span{display:flex;align-items:center;gap:5px;white-space:nowrap}
.go-tip{position:absolute;top:30px;background:#fff;border:1px solid #DCE3EC;border-radius:8px;padding:6px 9px;font-size:11.5px;box-shadow:0 4px 14px rgba(0,0,0,.1);pointer-events:none;line-height:1.5}
.go-novo-rdo{display:flex;align-items:flex-end;gap:8px}.go-novo-rdo input{width:auto}.go-novo-rdo .fl{align-self:center}
.go-rdo-lista{display:flex;flex-direction:column;gap:8px}
.go-rdo-card{display:flex;align-items:center;gap:14px;background:#fff;border-radius:12px;padding:12px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);text-decoration:none;color:inherit}
.go-rdo-card:hover{box-shadow:0 6px 16px rgba(0,0,0,.1)}
.go-rdo-data{font-size:16px;font-weight:800;color:var(--vs-blue);min-width:100px}.go-rdo-data span{display:block;font-size:11px;color:#8A94A6;font-weight:600}
.go-rdo-info{flex:1;display:flex;flex-direction:column;gap:3px}
.go-rdo-head{display:flex;align-items:center;gap:14px;flex-wrap:wrap;background:#fff;border-radius:12px;padding:10px 16px;box-shadow:0 1px 4px rgba(0,0,0,.08);margin-bottom:10px}
.go-rdo-titulo{font-size:15px;font-weight:800}
.go-linha3{display:grid;grid-template-columns:1.4fr 1fr 70px 30px;gap:6px;margin-bottom:6px;align-items:center}
.go-qtd{text-align:right}
.go-ativ{border:1px solid #E8EDF4;border-radius:10px;padding:8px 10px;margin-bottom:8px}
.go-ativ-t{display:flex;justify-content:space-between;gap:8px;align-items:center;font-size:13px}.go-ativ-t input{flex:1}
.go-ativ-c{display:flex;gap:10px;margin-top:6px;flex-wrap:wrap}.go-ativ-c .fgf{min-width:180px}
.go-ativ-i{font-size:11.5px;color:#64748B;margin-top:5px}
.go-picker{position:relative;margin:8px 0}
.go-sug{display:none;position:absolute;left:0;right:0;top:100%;background:#fff;border:1px solid #DCE3EC;border-radius:8px;box-shadow:0 8px 22px rgba(0,0,0,.12);z-index:20;max-height:300px;overflow:auto}
.go-sug-i{padding:7px 10px;font-size:12.5px;cursor:pointer;border-bottom:1px solid #F3F5F8}.go-sug-i:hover{background:#F0F5FD}
.go-fotos{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px}
.go-foto{display:flex;flex-direction:column;gap:4px;position:relative}.go-foto img{width:100%;aspect-ratio:4/3;object-fit:cover;border-radius:8px;cursor:pointer}
.go-foto input{font-size:11.5px!important;padding:4px 6px!important}.go-foto .card-btn{position:absolute;top:2px;right:2px;background:#fff}
.go-foto-arq{aspect-ratio:4/3;border:1px dashed #CBD5E1;border-radius:8px;display:flex;align-items:center;justify-content:center;text-align:center;font-size:11px;cursor:pointer;padding:6px;word-break:break-all}
.go-barra-acoes{position:sticky;bottom:0;display:flex;gap:8px;justify-content:flex-end;flex-wrap:wrap;background:rgba(240,244,248,.95);backdrop-filter:blur(4px);padding:10px 0;z-index:5}
@media(max-width:1100px){.go-tr{grid-template-columns:minmax(90px,.8fr) minmax(180px,2.4fr) 130px 130px 140px 92px}.go-tr>.go-resp,.go-tr>.go-dt:nth-child(5),.go-tr>.go-prod{display:none}}
@media(max-width:820px){
  .go-cols,.go-grid2,.go-grid3{grid-template-columns:1fr}.go-span2{grid-column:auto}
  .go-kpis{grid-template-columns:repeat(2,1fr)}.go-kpis .go-kpi:first-child{grid-column:span 2}
  .go-th{display:none}
  .go-tr{grid-template-columns:1fr auto;grid-template-areas:"cod sit" "desc desc" "av av" "dt acoes";row-gap:4px}
  .go-tr>.go-cod{grid-area:cod}.go-tr>.go-desc{grid-area:desc;white-space:normal}.go-tr>.go-sitc{grid-area:sit;text-align:right}
  .go-tr>.go-av{grid-area:av}.go-tr>.go-dt:nth-child(4){grid-area:dt}.go-tr>.go-acoes-l{grid-area:acoes}
  .go-toolbar select{max-width:none;flex:1 1 45%}
  .go-linha3{grid-template-columns:1fr 70px 30px}.go-linha3>:nth-child(2){display:none}
  .go-indic{flex-direction:column;align-items:flex-start}
  .go-rdo-card .obra-go{display:none}
}`;
  document.head.appendChild(s);
}
