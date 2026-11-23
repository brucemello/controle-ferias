// =========================================================================
// FÉRIAS INCENTIVADAS — campanhas da empresa com dias abonados
// -------------------------------------------------------------------------
// O gestor cadastra a campanha e suas opções (férias legais + dias abonados
// + retorno). O dev registra interesse em uma opção; o gestor confirma ou
// recusa. Ao confirmar, as férias e os abonos entram na planilha e na agenda.
// Abas: "Férias Incentivadas" (uma linha por opção) e "Adesões Incentivadas".
// =========================================================================

const INC = {
  abaCampanhas: "Férias Incentivadas",
  abaAdesoes: "Adesões Incentivadas",
  cabCampanhas: ["Campanha ID", "Campanha", "Opção", "Início férias", "Fim férias", "Dias abonados", "Retorno", "Prazo de adesão", "Ativa"],
  cabAdesoes: ["ID", "Data", "Colaborador", "Campanha ID", "Opção", "Status", "Registrado por", "Observação"]
};

function abaInc_(nome, cab) {
  const ss = planilha_();
  let aba = ss.getSheetByName(nome);
  if (!aba) {
    aba = ss.insertSheet(nome);
    aba.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight("bold").setBackground("#1f4e78").setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }
  return aba;
}

function dadosAba_(aba, nc) {
  return aba.getLastRow() > 1 ? aba.getRange(2, 1, aba.getLastRow() - 1, nc).getValues() : [];
}

function ehGestorAtual_() {
  const g = emailLider_();
  if (!g) return true;
  let u = "";
  try { u = Session.getActiveUser().getEmail(); } catch (e) {}
  return String(u).toLowerCase() === String(g).toLowerCase();
}

function listaDatas_(txt) {
  return String(txt || "").split(/[;,\n]/).map(s => parseData_(s.trim())).filter(d => d).sort((a, b) => a - b);
}

// ---------- Leitura para o dashboard ----------
function obterIncentivadas() {
  vincularPlanilha_();
  const abaC = abaInc_(INC.abaCampanhas, INC.cabCampanhas);
  const abaA = abaInc_(INC.abaAdesoes, INC.cabAdesoes);
  const hoje = new Date(); hoje.setHours(0, 0, 0, 0);

  const adesoes = dadosAba_(abaA, 8).map((r, i) => ({
    linha: i + 2, id: String(r[0]), data: r[1] instanceof Date ? fmt_(r[1]) : String(r[1]), nome: String(r[2]).toUpperCase(),
    campanha: String(r[3]), opcao: String(r[4]), status: String(r[5]), obs: String(r[7] || "")
  })).filter(a => a.id);

  // Quem já está fora (férias/folgas) para mostrar ocupação no período de cada opção
  let periodos = [];
  try {
    const d = obterDadosDashboard();
    periodos = d.periodos.concat(d.folgas).filter(p => !/INCENTIVADA/i.test(p.situacao || p.beneficio || ""));
  } catch (e) {}

  const campanhas = {};
  dadosAba_(abaC, 9).forEach(r => {
    const id = String(r[0]); if (!id) return;
    const ini = parseData_(r[3]), fim = parseData_(r[4]); if (!ini || !fim) return;
    const ab = listaDatas_(r[5]);
    const ret = parseData_(r[6]) || new Date(fim.getFullYear(), fim.getMonth(), fim.getDate() + 1);
    const prazo = parseData_(r[7]);
    if (!campanhas[id]) campanhas[id] = { id: id, nome: String(r[1]), prazo: prazo ? iso_(prazo) : "", ativa: String(r[8]).toUpperCase() !== "NÃO" && String(r[8]).toUpperCase() !== "NAO",
      encerrada: !!prazo && prazo < hoje, opcoes: [] };
    const ultimoDia = new Date(ret.getFullYear(), ret.getMonth(), ret.getDate() - 1);
    const fora = periodos.filter(p => parseData_(p.inicio) <= ultimoDia && parseData_(p.fim) >= ini).map(p => p.nome);
    campanhas[id].opcoes.push({
      opcao: String(r[2]), inicio: iso_(ini), fim: iso_(fim), diasFerias: diasEntre_(ini, fim),
      abonados: ab.map(iso_), retorno: iso_(ret), totalDescanso: diasEntre_(ini, ultimoDia),
      foraNoPeriodo: fora.filter((n, i) => fora.indexOf(n) === i),
      adesoes: adesoes.filter(a => a.campanha === id && a.opcao === String(r[2]) && a.status !== "Cancelado")
    });
  });

  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  let meuNome = "";
  try {
    const c = lerContatos_();
    Object.keys(c).forEach(n => { if (String(c[n]).toLowerCase() === String(usuario).toLowerCase()) meuNome = n; });
  } catch (e) {}

  let time = [];
  try { time = buscarNomesTime().map(n => String(n).toUpperCase()); } catch (e) {}

  return {
    campanhas: Object.keys(campanhas).map(k => campanhas[k]).sort((a, b) => (b.prazo || "").localeCompare(a.prazo || "")),
    ehGestor: ehGestorAtual_(), usuario: usuario, meuNome: meuNome, time: time.sort()
  };
}

// ---------- Gestor: campanhas ----------
// c: { id?, nome, prazo: iso, opcoes: [{ opcao, inicio, fim, abonados: [iso], retorno }] }
function salvarCampanhaIncentivada(c) {
  if (!ehGestorAtual_()) throw new Error("Só o gestor pode cadastrar campanhas.");
  const nome = String(c.nome || "").trim();
  if (!nome) throw new Error("Informe o nome da campanha.");
  if (!c.opcoes || !c.opcoes.length) throw new Error("Cadastre pelo menos uma opção.");
  const id = c.id || "INC-" + Utilities.formatDate(new Date(), tz_(), "yyyyMMddHHmmss");
  const prazo = parseData_(c.prazo);

  const linhas = c.opcoes.map((o, i) => {
    const ini = parseData_(o.inicio), fim = parseData_(o.fim);
    if (!ini || !fim || fim < ini) throw new Error("Opção " + (i + 1) + ": confira o início e o fim das férias.");
    const ab = (o.abonados || []).map(parseData_).filter(d => d).sort((a, b) => a - b);
    let ret = parseData_(o.retorno);
    if (!ret) { const base = ab.length ? ab[ab.length - 1] : fim; ret = proximoDiaUtil_(new Date(base.getFullYear(), base.getMonth(), base.getDate() + 1)); }
    return [id, nome, String(o.opcao || ("Opção " + (i + 1))).trim(), ini, fim, ab.map(fmt_).join("; "), ret, prazo || "", "Sim"];
  });

  const aba = abaInc_(INC.abaCampanhas, INC.cabCampanhas);
  // remove linhas antigas da mesma campanha
  dadosAba_(aba, 1).map((r, i) => String(r[0]) === id ? i + 2 : 0).filter(x => x).reverse().forEach(l => aba.deleteRow(l));
  const ini = Math.max(aba.getLastRow(), 1) + 1;
  aba.getRange(ini, 1, linhas.length, 9).setValues(linhas);
  aba.getRange(ini, 4, linhas.length, 2).setNumberFormat("dd/mm/yyyy");
  aba.getRange(ini, 7, linhas.length, 2).setNumberFormat("dd/mm/yyyy");
  return "Campanha \"" + nome + "\" salva com " + linhas.length + " opção(ões).";
}

function proximoDiaUtil_(d) {
  while (d.getDay() === 0 || d.getDay() === 6) d = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
  return d;
}

function encerrarCampanhaIncentivada(id, reabrir) {
  if (!ehGestorAtual_()) throw new Error("Só o gestor pode encerrar campanhas.");
  const aba = abaInc_(INC.abaCampanhas, INC.cabCampanhas);
  dadosAba_(aba, 1).forEach((r, i) => { if (String(r[0]) === id) aba.getRange(i + 2, 9).setValue(reabrir ? "Sim" : "Não"); });
  return reabrir ? "Campanha reaberta." : "Campanha encerrada: novas adesões bloqueadas.";
}

// ---------- Dev: interesse ----------
function registrarInteresseIncentivada(nome, campanhaId, opcao) {
  nome = String(nome || "").trim().toUpperCase();
  if (!nome) throw new Error("Escolha o colaborador.");
  const dados = obterIncentivadas();
  const camp = dados.campanhas.filter(c => c.id === campanhaId)[0];
  if (!camp) throw new Error("Campanha não encontrada.");
  if (!camp.ativa || camp.encerrada) throw new Error("O prazo de adesão desta campanha já terminou.");
  const op = camp.opcoes.filter(o => o.opcao === opcao)[0];
  if (!op) throw new Error("Opção não encontrada.");

  const aba = abaInc_(INC.abaAdesoes, INC.cabAdesoes);
  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  const linhas = dadosAba_(aba, 8);
  let mudou = false;
  linhas.forEach((r, i) => {
    if (String(r[2]).toUpperCase() === nome && String(r[3]) === campanhaId && String(r[5]) !== "Cancelado") {
      if (String(r[5]) === "Confirmado") throw new Error("A adesão de " + primeiroNome_(nome) + " já foi confirmada pelo gestor. Fale com ele para mudar.");
      aba.getRange(i + 2, 2, 1, 7).setValues([[new Date(), nome, campanhaId, opcao, "Interesse", usuario, "Trocou de opção"]]);
      mudou = true;
    }
  });
  if (!mudou) aba.appendRow(["ADS-" + Utilities.getUuid().slice(0, 8), new Date(), nome, campanhaId, opcao, "Interesse", usuario, ""]);

  try {
    if (emailLider_()) enviarEmail_(emailLider_(), "🏝️ Interesse em férias incentivadas: " + nome, "Férias incentivadas",
      "<p><b>" + esc_(nome) + "</b> " + (mudou ? "trocou para" : "tem interesse na") + " <b>" + esc_(opcao) + "</b> da campanha <b>" + esc_(camp.nome) + "</b>: " +
      "férias de " + fmt_(parseData_(op.inicio)) + " a " + fmt_(parseData_(op.fim)) + ", retorno em " + fmt_(parseData_(op.retorno)) + ".</p>" +
      "<p>Para confirmar ou recusar, abra o dashboard → <b>Férias incentivadas</b>.</p>");
  } catch (e) {}
  return (mudou ? "Interesse atualizado: " : "Interesse registrado: ") + opcao + ". O gestor vai confirmar.";
}

function cancelarInteresseIncentivada(nome, campanhaId) {
  nome = String(nome || "").trim().toUpperCase();
  const aba = abaInc_(INC.abaAdesoes, INC.cabAdesoes);
  let n = 0;
  dadosAba_(aba, 8).forEach((r, i) => {
    if (String(r[2]).toUpperCase() === nome && String(r[3]) === campanhaId && String(r[5]) === "Interesse") { aba.getRange(i + 2, 6).setValue("Cancelado"); n++; }
  });
  if (!n) throw new Error("Não há interesse em aberto para cancelar.");
  return "Interesse cancelado.";
}

// ---------- Gestor: decidir ----------
function decidirIncentivada(adesaoId, confirmar, obs) {
  if (!ehGestorAtual_()) throw new Error("Só o gestor pode confirmar ou recusar.");
  const aba = abaInc_(INC.abaAdesoes, INC.cabAdesoes);
  const linhas = dadosAba_(aba, 8);
  const i = linhas.findIndex(r => String(r[0]) === adesaoId);
  if (i < 0) throw new Error("Adesão não encontrada.");
  const r = linhas[i];
  if (String(r[5]) !== "Interesse") throw new Error("Esta adesão já foi " + String(r[5]).toLowerCase() + ".");
  const nome = String(r[2]).toUpperCase();
  const camp = obterIncentivadas().campanhas.filter(c => c.id === String(r[3]))[0];
  const op = camp && camp.opcoes.filter(o => o.opcao === String(r[4]))[0];
  if (!op) throw new Error("A opção desta adesão não existe mais.");

  if (confirmar) lancarIncentivada_(nome, camp, op);
  aba.getRange(i + 2, 6).setValue(confirmar ? "Confirmado" : "Recusado");
  if (obs) aba.getRange(i + 2, 8).setValue(obs);

  const email = obterEmail_(nome);
  if (email) {
    enviarEmail_(email, (confirmar ? "✅ Férias incentivadas confirmadas" : "❌ Férias incentivadas não confirmadas") + " (" + op.opcao + ")",
      confirmar ? "Férias incentivadas confirmadas 🏝️" : "Férias incentivadas",
      "<p>Olá, <b>" + esc_(primeiroNome_(nome)) + "</b>!</p>" +
      (confirmar
        ? "<p>Sua adesão à <b>" + esc_(op.opcao) + "</b> da campanha <b>" + esc_(camp.nome) + "</b> foi <b>confirmada</b>:</p><ul>" +
          "<li>Férias legais: <b>" + fmt_(parseData_(op.inicio)) + " a " + fmt_(parseData_(op.fim)) + "</b> (" + op.diasFerias + " dias)</li>" +
          (op.abonados.length ? "<li>Dias abonados: <b>" + op.abonados.map(d => fmt_(parseData_(d))).join(", ") + "</b></li>" : "") +
          "<li>Retorno: <b>" + fmt_(parseData_(op.retorno)) + "</b> · " + op.totalDescanso + " dias de descanso</li></ul>" +
          "<p>Já está na agenda do time. Lembre de solicitar as férias legais no sistema do RH.</p>"
        : "<p>Sua adesão à <b>" + esc_(op.opcao) + "</b> da campanha <b>" + esc_(camp.nome) + "</b> não foi confirmada.</p>") +
      (obs ? "<div style='background:#f8f9fa;border-left:4px solid #1a73e8;padding:10px 12px'><b>Comentário do gestor:</b><br>" + esc_(obs) + "</div>" : ""));
  }
  return confirmar ? primeiroNome_(nome) + ": confirmado, lançado na planilha e na agenda." : primeiroNome_(nome) + ": recusado.";
}

// Férias legais na aba Controle + abonos na aba Folgas + um evento na agenda até o último dia de descanso
function lancarIncentivada_(nome, camp, op) {
  const ss = planilha_();
  const ini = parseData_(op.inicio), fim = parseData_(op.fim), ret = parseData_(op.retorno);
  const ctrl = ss.getSheetByName(CONFIG_APROVACAO.abaControle);
  const linhaC = Math.max(ctrl.getLastRow(), 1) + 1;
  ctrl.getRange(linhaC, 1, 1, 17).setValues([[
    "-", nome, "-", "-", "Férias Incentivadas (" + op.opcao + ")", "", "", "",
    "Sugestão", ini, fim, op.diasFerias, ret, "", "Espelhado no Calendar", new Date(), camp.nome
  ]]);
  ctrl.getRange(linhaC, 1, 1, 17).setBackground("#e0f2f1");

  const folgas = ss.getSheetByName("Folgas e Emendas - 2026") || inicializarTodasAbas();
  op.abonados.forEach(d => {
    const dd = parseData_(d);
    folgas.appendRow([nome, "ABONO (Férias Incentivadas)", iso_(dd), 1, iso_(new Date(dd.getFullYear(), dd.getMonth(), dd.getDate() + 1)), new Date(), "Férias Incentivadas | Espelhado no Calendar"]);
  });
  const ultimo = new Date(ret.getFullYear(), ret.getMonth(), ret.getDate() - 1);
  injetarEventoNoCalendarioGoogle("Ausência: " + nome + " [Férias Incentivadas - " + op.opcao + "]", iso_(ini), iso_(ultimo));
}
