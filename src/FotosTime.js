// =========================================================================
// FOTOS DO TIME — vêm do Agente A.I.F.R.E.A.D (aba TIME_FOTOS: e-mail → foto)
// -------------------------------------------------------------------------
// sincronizarFotosTime() lê as fotos da planilha do Agente, liga cada e-mail
// a uma pessoa do time (pela aba Contatos ou pelo nome no e-mail) e guarda
// na aba "Fotos" desta planilha. Também completa e-mails vazios em Contatos.
// O dashboard busca as fotos com obterFotosTime().
// =========================================================================

const FOTOS_CONFIG = {
  aba: "Fotos",
  abaOrigem: "TIME_FOTOS",
  idAgenteOriginal: "1hxXH0Omh-HvqWwrhH2gnriAQuwcCa4_ZMUL8VD1NbsI" // Agente A.I.F.R.E.A.D Controle Executivo (planilha original)
};

function idPlanilhaFotos_() {
  const p = PropertiesService.getScriptProperties().getProperty("ID_PLANILHA_FOTOS");
  if (p) return p;
  return ehPlanilhaOriginal_() ? FOTOS_CONFIG.idAgenteOriginal : "";
}

// Liga e-mail → nome: tokens do e-mail precisam aparecer no nome (ex.: silva.carlos → CARLOS EDUARDO NONATO DA SILVA)
function nomePorEmail_(email, nomes) {
  const local = String(email).split("@")[0].toLowerCase().replace(/[0-9]/g, "");
  const toks = local.split(/[._-]+/).filter(t => t.length >= 2);
  if (!toks.length) return "";
  const cand = nomes.filter(n => {
    const nt = normalizar_(n).toLowerCase().split(/\s+/);
    return toks.every(t => nt.some(x => x === t || (t.length > 3 && t.indexOf(x) >= 0 && x.length >= 4)));
  });
  return cand.length === 1 ? cand[0] : "";
}

function sincronizarFotosTime(idOpcional) {
  const props = PropertiesService.getScriptProperties();
  if (idOpcional) props.setProperty("ID_PLANILHA_FOTOS", String(idOpcional).trim());
  const id = idPlanilhaFotos_();
  if (!id) throw new Error("Informe a planilha do Agente A.I.F.R.E.A.D (onde ficam as fotos) em Configuração.");
  const origem = SpreadsheetApp.openById(id).getSheetByName(FOTOS_CONFIG.abaOrigem);
  if (!origem || origem.getLastRow() < 2) throw new Error("A planilha informada não tem a aba " + FOTOS_CONFIG.abaOrigem + " com fotos.");

  const ss = planilha_();
  let nomes = [];
  try { nomes = buscarNomesTime().map(n => String(n).trim().toUpperCase()); } catch (e) {}
  if (!nomes.length) throw new Error("Cadastre as pessoas do time antes (Configuração → Pessoas do time).");

  // e-mails já cadastrados em Contatos
  const contatos = lerContatos_(); // NOME normalizado → e-mail
  const porEmail = {};
  Object.keys(contatos).forEach(n => { if (contatos[n]) porEmail[String(contatos[n]).toLowerCase()] = n; });
  const gestor = String(emailLider_() || "").toLowerCase();

  const fotos = [], emailsNovos = {};
  origem.getRange(2, 1, origem.getLastRow() - 1, 2).getValues().forEach(r => {
    const chave = String(r[0] || "").trim().toLowerCase(), foto = String(r[1] || "");
    if (!chave || foto.indexOf("data:image/") !== 0) return;
    let email = chave, nome = "";
    if (chave === "lid_tl" && gestor) email = gestor;           // foto do Team Leader = gestor configurado
    else if (chave.indexOf("@") < 0) return;                     // demais gestores (tribe, gerentes, PO) ficam de fora
    if (porEmail[email]) nome = nomes.filter(n => normalizar_(n) === porEmail[email])[0] || "";
    if (!nome) { nome = nomePorEmail_(email, nomes); if (nome && !contatos[normalizar_(nome)]) emailsNovos[nome] = email; }
    if (nome) fotos.push([nome, email, foto, new Date()]);
  });

  // grava a aba Fotos
  let aba = ss.getSheetByName(FOTOS_CONFIG.aba);
  if (!aba) { aba = ss.insertSheet(FOTOS_CONFIG.aba); }
  aba.clearContents();
  aba.getRange(1, 1, 1, 4).setValues([["Nome", "E-mail", "Foto", "Atualizado em"]]).setFontWeight("bold");
  if (fotos.length) aba.getRange(2, 1, fotos.length, 4).setValues(fotos);
  try { aba.hideSheet(); } catch (e) {}

  // completa e-mails vazios na aba Contatos
  const qtdEmails = Object.keys(emailsNovos).length;
  if (qtdEmails) {
    const c = ss.getSheetByName("Contatos");
    if (c && c.getLastRow() > 1) {
      const vals = c.getRange(2, 1, c.getLastRow() - 1, 2).getValues();
      vals.forEach((r, i) => { const n = String(r[0]).trim().toUpperCase(); if (!r[1] && emailsNovos[n]) c.getRange(i + 2, 2).setValue(emailsNovos[n]); });
      const jaTem = {}; vals.forEach(r => jaTem[String(r[0]).trim().toUpperCase()] = 1);
      const faltam = Object.keys(emailsNovos).filter(n => !jaTem[n]).map(n => [n, emailsNovos[n]]);
      if (faltam.length) c.getRange(c.getLastRow() + 1, 1, faltam.length, 2).setValues(faltam);
    }
    _cacheEmails = null;
  }

  const semFoto = nomes.filter(n => !fotos.some(f => f[0] === n));
  props.setProperty("FOTOS_ATUALIZADAS", new Date().toISOString());
  return fotos.length + " foto(s) do time atualizada(s)" + (qtdEmails ? " e " + qtdEmails + " e-mail(s) completado(s) em Contatos" : "") + "." +
    (semFoto.length ? " Sem foto: " + semFoto.map(primeiroNome_).join(", ") + "." : "");
}

// Para o dashboard: { NOME: dataUrl }
function obterFotosTime() {
  const aba = planilha_().getSheetByName(FOTOS_CONFIG.aba);
  const out = {};
  if (aba && aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0] && String(r[2]).indexOf("data:image/") === 0) out[String(r[0]).toUpperCase()] = String(r[2]); });
  return out;
}
