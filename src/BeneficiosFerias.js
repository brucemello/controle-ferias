// =========================================================================
// BENEFÍCIOS — cadastro de admissão (GiftDay) e MVT, e situação do ResetDay
// Usado pela tela "Benefícios" do dashboard. Grava nas abas "Admisssão" e "MVT".
// =========================================================================

const MESES_MVT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

function obterBeneficios() {
  const ss = planilha_();
  let nomes = [];
  try { nomes = buscarNomesTime().map(n => String(n).trim().toUpperCase()); } catch (e) {}

  const mvtPorNome = {};
  const abaMvt = ss.getSheetByName("MVT");
  if (abaMvt && abaMvt.getLastRow() > 1) {
    abaMvt.getRange(2, 1, abaMvt.getLastRow() - 1, 5).getValues().forEach(r => {
      const n = normalizar_(r[0]);
      if (!n) return;
      if (!mvtPorNome[n]) mvtPorNome[n] = [];
      const qtd = parseInt(r[3], 10);
      if (!r[1] && !r[2] && (isNaN(qtd) || qtd === 0)) return; // linha só com o nome
      mvtPorNome[n].push({ mes: mesMvtNormalizado_(r[1]), ano: parseInt(r[2], 10) || "", utilizou: normalizar_(r[4]) === "SIM" });
    });
  }

  let usuario = "";
  try { usuario = Session.getActiveUser().getEmail(); } catch (e) {}
  const gestor = emailLider_();

  return {
    ano: new Date().getFullYear(),
    meses: MESES_MVT,
    podeEditar: !gestor || String(usuario).toLowerCase() === String(gestor).toLowerCase(),
    pessoas: nomes.map(n => {
      const s = consultarSaldosColaborador(n);
      return {
        nome: n,
        admissao: s.admissao,
        anos: s.anosEmpresa,
        gift: { total: s.giftTotal, usado: s.giftUsado, saldo: s.giftSaldo },
        reset: { saldo: s.resetSaldo, usadoEm: s.resetUsadoEm },
        mvt: mvtPorNome[normalizar_(n)] || [],
        mvtSaldo: s.mvtSaldo
      };
    })
  };
}

// lista: [{ nome, admissao: "yyyy-mm-dd" | "", mvt: [{mes, ano, utilizou}] }]
function salvarBeneficios(lista) {
  const gestor = emailLider_();
  if (gestor) validarGestor_();
  if (!Array.isArray(lista)) throw new Error("Nada para salvar.");
  const ss = planilha_();
  const nomesEditados = {};
  lista.forEach(p => { nomesEditados[normalizar_(p.nome)] = true; });

  // ---- Admissão: mantém quem não foi editado, regrava os editados ----
  const abaAdm = ss.getSheetByName("Admisssão") || ss.insertSheet("Admisssão");
  const restoAdm = abaAdm.getLastRow() > 1
    ? abaAdm.getRange(2, 1, abaAdm.getLastRow() - 1, 2).getValues().filter(r => r[0] && !nomesEditados[normalizar_(r[0])])
    : [];
  const novasAdm = lista.filter(p => p.admissao).map(p => {
    const d = parseData_(p.admissao);
    if (!d) throw new Error("Data de admissão inválida para " + p.nome + ".");
    return [String(p.nome).toUpperCase(), d];
  });
  regravarAba_(abaAdm, ["Nome", "Admissão"], restoAdm.concat(novasAdm), "dd/mm/yyyy", 2);

  // ---- MVT ----
  const abaMvt = ss.getSheetByName("MVT") || ss.insertSheet("MVT");
  const restoMvt = abaMvt.getLastRow() > 1
    ? abaMvt.getRange(2, 1, abaMvt.getLastRow() - 1, 5).getValues().filter(r => r[0] && !nomesEditados[normalizar_(r[0])])
    : [];
  const novasMvt = [];
  lista.forEach(p => {
    const itens = (p.mvt || []).filter(m => m.mes && m.ano);
    if (!itens.length) { novasMvt.push([String(p.nome).toUpperCase(), "", "", 0, ""]); return; }
    itens.forEach(m => novasMvt.push([String(p.nome).toUpperCase(), m.mes, parseInt(m.ano, 10), 1, m.utilizou ? "Sim" : "Não"]));
  });
  regravarAba_(abaMvt, ["Nome", "Mes", "Ano", "Quantidade", "Utilizou"], restoMvt.concat(novasMvt), null);

  return "Benefícios salvos: " + novasAdm.length + " data(s) de admissão e " + novasMvt.filter(r => r[3]).length + " MVT(s).";
}

function regravarAba_(aba, cabecalho, linhas, formatoData, colData) {
  const nc = cabecalho.length;
  if (aba.getLastRow() > 1) aba.getRange(2, 1, aba.getLastRow() - 1, Math.max(nc, aba.getLastColumn())).clearContent();
  aba.getRange(1, 1, 1, nc).setValues([cabecalho]).setFontWeight("bold");
  if (linhas.length) {
    aba.getRange(2, 1, linhas.length, nc).setValues(linhas.map(l => { const x = l.slice(0, nc); while (x.length < nc) x.push(""); return x; }));
    if (formatoData && colData) aba.getRange(2, colData, linhas.length, 1).setNumberFormat(formatoData);
  }
}

function mesMvtNormalizado_(v) {
  const s = normalizar_(v);
  if (!s) return "";
  const i = MESES_MVT.findIndex(m => normalizar_(m).indexOf(s.slice(0, 3)) === 0);
  return i >= 0 ? MESES_MVT[i] : String(v);
}
