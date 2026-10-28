/**
 * Declaração de Vale-Transporte — backend (Google Apps Script)
 * ------------------------------------------------------------
 * Grava as respostas na planilha onde este script está instalado
 * e as fotos (comprovante, cartão, carta) numa pasta privada do Google Drive.
 *
 * 1) Troque a SENHA_RH abaixo por uma senha forte (só o RH deve saber).
 * 2) Execute a função "configurar" uma vez (menu Executar).
 * 3) Implante como "App da Web" com acesso "Qualquer pessoa".
 * Veja o passo a passo completo no arquivo LEIA-ME.md.
 */

const SENHA_RH = "TROQUE-ESTA-SENHA";          // senha do painel do RH
const NOME_PASTA = "Vale-Transporte - Anexos";  // pasta no Drive para as fotos

const EMPRESAS_INICIAIS = [
  ["VEGAS ALARMES MONITORADOS LTDA", "46.937.114/0001-63"],
  ["VEGAS MONITORAMENTO ELETRÔNICO E SERVIÇOS TERCEIRIZADOS LTDA", "50.196.573/0001-00"],
  ["VEGAS MONITORAMENTO ELETRÔNICO E SERVIÇOS TERCEIRIZADOS LTDA (FILIAL)", "50.196.573/0002-91"],
  ["VEGAS SEGURANÇA E TECNOLOGIA PREVENTIVA LTDA", "33.549.671/0001-65"],
  ["VEGAS VIGILÂNCIA E SEGURANÇA LTDA", "29.890.721/0001-04"],
  ["SERVIÇOS CORPORATIVOS VGS LTDA", "46.955.475/0001-32"]
];

const COLS_CONVITES = ["Código", "Nome do funcionário", "Razão social", "CNPJ", "WhatsApp", "Criado em", "Situação", "Respondido em"];
const COLS_RESPOSTAS = ["Enviado em", "Código", "Razão social", "CNPJ", "Nome", "Opção", "Local", "Data da declaração",
  "Rua", "Número", "Complemento", "Bairro", "Cidade", "UF", "CEP",
  "Ida (empresa / linha / valor)", "Volta (empresa / linha / valor)", "Custo diário (R$)",
  "Nº do cartão", "Possui cartão", "Comprovante de residência", "Foto do cartão", "Carta de próprio punho",
  "Comprovante com data dos últimos 2 meses (declarado)", "Declarações aceitas", "Termo LGPD aceito", "Versão do termo",
  "Assinatura eletrônica", "Dispositivo"];

/* ================= CONFIGURAÇÃO (executar 1 vez) ================= */
function configurar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emp = aba_("Empresas", ["Razão social", "CNPJ"]);
  if (emp.getLastRow() < 2) emp.getRange(2, 1, EMPRESAS_INICIAIS.length, 2).setValues(EMPRESAS_INICIAIS);
  aba_("Convites", COLS_CONVITES);
  aba_("Respostas", COLS_RESPOSTAS);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("PASTA_ID")) {
    const pasta = DriveApp.createFolder(NOME_PASTA);
    props.setProperty("PASTA_ID", pasta.getId());
  }
  const padrao = ss.getSheetByName("Página1") || ss.getSheetByName("Sheet1");
  if (padrao && ss.getSheets().length > 1 && padrao.getLastRow() === 0) ss.deleteSheet(padrao);
  Logger.log("Configuração concluída. Agora implante como App da Web.");
}

function aba_(nome, cabecalho) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) sh = ss.insertSheet(nome);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]).setFontWeight("bold").setBackground("#16243A").setFontColor("#FFFFFF");
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ================= ROTAS ================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === "convite") return json_(buscarConvite_(p.codigo));
    return json_({ ok: true, servico: "vale-transporte" });
  } catch (err) { return json_({ ok: false, erro: String(err.message || err) }); }
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, erro: "Requisição inválida." }); }
  try {
    switch (d.action) {
      case "enviar":        return json_(enviar_(d));
      case "login":         return json_(rh_(d, () => ({ ok: true })));
      case "empresas":      return json_(rh_(d, () => ({ ok: true, empresas: empresas_() })));
      case "gerarConvite":  return json_(rh_(d, () => gerarConvite_(d)));
      case "excluirConvite":return json_(rh_(d, () => excluirConvite_(d.codigo)));
      case "listar":        return json_(rh_(d, () => listar_()));
      case "pdf":           return json_(rh_(d, () => gerarPdf_(d)));
      default:              return json_({ ok: false, erro: "Ação desconhecida." });
    }
  } catch (err) { return json_({ ok: false, erro: String(err.message || err) }); }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function rh_(d, fn) {
  if (!d.chave || d.chave !== SENHA_RH) { Utilities.sleep(800); return { ok: false, erro: "Senha incorreta.", auth: false }; }
  return fn();
}

/* ================= DADOS ================= */
function empresas_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Empresas");
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getDisplayValues()
    .filter(r => r[0] && r[1]).map(r => ({ razao: r[0].trim(), cnpj: r[1].trim() }));
}

function linhas_(nome) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
}

function buscarConvite_(codigo) {
  codigo = String(codigo || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
  if (codigo.length !== 6) return { ok: false, erro: "Código inválido." };
  const r = linhas_("Convites").find(l => l[0] === codigo);
  if (!r) return { ok: false, erro: "Código não encontrado." };
  if (r[6] === "Excluído") return { ok: false, erro: "Este link foi cancelado pelo RH." };
  return { ok: true, convite: { codigo: r[0], nome: r[1], empresa: { razao: r[2], cnpj: r[3] } } };
}

function novoCodigo_() {
  const a = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  const existentes = new Set(linhas_("Convites").map(l => l[0]));
  for (let t = 0; t < 50; t++) {
    let c = ""; for (let i = 0; i < 6; i++) c += a.charAt(Math.floor(Math.random() * a.length));
    if (!existentes.has(c)) return c;
  }
  throw new Error("Não foi possível gerar um código. Tente de novo.");
}

function gerarConvite_(d) {
  const emp = empresas_().find(x => x.cnpj === d.cnpj);
  const nome = String(d.nome || "").trim();
  if (!emp) return { ok: false, erro: "Empresa não encontrada no cadastro." };
  if (nome.length < 3) return { ok: false, erro: "Informe o nome do funcionário." };
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const codigo = novoCodigo_();
    const agora = Utilities.formatDate(new Date(), "America/Sao_Paulo", "dd/MM/yyyy HH:mm");
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Convites")
      .appendRow([codigo, seguro_(nome), emp.razao, emp.cnpj, seguro_(d.tel || ""), agora, "Pendente", ""]);
    return { ok: true, convite: { codigo, nome, empresa: emp, tel: d.tel || "", criadoEm: agora, situacao: "Pendente" } };
  } finally { lock.releaseLock(); }
}

function excluirConvite_(codigo) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Convites");
  const dados = linhas_("Convites");
  const i = dados.findIndex(l => l[0] === codigo);
  if (i < 0) return { ok: false, erro: "Link não encontrado." };
  sh.getRange(i + 2, 7).setValue("Excluído");
  return { ok: true };
}

function listar_() {
  const convites = linhas_("Convites").filter(l => l[6] !== "Excluído").map(l => ({
    codigo: l[0], nome: l[1], empresa: { razao: l[2], cnpj: l[3] }, tel: l[4], criadoEm: l[5], situacao: l[6], respondidoEm: l[7]
  })).reverse();
  const respostas = linhas_("Respostas").map(l => {
    const o = {}; COLS_RESPOSTAS.forEach((c, i) => o[c] = l[i]); return o;
  }).reverse();
  return { ok: true, convites, respostas, empresas: empresas_() };
}

/* ================= ENVIO DO FUNCIONÁRIO ================= */
function enviar_(d) {
  const cv = buscarConvite_(d.codigo);
  if (!cv.ok) return cv;
  if (d.opcao !== "opto" && d.opcao !== "nao_opto") return { ok: false, erro: "Escolha uma opção." };
  if (!d.lgpd || !d.assinatura) return { ok: false, erro: "Aceite o termo e confirme a assinatura." };
  if (d.opcao === "opto" && !d.fotoComprovante) return { ok: false, erro: "Falta a foto do comprovante de residência." };
  if (d.opcao === "nao_opto" && !d.fotoCarta) return { ok: false, erro: "Falta a foto da carta de próprio punho." };

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const agora = new Date();
    const carimbo = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");
    const base = Utilities.formatDate(agora, "America/Sao_Paulo", "yyyyMMdd-HHmm") + "_" + d.codigo + "_" + limparNome_(d.nome);
    const linkComp = d.opcao === "opto" ? salvarFoto_(d.fotoComprovante, base + "_comprovante") : "";
    const linkCart = d.fotoCartao ? salvarFoto_(d.fotoCartao, base + "_cartao") : "";
    const linkCarta = d.opcao === "nao_opto" ? salvarFoto_(d.fotoCarta, base + "_carta") : "";
    const e = d.endereco || {};
    const trecho = l => (l || []).map(x => `${x.empresa} / ${x.linha} / R$ ${Number(x.valor || 0).toFixed(2).replace(".", ",")}`).join(" + ");
    const custo = [].concat(d.ida || [], d.volta || []).reduce((a, x) => a + Number(x.valor || 0), 0);
    const emp = cv.convite.empresa;

    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respostas").appendRow([
      carimbo, d.codigo, emp.razao, emp.cnpj, seguro_(d.nome), d.opcao === "opto" ? "Opta" : "Não opta / cancelou",
      seguro_(d.local), seguro_(d.dataTexto),
      seguro_(e.rua), seguro_(e.numero), seguro_(e.complemento), seguro_(e.bairro), seguro_(e.cidade), seguro_(e.uf), seguro_(e.cep),
      seguro_(trecho(d.ida)), seguro_(trecho(d.volta)), d.opcao === "opto" ? custo.toFixed(2).replace(".", ",") : "",
      seguro_(d.cartaoNumero || ""), d.opcao === "opto" ? (d.semCartao ? "Não" : "Sim") : (linkCart || d.cartaoNumero ? "Sim" : ""),
      linkComp, linkCart, linkCarta,
      d.opcao === "opto" ? (d.comprovanteData ? "Sim" : "Não") : "", d.declaracoes ? "Sim" : "Não",
      "Sim", seguro_(d.lgpdVersao || ""), "Sim", seguro_(String(d.dispositivo || "").slice(0, 150))
    ]);

    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Convites");
    const i = linhas_("Convites").findIndex(l => l[0] === d.codigo);
    if (i >= 0) sh.getRange(i + 2, 7, 1, 2).setValues([[d.opcao === "opto" ? "Respondeu - opta" : "Respondeu - não opta", carimbo]]);
    return { ok: true, empresa: emp, enviadoEm: carimbo };
  } finally { lock.releaseLock(); }
}

function salvarFoto_(dataUrl, nome) {
  const m = String(dataUrl || "").match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error("Foto inválida.");
  const pasta = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("PASTA_ID"));
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], nome + ".jpg");
  return pasta.createFile(blob).getUrl();
}

function limparNome_(s) { return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Za-z0-9]+/g, "-").slice(0, 40); }
/* Impede que textos digitados virem fórmulas na planilha */
function seguro_(v) { v = String(v == null ? "" : v).slice(0, 500); return /^[=+\-@]/.test(v) ? "'" + v : v; }

/* ================= PDF DA FICHA ================= */
const TXT_DECL_OPTA = [
  "Comprometo-me a atualizar as informações acima sempre que ocorrerem alterações e a utilizar os vales-transporte que me forem concedidos exclusivamente no percurso indicado.",
  "Estou ciente de que a declaração inexata que induza o empregador em erro, ou o uso indevido dos vales-transporte, configura justa causa para rescisão do contrato de trabalho por ato de improbidade (art. 482, \"a\", da CLT)."
];
const TXT_DECL_NAO = "Declaro que não desejo receber o vale-transporte (ou solicito o cancelamento do benefício atual) e estou ciente de que posso solicitá-lo novamente a qualquer momento, preenchendo nova declaração.";

function gerarPdf_(d) {
  const r = linhas_("Respostas").map(l => { const o = {}; COLS_RESPOSTAS.forEach((c, i) => o[c] = l[i]); return o; })
    .find(o => o["Código"] === d.codigo && o["Enviado em"] === d.enviadoEm);
  if (!r) return { ok: false, erro: "Resposta não encontrada na planilha." };
  const nome = "Ficha VT - " + limparNome_(r["Nome"]) + " - " + r["Código"] + ".pdf";
  const pdf = Utilities.newBlob(htmlFicha_(r), MimeType.HTML, "ficha.html").getAs(MimeType.PDF).setName(nome);
  return { ok: true, nome, pdf: Utilities.base64Encode(pdf.getBytes()) };
}

function imgDrive_(url) {
  const m = String(url || "").match(/\/d\/([\w-]+)/) || String(url || "").match(/[?&]id=([\w-]+)/);
  if (!m) return "";
  try {
    const b = DriveApp.getFileById(m[1]).getBlob();
    return "data:" + b.getContentType() + ";base64," + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ""; }
}

function htmlFicha_(r) {
  const h = v => String(v == null ? "" : v).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const opta = r["Opção"] === "Opta";
  const campo = (t, v, cs) => `<td${cs ? ` colspan="${cs}"` : ""} style="border:1px solid #999;padding:5px 7px;vertical-align:top"><div style="font-size:8pt;color:#555">${t}</div><div style="font-size:10pt">${h(v) || "&nbsp;"}</div></td>`;
  const titulo = t => `<div style="background:#16243A;color:#fff;font-weight:bold;font-size:10pt;padding:5px 7px;margin-top:12px">${t}</div>`;
  const caixa = on => on ? "[X]" : "[&nbsp;&nbsp;]";
  const trechos = txt => {
    const l = String(txt || "").split(" + ").filter(Boolean).map(x => x.split(" / "));
    if (!l.length) return `<tr><td colspan="3" style="border:1px solid #999;padding:5px 7px">—</td></tr>`;
    return l.map(p => `<tr><td style="border:1px solid #999;padding:5px 7px">${h(p[0])}</td><td style="border:1px solid #999;padding:5px 7px">${h(p[1])}</td><td style="border:1px solid #999;padding:5px 7px;text-align:right">${h(p[2])}</td></tr>`).join("");
  };
  const tabTrecho = (t, txt) => `<p style="font-size:9pt;font-weight:bold;margin:8px 0 3px">${t}</p>
    <table style="width:100%;border-collapse:collapse;font-size:9.5pt"><tr style="background:#eee"><th style="border:1px solid #999;padding:4px;text-align:left">Empresa de ônibus</th><th style="border:1px solid #999;padding:4px;text-align:left">Linha</th><th style="border:1px solid #999;padding:4px;text-align:right">Valor</th></tr>${trechos(txt)}</table>`;
  const anexo = (t, url) => { const src = imgDrive_(url); return src ? `<div style="page-break-before:always">${titulo(t)}<p style="text-align:center;margin-top:10px"><img src="${src}" width="520"></p></div>` : ""; };

  let corpo = `<table style="width:100%;border-collapse:collapse"><tr>${campo("Nome do empregado", r["Nome"], 2)}${campo("Local e data", r["Local"] + ", " + r["Data da declaração"])}</tr></table>
    ${titulo("Opção")}
    <p style="font-size:10pt;margin:6px 0">${caixa(opta)} Opto pelo vale-transporte &nbsp;&nbsp;&nbsp; ${caixa(!opta)} Não opto / solicito o cancelamento</p>`;
  if (opta) {
    corpo += `${titulo("Endereço residencial")}
    <table style="width:100%;border-collapse:collapse"><tr>${campo("Rua / Avenida", r["Rua"], 3)}${campo("Número", r["Número"])}</tr>
    <tr>${campo("Complemento", r["Complemento"])}${campo("Bairro", r["Bairro"])}${campo("Cidade / UF", r["Cidade"] + " / " + r["UF"])}${campo("CEP", r["CEP"])}</tr></table>
    ${titulo("Meios de transporte utilizados")}
    ${tabTrecho("Da residência para o trabalho", r["Ida (empresa / linha / valor)"])}
    ${tabTrecho("Do trabalho para a residência", r["Volta (empresa / linha / valor)"])}
    <p style="font-size:10pt;text-align:right;margin:6px 0"><b>Custo diário informado: R$ ${h(r["Custo diário (R$)"])}</b></p>`;
  }
  if (r["Possui cartão"] || r["Nº do cartão"] || r["Foto do cartão"])
    corpo += `${titulo("Cartão de passagem")}<table style="width:100%;border-collapse:collapse"><tr>${campo("Possui cartão", r["Possui cartão"] || "—")}${campo("Número do cartão", r["Nº do cartão"] || "—")}${campo("Foto anexada", r["Foto do cartão"] ? "Sim" : "Não")}</tr></table>`;
  corpo += `${titulo("Declarações")}<div style="font-size:9.5pt;line-height:1.45">` +
    (opta ? TXT_DECL_OPTA.map(t => `<p style="margin:6px 0">[X] ${h(t)}</p>`).join("") +
            `<p style="margin:6px 0">[${r["Comprovante com data dos últimos 2 meses (declarado)"] === "Sim" ? "X" : "&nbsp;&nbsp;"}] Declaro que o comprovante de residência anexado tem data dos últimos dois meses.</p><p style="margin:6px 0;color:#555">O desconto em folha é de até 6% do salário básico, conforme a Lei nº 7.418/1985.</p>`
          : `<p style="margin:6px 0">[X] ${h(TXT_DECL_NAO)}</p><p style="margin:6px 0">Carta de próprio punho assinada pelo empregado anexada a esta ficha.</p>`) + `</div>
    ${titulo("Registro da assinatura eletrônica")}
    <table style="width:100%;border-collapse:collapse"><tr>${campo("Enviado em", r["Enviado em"])}${campo("Código da requisição", r["Código"])}${campo("Termo LGPD", "Aceito · versão " + r["Versão do termo"])}</tr>
    <tr>${campo("Dispositivo", r["Dispositivo"], 3)}</tr></table>
    <table style="width:100%;margin-top:40px"><tr><td style="width:50%;text-align:center;font-size:9pt"><div style="border-top:1px solid #000;margin:0 20px;padding-top:3px">${h(r["Nome"])}<br>Empregado</div></td>
    <td style="width:50%;text-align:center;font-size:9pt"><div style="border-top:1px solid #000;margin:0 20px;padding-top:3px">${h(r["Razão social"])}<br>Empregador</div></td></tr></table>`;

  return `<html><head><meta charset="utf-8"></head><body style="font-family:Arial,Helvetica,sans-serif;color:#111;margin:0">
    <table style="width:100%;border-collapse:collapse;border-bottom:2px solid #16243A"><tr>
      <td><div style="font-size:14pt;font-weight:bold">Declaração de Beneficiário de Vale-Transporte</div><div style="font-size:9pt;color:#555">Lei nº 7.418/1985 · Decreto nº 10.854/2021</div></td>
      <td style="text-align:right;font-size:9pt"><b>${h(r["Razão social"])}</b><br>CNPJ ${h(r["CNPJ"])}</td></tr></table>
    ${corpo}
    ${opta ? anexo("Anexo · Comprovante de residência", r["Comprovante de residência"]) : anexo("Anexo · Carta de próprio punho", r["Carta de próprio punho"])}
    ${anexo("Anexo · Cartão de passagem", r["Foto do cartão"])}
  </body></html>`;
}
