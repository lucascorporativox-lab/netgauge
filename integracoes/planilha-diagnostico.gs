/**
 * NetGauge · Planilha do formulário de diagnóstico (Google Apps Script)
 *
 * Recebe os cadastros do modal de diagnóstico do site, completos ou não,
 * e grava uma linha por pessoa na aba "Diagnósticos". O site envia o mesmo
 * "id" a cada atualização, então a linha é atualizada em vez de duplicada.
 *
 * COMO PUBLICAR (uma vez só):
 * 1. Crie uma planilha no Google Sheets (ex.: "NetGauge - Diagnósticos").
 * 2. Menu Extensões > Apps Script. Apague o conteúdo e cole este arquivo inteiro. Salve.
 * 3. Clique em Implantar > Nova implantação > tipo "App da Web".
 *      - Executar como: Eu
 *      - Quem pode acessar: Qualquer pessoa
 *    Autorize o acesso quando o Google pedir.
 * 4. Copie a URL do App da Web (termina em /exec) e cole em PLANILHA_URL no index.html.
 *
 * Ao alterar este código depois, use Implantar > Gerenciar implantações > editar >
 * Versão: "Nova versão", para manter a mesma URL.
 */

const ABA = 'Diagnósticos';
const COLUNAS = [
  ['id', 'ID'],
  ['atualizado_em', 'Atualizado em'],
  ['status', 'Status'],
  ['nome', 'Nome'],
  ['whatsapp', 'WhatsApp'],
  ['provedor', 'Provedor'],
  ['cidade', 'Cidade / UF'],
  ['interesse', 'Interesse'],
  ['problemas', 'Problemas hoje'],
  ['erp', 'ERP'],
  ['monitoramento', 'Monitoramento'],
  ['observacoes', 'Observações'],
  ['acesso_via', 'Acesso via'],
  ['pagina', 'Página'],
];

/* abrir a URL /exec no navegador mostra esta mensagem: serve para conferir se a implantação está no ar */
function doGet() {
  return resposta('NetGauge: planilha de diagnósticos funcionando. O site grava os cadastros por aqui.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (!/^[A-Za-z0-9-]{8,64}$/.test(String(d.id || ''))) return resposta('id inválido');
    if (String(d.whatsapp || '').replace(/\D/g, '').length < 10) return resposta('sem whatsapp');

    const aba = obterAba();
    const linha = COLUNAS.map(([chave]) => (chave === 'atualizado_em' ? new Date() : limpar(d[chave])));

    const total = aba.getLastRow() - 1;
    const ids = total > 0 ? aba.getRange(2, 1, total, 1).getValues().map((r) => r[0]) : [];
    const pos = ids.indexOf(d.id);
    if (pos >= 0) {
      const iStatus = COLUNAS.findIndex(([c]) => c === 'status');
      const atual = aba.getRange(pos + 2, iStatus + 1).getValue();
      if (atual === 'concluído') linha[iStatus] = 'concluído'; // nunca volta para "incompleto"
      aba.getRange(pos + 2, 1, 1, linha.length).setValues([linha]);
    } else {
      aba.appendRow(linha);
    }
    return resposta('ok');
  } catch (err) {
    console.error('Erro ao gravar diagnóstico:', err);
    return resposta('erro');
  } finally {
    lock.releaseLock();
  }
}

function obterAba() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  let aba = planilha.getSheetByName(ABA);
  if (!aba) {
    aba = planilha.insertSheet(ABA);
    aba.appendRow(COLUNAS.map(([, titulo]) => titulo));
    aba.setFrozenRows(1);
    aba.getRange(1, 1, 1, COLUNAS.length).setFontWeight('bold');
  }
  return aba;
}

/* texto simples, limitado e sem virar fórmula na planilha */
function limpar(v) {
  let s = String(v == null ? '' : v).slice(0, 2000);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function resposta(txt) {
  return ContentService.createTextOutput(txt);
}
