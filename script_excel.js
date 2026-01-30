const SPREADSHEET_ID = '1O3xoPVPf__VIRbZwh-hxMBTT8WDEJual2ju0FKTZKus';
const DOC_TEMPLATE_ID = '1vqJwluEpgNG0zRCfSMFF10CUWrOMGASMLjkW1yXxHxE';
const OUTPUT_FOLDER_ID = '1qNB12ZkYdIyjzHPSyxojXxXz4GiYGesB';

function doPost(e) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { return ContentService.createTextOutput("Busy"); }

  try {
    const jsonString = e.postData.contents;
    const netlifyData = JSON.parse(jsonString);
    const d = netlifyData.data;

    // --- 1. LIMPEZA E FORMATAÇÃO AUTOMÁTICA ---
    if (d.cnpj) d.cnpj = formatCNPJ(d.cnpj);
    if (d.cpf) d.cpf = formatCPF(d.cpf);
    if (d.cep) d.cep = formatCEP(d.cep);
    if (d.cepSocio) d.cepSocio = formatCEP(d.cepSocio);
    // ------------------------------------------

    // --- 2. PROTEÇÃO CONTRA DUPLICIDADE (CACHE DE 1 HORA) ---
    // Cria uma ID única baseada nos dados principais
    const uniqueId = (d.cnpj || "") + "_" + (d.razaoSocial || "") + "_" + (d.emailEmpresa || "");
    const cleanId = uniqueId.replace(/[^a-zA-Z0-9]/g, "");

    const scriptProperties = PropertiesService.getScriptProperties();
    const lastTime = scriptProperties.getProperty(cleanId);
    const now = new Date().getTime();

    // Se o Netlify tentar reenviar em 10, 20 ou 50 minutos, será ignorado.
    if (lastTime && (now - parseInt(lastTime) < 3600000)) {
      return ContentService.createTextOutput("Duplicado (Cache 1h) - Ignorado");
    }
    
    // Grava o horário atual na memória
    scriptProperties.setProperty(cleanId, now.toString());
    // ------------------------------------------------------------

    saveToSheet(d);
    SpreadsheetApp.flush(); // Garante que salvou na planilha antes de gerar arquivos
    generateFiles(d);

    return ContentService.createTextOutput("Sucesso");

  } catch (error) {
    return ContentService.createTextOutput("Erro: " + error.toString());
  } finally {
    lock.releaseLock();
  }
}

// --- FUNÇÕES DE FORMATAÇÃO ---

function formatCPF(v) {
  v = v.toString().replace(/\D/g, ""); // Remove letras
  if (v.length === 11) {
    // Formato: XXX.XXX.XXX-XX
    return v.replace(/(\d{3})(\d{3})(\d{3})(\d{2})/, "$1.$2.$3-$4");
  }
  return v;
}

function formatCNPJ(v) {
  v = v.toString().replace(/\D/g, ""); 
  if (v.length === 14) {
    // Formato: XX.XXX.XXX/XXXX-XX
    return v.replace(/(\d{2})(\d{3})(\d{3})(\d{4})(\d{2})/, "$1.$2.$3/$4-$5");
  }
  return v;
}

function formatCEP(v) {
  v = v.toString().replace(/\D/g, ""); 
  if (v.length === 8) {
    // Formato: XX.XXX-XXX
    return v.replace(/(\d{2})(\d{3})(\d{3})/, "$1.$2-$3");
  }
  return v;
}

// ----------------------------------------

function saveToSheet(d) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getActiveSheet();
  sheet.appendRow([
    d.razaoSocial, d.cnpj, d.endereco, d.bairro, d.cidade, d.uf, d.cep, 
    d.telefone, d.celular, d.emailEmpresa, d.banco, d.agencia, d.conta, 
    d.pix, d.nomeSocio, d.cpf, d.rg, d.orgaoExpedidor, d.dataEmissao, 
    d.nascimento, d.nacionalidade, d.estadoCivil, d.profissao, d.emailSocio, 
    d.enderecoSocio, d.bairroSocio, d.cidadeSocio, d.ufSocio, d.cepSocio,
    new Date()
  ]);
}

function generateFiles(d) {
  const folder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  const token = ScriptApp.getOAuthToken();
  const nomeBase = (d.razaoSocial || "Cliente") + " - " + (d.cnpj || "");
  
  // 1. DOCX
  const tempDocFile = DriveApp.getFileById(DOC_TEMPLATE_ID).makeCopy(nomeBase, folder);
  const doc = DocumentApp.openById(tempDocFile.getId());
  const body = doc.getBody();
  
  const campos = [
    "razaoSocial", "cnpj", "cep", "endereco", "bairro", "cidade", "uf", 
    "telefone", "celular", "emailEmpresa", "banco", "agencia", "conta", 
    "pix", "nomeSocio", "cpf", "rg", "orgaoExpedidor", "dataEmissao", 
    "nascimento", "nacionalidade", "estadoCivil", "profissao", 
    "emailSocio", "cepSocio", "enderecoSocio", "bairroSocio", 
    "cidadeSocio", "ufSocio"
  ];

  campos.forEach(campo => {
    // Procura por {camelCase} no documento
    body.replaceText("{" + campo + "}", d[campo] || "");
  });
  
  doc.saveAndClose();

  const docxUrl = "https://docs.google.com/feeds/download/documents/export/Export?id=" + tempDocFile.getId() + "&exportFormat=doc";
  const docxBlob = UrlFetchApp.fetch(docxUrl, {headers: {Authorization: "Bearer " + token}}).getBlob();
  docxBlob.setName(nomeBase + ".docx");
  folder.createFile(docxBlob);
  tempDocFile.setTrashed(true);

  // 2. XLSX
  const tempSheet = SpreadsheetApp.create(nomeBase);
  const tempSheetFile = DriveApp.getFileById(tempSheet.getId());
  tempSheetFile.moveTo(folder);
  const s = tempSheet.getActiveSheet();
  
  s.appendRow(["CAMPO", "VALOR"]);
  for (const key in d) {
    if (typeof d[key] !== 'object') {
      s.appendRow([key, d[key]]);
    }
  }
  SpreadsheetApp.flush();

  const xlsxUrl = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + tempSheet.getId() + "&exportFormat=xlsx";
  const xlsxBlob = UrlFetchApp.fetch(xlsxUrl, {headers: {Authorization: "Bearer " + token}}).getBlob();
  xlsxBlob.setName(nomeBase + ".xlsx");
  folder.createFile(xlsxBlob);
  tempSheetFile.setTrashed(true);
}