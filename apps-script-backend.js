// ===================================================================
// SHOPEE OCORRÊNCIAS — Google Apps Script Backend
// ===================================================================
// Deploy como Web App: Extensões > Apps Script > Implantar > Novo deployment
// Tipo: App da Web | Executar como: Eu | Acesso: Qualquer pessoa
// ===================================================================

const SPREADSHEET_ID = '104dpraVtTatQDnqCvHm_13huPGsN2BrNEIpMShwO9xw'; // App Ocorrência
const HISTORICO_SPREADSHEET_ID = '1YAYwh0Gc61D4roOJ73qnT5Yo_UBqXbA9E-bPmyqx-OM';
const SHEET_NAME = 'Sheet1';
const HISTORICO_SHEET = 'HISTORICO';
const ALLOWED_DOMAINS = ['shopee.com', 'shopeemobile-external.com'];

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  try {
    const action = e.parameter.action;
    const token = e.parameter.token;
    
    // Validar token Google
    if (!token) {
      return output.setContent(JSON.stringify({ error: 'Token obrigatório' }));
    }
    
    const userInfo = validateToken(token);
    if (!userInfo) {
      return output.setContent(JSON.stringify({ error: 'Token inválido' }));
    }
    
    // Verificar domínio permitido
    const domain = userInfo.email.split('@')[1];
    if (!ALLOWED_DOMAINS.includes(domain)) {
      return output.setContent(JSON.stringify({ error: 'Domínio não autorizado' }));
    }
    
    let result;
    
    switch (action) {
      case 'getDrivers':
        result = getDrivers(e.parameter.query);
        break;
      case 'addOcorrencia':
        const data = JSON.parse(e.parameter.data);
        data.operador = userInfo.email;
        data.operadorNome = userInfo.name;
        result = addOcorrencia(data);
        break;
      case 'getOcorrencias':
        result = getOcorrencias(e.parameter);
        break;
      case 'updateStatus':
        result = updateStatus(e.parameter.rowId, e.parameter.newStatus, userInfo.email);
        break;
      case 'getStats':
        result = getStats();
        break;
      default:
        result = { error: 'Ação desconhecida' };
    }
    
    return output.setContent(JSON.stringify(result));
    
  } catch (err) {
    return output.setContent(JSON.stringify({ error: err.toString() }));
  }
}

function validateToken(token) {
  try {
    const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + token;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    
    if (data.error) return null;
    
    return {
      email: data.email,
      name: data.name || data.email.split('@')[0],
      picture: data.picture
    };
  } catch (e) {
    return null;
  }
}

function getDrivers(query) {
  const ss = SpreadsheetApp.openById(HISTORICO_SPREADSHEET_ID);
  const sheet = ss.getSheetByName(HISTORICO_SHEET);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const drivers = [];
  for (let i = 1; i < data.length; i++) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    
    if (query) {
      const q = query.toLowerCase();
      if (String(row.driver_id).includes(q) || 
          String(row.name).toLowerCase().includes(q) ||
          String(row.phone).includes(q)) {
        drivers.push(row);
      }
    } else {
      drivers.push(row);
    }
    
    if (drivers.length >= 50) break;
  }
  
  return { drivers };
}

function addOcorrencia(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  // Criar cabeçalhos se a planilha estiver vazia
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Data/Hora', 'Driver ID', 'Nome', 'HUB', 'Preferências',
      'Perfil', 'Baú', 'Qtd Pacotes', 'Ocorrência', 'Motivo',
      'Status', 'Código AT', 'Observações', 'Operador',
      'Operador Email', 'Latitude', 'Longitude', 'Imagem URL'
    ]);
  }
  
  const row = [
    new Date().toLocaleString('pt-BR', { timeZone: 'America/Sao_Paulo' }),
    data.driverId,
    data.nome,
    data.hub,
    data.preferencias,
    data.profile,
    data.bau,
    data.qtdPacotes,
    data.ocorrencia,
    data.motivo,
    data.status || 'Pendente',
    data.atCode,
    data.observacao,
    data.operadorNome,
    data.operador,
    data.latitude || '',
    data.longitude || '',
    data.imagemUrl || ''
  ];
  
  sheet.appendRow(row);
  
  // Enviar alerta se > 100 pacotes
  if (parseInt(data.qtdPacotes) > 100) {
    sendAlert(data);
  }
  
  return { success: true, row: sheet.getLastRow() };
}

function getOcorrencias(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) return { ocorrencias: [] };
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const ocorrencias = [];
  
  for (let i = data.length - 1; i >= 1; i--) {
    const row = {};
    headers.forEach((h, j) => { row[h] = data[i][j]; });
    row._rowIndex = i + 1;
    
    // Aplicar filtros
    if (params.hub && row['HUB'] !== params.hub) continue;
    if (params.status && row['Status'] !== params.status) continue;
    if (params.motivo && row['Motivo'] !== params.motivo) continue;
    if (params.perfil && row['Perfil'] !== params.perfil) continue;
    
    ocorrencias.push(row);
    if (ocorrencias.length >= 200) break;
  }
  
  return { ocorrencias };
}

function updateStatus(rowId, newStatus, operadorEmail) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  const statusCol = 11; // Coluna K (Status)
  sheet.getRange(parseInt(rowId), statusCol).setValue(newStatus);
  
  return { success: true };
}

function getStats() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) return { total: 0, pendentes: 0, ok: 0, excesso: 0, poucos: 0 };
  
  const data = sheet.getDataRange().getValues();
  let total = 0, pendentes = 0, ok = 0, excesso = 0, poucos = 0;
  const motivos = {};
  const perfis = {};
  
  for (let i = 1; i < data.length; i++) {
    total++;
    const status = data[i][10];
    const qtd = parseInt(data[i][7]) || 0;
    const motivo = data[i][9];
    const perfil = data[i][5];
    
    if (status === 'Pendente') pendentes++;
    if (status === 'OK') ok++;
    if (qtd > 100) excesso++;
    if (qtd > 0 && qtd < 30) poucos++;
    
    if (motivo) motivos[motivo] = (motivos[motivo] || 0) + 1;
    if (perfil) perfis[perfil] = (perfis[perfil] || 0) + 1;
  }
  
  return { total, pendentes, ok, excesso, poucos, motivos, perfis };
}

function sendAlert(data) {
  try {
    const subject = `⚠️ ALERTA: Ocorrência >100 pacotes — ${data.nome}`;
    const body = `
Driver: ${data.nome} (ID: ${data.driverId})
HUB: ${data.hub}
Pacotes: ${data.qtdPacotes}
Motivo: ${data.motivo}
Ocorrência: ${data.ocorrencia}
Operador: ${data.operadorNome}
AT: ${data.atCode}
    `;
    // Enviar para o operador registrado
    MailApp.sendEmail(data.operador, subject, body);
  } catch (e) {
    // Silenciar erros de email
  }
}
