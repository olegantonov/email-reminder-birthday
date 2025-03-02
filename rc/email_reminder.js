// Constantes para evitar erros de digitação e facilitar manutenções
const SHEET_EMAIL = "Pessoal.Email";
const SHEET_SAUDE = "Pessoal.Saude";
const RANGE_EMAIL = "B2:F";
const EMAIL_ADMIN = "daniel.marques@senado.leg.br"; // E-mail do administrador/gestor para os lembretes semanais

// Função para validar e-mail usando expressão regular
function validarEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// Obtém a lista de e-mails da planilha "Pessoal.Email"
function obterEmailsParaEnvio() {
  console.log("Obtendo a planilha de e-mails.");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_EMAIL);
  if (!sheet) {
    console.error("Não foi possível encontrar a aba '" + SHEET_EMAIL + "'. Verifique o nome da aba.");
    return [];
  }
  
  const range = sheet.getRange(RANGE_EMAIL);
  const values = range.getValues();
  
  // Filtra os usuários que possuem "SIM" na coluna B e e-mails válidos na coluna F
  const emailsParaEnvio = values
    .filter(row => row[0].toString().toUpperCase() === "SIM" && validarEmail(row[4]))
    .map(row => row[4].trim());
    
  console.log("E-mails filtrados com sucesso.");
  return emailsParaEnvio;
}

/* 
  Função auxiliar para obter aniversariantes do mês a partir da aba "Pessoal.Saude".
  Considera que a data de nascimento está na coluna G e o e-mail na coluna J.
  A data é formatada para "dd/MM" (apenas dia e mês).
*/
function obterAniversariantesMes(sheet) {
  const hoje = new Date();
  const mesAtual = hoje.getMonth() + 1;
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let aniversariantesMes = [];
  
  // Começa na linha 2, presumindo que a primeira linha contém cabeçalhos
  for (let i = 1; i < values.length; i++) {
    let nome = values[i][1];
    let dataNascimento = new Date(values[i][6]); // Coluna G
    let email = values[i][9];              // Coluna J
    
    if (!nome || !dataNascimento || isNaN(dataNascimento.getTime()) || !email || !validarEmail(email)) {
      console.warn("Dados inválidos na linha " + (i + 1) + ". Pulando...");
      continue;
    }
    
    let mesAniversario = dataNascimento.getMonth() + 1;
    if (mesAniversario === mesAtual) {
      let dataFormatada = Utilities.formatDate(dataNascimento, Session.getScriptTimeZone(), "dd/MM");
      aniversariantesMes.push({
        nome: nome,
        dataNascimento: dataNascimento,
        dataFormatada: dataFormatada,
        email: email.trim()
      });
    }
  }
  
  // Ordena os aniversariantes pelo dia (extraído da data de nascimento)
  aniversariantesMes.sort((a, b) => a.dataNascimento.getDate() - b.dataNascimento.getDate());
  return aniversariantesMes;
}

/* 
  Função auxiliar para obter aniversariantes que fazem aniversário hoje.
  Compara o dia e o mês da data de nascimento com o dia e mês atuais.
*/
function obterAniversariantesHoje(sheet) {
  const hoje = new Date();
  const diaHoje = hoje.getDate();
  const mesHoje = hoje.getMonth() + 1;
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  let aniversariantesHoje = [];
  
  for (let i = 1; i < values.length; i++) {
    let nome = values[i][1];
    let dataNascimento = new Date(values[i][6]); // Coluna G
    let email = values[i][9];              // Coluna J
    
    if (!nome || !dataNascimento || isNaN(dataNascimento.getTime()) || !email || !validarEmail(email)) {
      console.warn("Dados inválidos na linha " + (i + 1) + ". Pulando...");
      continue;
    }
    
    if (dataNascimento.getDate() === diaHoje && (dataNascimento.getMonth() + 1) === mesHoje) {
      let dataFormatada = Utilities.formatDate(dataNascimento, Session.getScriptTimeZone(), "dd/MM");
      aniversariantesHoje.push({
        nome: nome,
        dataNascimento: dataNascimento,
        dataFormatada: dataFormatada,
        email: email.trim()
      });
    }
  }
  
  return aniversariantesHoje;
}

/* 
  Gera o template HTML para o e-mail semanal com os aniversariantes do mês.
  isTeste indica se é um envio de teste.
*/
function buildEmailBodySemanal(aniversariantes, isTeste) {
  let titulo = isTeste ? "🎉 [TESTE] Aniversariantes do Mês" : "🎉 Aniversariantes do Mês";
  let corpoEmailHtml = `
    <div style="font-family: sans-serif;">
      <h4>Caros colegas,</h4>
      <p>Segue a lista de aniversariantes do mês${isTeste ? " (TESTE)" : ""}:</p>
      <ul>
  `;
  
  aniversariantes.forEach(aniversariante => {
    corpoEmailHtml += `
      <li>
        <strong>Nome:</strong> ${aniversariante.nome}<br>
        <strong>Data de Aniversário:</strong> ${aniversariante.dataFormatada}<br>
        <strong>E-mail:</strong> <a href="mailto:${aniversariante.email}">${aniversariante.email}</a>
      </li>
    `;
  });
  
  corpoEmailHtml += `
      </ul>
      <p>Por favor, enviem suas felicitações! 🎈🎂</p>
      <p>Atenciosamente,<br>
         Equipe do Gabinete do Senador Astronauta Marcos Pontes<br><br>
         Desenvolvido por Daniel Marques (<a href="mailto:${EMAIL_ADMIN}">${EMAIL_ADMIN}</a>)
      </p>
    </div>
  `;
  
  return { titulo: titulo, corpo: corpoEmailHtml };
}

/* 
  Gera o template HTML para o e-mail diário de aniversário, com mensagem personalizada.
*/
function buildEmailBodyDiario(aniversariante) {
  let titulo = `Feliz Aniversário, ${aniversariante.nome}!`;
  let corpoEmailHtml = `
    <div style="font-family: sans-serif;">
      <h4>Olá ${aniversariante.nome},</h4>
      <p>Hoje é um dia especial: <strong>${aniversariante.dataFormatada}</strong> - o dia do seu aniversário!</p>
      <p>Que seu dia seja repleto de alegria e felicidades.</p>
      <p>Atenciosamente,<br>
         Equipe do Gabinete do Senador Astronauta Marcos Pontes<br><br>
         Desenvolvido por Daniel Marques (<a href="mailto:${EMAIL_ADMIN}">${EMAIL_ADMIN}</a>)
      </p>
    </div>
  `;
  return { titulo: titulo, corpo: corpoEmailHtml };
}

/* 
  Função para enviar o e-mail semanal com os aniversariantes do mês.
  Este envio ocorre apenas nas segundas-feiras.
*/
function enviarEmailAniversariantesSemanal() {
  console.log("Iniciando o processo de envio semanal de e-mails.");
  const hoje = new Date();
  
  // Verifica se hoje é segunda-feira (getDay() retorna 1 para segunda-feira)
  if (hoje.getDay() !== 1) {
    console.log('Hoje não é segunda-feira. Nenhum e-mail semanal será enviado.');
    return;
  }
  
  const sheetPessoal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SAUDE);
  if (!sheetPessoal) {
    console.error("Não foi possível encontrar a aba '" + SHEET_SAUDE + "'. Verifique o nome da aba.");
    return;
  }
  
  const aniversariantesMes = obterAniversariantesMes(sheetPessoal);
  
  if (aniversariantesMes.length > 0) {
    enviaEmailSemanal(aniversariantesMes);
  } else {
    console.log('Não há aniversariantes no mês atual.');
  }
}

/* 
  Envia o e-mail semanal para o e-mail do administrador.
*/
function enviaEmailSemanal(aniversariantesMes) {
  const { titulo, corpo } = buildEmailBodySemanal(aniversariantesMes, false);
  try {
    MailApp.sendEmail({
      to: EMAIL_ADMIN,
      subject: titulo,
      htmlBody: corpo
    });
    console.log('E-mail semanal enviado com sucesso para ' + EMAIL_ADMIN + '.');
  } catch (error) {
    console.error("Erro ao enviar e-mail semanal: " + error.message);
  }
}

/* 
  Função para enviar, diariamente, e-mails de felicitações aos aniversariantes do dia.
*/
function enviarEmailDiarioAniversariantes() {
  console.log("Iniciando o processo diário de envio de e-mails de aniversário.");
  const sheetPessoal = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SAUDE);
  if (!sheetPessoal) {
    console.error("Não foi possível encontrar a aba '" + SHEET_SAUDE + "'. Verifique o nome da aba.");
    return;
  }
  
  const aniversariantesHoje = obterAniversariantesHoje(sheetPessoal);
  
  if (aniversariantesHoje.length > 0) {
    aniversariantesHoje.forEach(aniversariante => {
      enviaEmailDiario(aniversariante);
    });
  } else {
    console.log("Nenhum aniversariante encontrado para hoje.");
  }
}

/* 
  Envia um e-mail diário de felicitações para o aniversariante.
*/
function enviaEmailDiario(aniversariante) {
  const { titulo, corpo } = buildEmailBodyDiario(aniversariante);
  try {
    MailApp.sendEmail({
      to: aniversariante.email,
      subject: titulo,
      htmlBody: corpo
    });
    console.log('E-mail de aniversário enviado com sucesso para ' + aniversariante.email);
  } catch (error) {
    console.error("Erro ao enviar e-mail diário para " + aniversariante.email + ": " + error.message);
  }
}

/* 
  Cria um trigger semanal para o envio automático do e-mail de lembrete (toda segunda-feira).
*/
function criarTriggerSemanal() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove triggers duplicados
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === 'enviarEmailAniversariantesSemanal') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  try {
    ScriptApp.newTrigger('enviarEmailAniversariantesSemanal')
             .timeBased()
             .everyWeeks(1)
             .onWeekDay(ScriptApp.WeekDay.MONDAY)
             .atHour(8)
             .create();
    console.log('Trigger semanal criado com sucesso.');
  } catch (error) {
    console.error("Erro ao criar trigger semanal: " + error.message);
  }
}

/* 
  Cria um trigger diário para verificar e enviar e-mails de aniversário.
*/
function criarTriggerDiario() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove triggers duplicados para o envio diário
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === 'enviarEmailDiarioAniversariantes') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  try {
    ScriptApp.newTrigger('enviarEmailDiarioAniversariantes')
             .timeBased()
             .everyDays(1)
             .atHour(8)
             .create();
    console.log('Trigger diário criado com sucesso.');
  } catch (error) {
    console.error("Erro ao criar trigger diário: " + error.message);
  }
}

/* 
  Função de configuração inicial que cria ambos os triggers: semanal e diário.
*/
function configurar() {
  criarTriggerSemanal();
  criarTriggerDiario();
  console.log('Configuração inicial finalizada com sucesso.');
}
