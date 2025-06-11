const ABA_CADASTRO = "Cadastro de Animais";
const ABA_HISTORICO = "Histórico Médico";
const ABA_CONFIGURACOES = "Configurações";
const COLUNA_ID = 1; // Coluna do ID na aba de Cadastro
let animalIdParaMedicacao; // Variável para armazenar o ID do animal

function getListaMedicacoes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaConfiguracoes = ss.getSheetByName(ABA_CONFIGURACOES);
  if (abaConfiguracoes) {
    var colunaMedicacoes = abaConfiguracoes.getRange("A:A").getValues().flat().filter(String);
    colunaMedicacoes.shift(); // Remove o cabeçalho, se houver
    return colunaMedicacoes;
  }
  return [];
}

function gerarProximoId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCadastro = ss.getSheetByName(ABA_CADASTRO);
  var ids = abaCadastro.getRange(2, COLUNA_ID, abaCadastro.getLastRow() - 1).getValues().flat().filter(String); // Pega todos os IDs existentes

  if (ids.length === 0) {
    return "001";
  }

  var ultimoIdStr = ids[ids.length - 1];
  var ultimoIdNum = parseInt(ultimoIdStr, 10);

  if (isNaN(ultimoIdNum)) {
    // Se o último ID não for um número, começamos do 1
    return "001";
  }

  var proximoIdNum = ultimoIdNum + 1;
  return proximoIdNum.toString().padStart(3, '0'); // Formata para 3 dígitos com zeros à esquerda
}

function mostrarFormularioCadastro() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormularioCadastroAnimal')
      .setWidth(600) // Ajuste a largura conforme necessário
      .setHeight(500); // Ajuste a altura conforme necessário
  ui.showModalDialog(htmlOutput, 'Cadastrar Novo Animal');
}

function cadastrarNovoAnimalDoFormulario(nome, especie, raca, sexo, dataNascimento, cor, peso, tutorNome, tutorTelefone, tutorEmail, observacoes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCadastro = ss.getSheetByName(ABA_CADASTRO);
  var id = gerarProximoId(); // Gera o próximo ID sequencial
  var novaLinha = [id, nome, especie, raca, sexo, dataNascimento, cor, peso, tutorNome, tutorTelefone, tutorEmail, observacoes];
  abaCadastro.appendRow(novaLinha);
  SpreadsheetApp.getUi().alert('Animal "' + nome + '" cadastrado com ID: ' + id);
}

function cadastrarNovoAnimal(nome) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCadastro = ss.getSheetByName(ABA_CADASTRO);
  var id = gerarProximoId(); // Gera o próximo ID sequencial
  var novaLinha = [id, nome, "", "", "", "", "", "", "", "", "", ""]; // Preenche com dados básicos
  abaCadastro.appendRow(novaLinha);
  SpreadsheetApp.getUi().alert('Animal "' + nome + '" cadastrado com ID: ' + id);
}

function mostrarFormularioConsulta() {
  var ui = SpreadsheetApp.getUi();
  var resultado = ui.prompt(
      'Consultar Animal',
      'Digite o ID do animal para consulta:',
      ui.ButtonSet.OK_CANCEL);

  if (resultado.getSelectedButton() == ui.Button.OK) {
    var idAnimal = resultado.getResponseText();
    if (idAnimal) {
      consultarAnimal(idAnimal);
    }
  }
}

function consultarAnimal(idAnimal) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaCadastro = ss.getSheetByName(ABA_CADASTRO);
  var abaHistorico = ss.getSheetByName(ABA_HISTORICO);
  var dataCadastro = abaCadastro.getDataRange().getValues();
  var animalEncontrado = null;
  var linhaCadastro = -1;

  // Busca o animal pelo ID
  for (var i = 1; i < dataCadastro.length; i++) {
    if (dataCadastro[i][COLUNA_ID - 1] == idAnimal) {
      animalEncontrado = dataCadastro[i];
      linhaCadastro = i + 1;
      break;
    }
  }

  if (animalEncontrado) {
    var mensagem = "--- Ficha do Animal ---\n";
    var cabecalhoCadastro = abaCadastro.getRange(1, 1, 1, abaCadastro.getLastColumn()).getValues()[0];
    for (var j = 0; j < cabecalhoCadastro.length; j++) {
      mensagem += cabecalhoCadastro[j] + ": " + animalEncontrado[j] + "\n";
    }

    mensagem += "\n--- Histórico Médico ---\n";
    var cabecalhoHistorico = abaHistorico.getRange(1, 1, 1, abaHistorico.getLastColumn()).getValues()[0];
    var dataHistorico = abaHistorico.getDataRange().getValues();
    var historicoAnimal = [];
    var administracoesMedicacao = [];
    var outrosEventos = [];

    for (var k = 1; k < dataHistorico.length; k++) {
      if (dataHistorico[k][0] == idAnimal) {
        var evento = {};
        for (var l = 0; l < cabecalhoHistorico.length; l++) {
          evento[cabecalhoHistorico[l]] = dataHistorico[k][l];
        }
        if (evento['Tipo de Evento (Cadastro, Internação, Alta, Medicação, Consulta, etc.)'] === 'Medicação (Admin)') {
          administracoesMedicacao.push(evento);
        } else {
          outrosEventos.push(evento);
        }
      }
    }

    if (outrosEventos.length > 0) {
      mensagem += "\n--- Outros Eventos ---\n";
      outrosEventos.forEach(function(evento) {
        for (var chave in evento) {
          mensagem += chave + ": " + evento[chave] + " | ";
        }
        mensagem += "\n";
      });
    }

    if (administracoesMedicacao.length > 0) {
      mensagem += "\n--- Medicações Administradas ---\n";
      administracoesMedicacao.forEach(function(evento) {
        mensagem += "Data/Hora: " + Utilities.formatDate(new Date(evento['Data e Hora do Evento']), Session.getTimeZone(), "dd/MM/yyyy HH:mm") + " | ";
        mensagem += "Medicação: " + evento['Medicação (se aplicável)'] + " | ";
        mensagem += "Dose: " + evento['Dose (se aplicável)'] + " | ";
        mensagem += "Via: " + evento['Via de Administração (se aplicável)'] + " | ";
        if (evento['Observações']) {
          mensagem += "Observações: " + evento['Observações'] + " | ";
        }
        mensagem += "\n";
      });
    } else {
      mensagem += "\nNenhuma medicação administrada registrada para este animal.\n";
    }

    SpreadsheetApp.getUi().alert('Informações do Animal (ID: ' + idAnimal + ')', mensagem, SpreadsheetApp.getUi().ButtonSet.OK);

  } else {
    SpreadsheetApp.getUi().alert('Erro', 'Animal com ID "' + idAnimal + '" não encontrado.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function solicitarIdParaMedicacao() {
  var ui = SpreadsheetApp.getUi();
  var resultadoId = ui.prompt(
      'Registrar Medicação',
      'Digite o ID do animal:',
      ui.ButtonSet.OK_CANCEL);
  if (resultadoId.getSelectedButton() == ui.Button.OK) {
    animalIdParaMedicacao = resultadoId.getResponseText();
    if (animalIdParaMedicacao) {
      mostrarDialogoMedicacao(animalIdParaMedicacao); // Passa o ID para a função do diálogo
    }
  }
}

function mostrarDialogoMedicacao(animalId) {
  var ui = SpreadsheetApp.getUi();
  var listaMedicacoes = getListaMedicacoes();
  var htmlTemplate = HtmlService.createTemplateFromFile('FormularioMedicacao');
  htmlTemplate.listaMedicacoes = listaMedicacoes;
  htmlTemplate.animalId = animalId; // Passa o ID para o template
  var htmlOutput = htmlTemplate.evaluate()
      .setWidth(300)
      .setHeight(200);
  ui.showModalDialog(htmlOutput, 'Registrar Medicação');
}

function registrarMedicacaoDoFormulario(animalId, medicacao, dose, via) { // Recebe o animalId corretamente
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaHistorico = ss.getSheetByName(ABA_HISTORICO);
  var dataHora = new Date();
  var responsavel = Session.getActiveUser().getEmail();
  var novaLinha = [animalId, dataHora, "Medicação", medicacao, dose, via, responsavel, ""];
  abaHistorico.appendRow(novaLinha);
  SpreadsheetApp.getUi().alert('Medicação "' + medicacao + '" registrada para o animal com ID: ' + animalId);
}

function registrarInternacao(idAnimal, ambulatorio) {
  registrarEventoHistorico(idAnimal, "Internação", "", "", "", "Ambulatório: " + ambulatorio);
}

function registrarAlta(idAnimal, ambulatorio) {
  registrarEventoHistorico(idAnimal, "Alta", "", "", "", "Ambulatório: " + ambulatorio);
}

function registrarConsulta(idAnimal, observacoes, ambulatorio) {
  registrarEventoHistorico(idAnimal, "Consulta", "", "", "", "Ambulatório: " + ambulatorio + (observacoes ? " | " + observacoes : ""));
}

function registrarEventoHistorico(idAnimal, tipoEvento, medicacao, dose, via, observacoes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaHistorico = ss.getSheetByName(ABA_HISTORICO);
  var dataHora = new Date();
  var responsavel = Session.getActiveUser().getEmail();
  var novaLinha = [idAnimal, dataHora, tipoEvento, medicacao || "", dose || "", via || "", responsavel, observacoes || ""];
  abaHistorico.appendRow(novaLinha);
  SpreadsheetApp.getUi().alert('Evento "' + tipoEvento + '" registrado para o animal com ID: ' + idAnimal);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Gestão Clínica')
      .addItem('Cadastrar Novo Animal', 'mostrarFormularioCadastro')
      .addItem('Consultar Animal', 'mostrarFormularioConsulta')
      .addSubMenu(ui.createMenu('Registrar Evento')
          .addItem('Medicação', 'solicitarIdParaMedicacao')
          .addItem('Internação', 'solicitarIdParaInternacao')
          .addItem('Alta', 'solicitarIdParaAlta')
          .addItem('Consulta', 'solicitarIdParaConsulta')
          .addItem('Administrar Medicação (Internação)', 'solicitarIdParaAdministrarMedicacao')) // Novo item de menu
      .addToUi();
}

function solicitarIdParaAdministrarMedicacao() {
  var ui = SpreadsheetApp.getUi();
  var resultadoId = ui.prompt(
      'Administrar Medicação',
      'Digite o ID do animal:',
      ui.ButtonSet.OK_CANCEL);
  if (resultadoId.getSelectedButton() == ui.Button.OK) {
    animalIdParaAdministrarMedicacao = resultadoId.getResponseText();
    if (animalIdParaAdministrarMedicacao) {
      mostrarFormularioAdministrarMedicacao(animalIdParaAdministrarMedicacao);
    }
  }
}


function mostrarFormularioAdministrarMedicacao(animalId) {
  var ui = SpreadsheetApp.getUi();
  var listaMedicacoes = getListaMedicacoes();
  var htmlTemplate = HtmlService.createTemplateFromFile('FormularioAdministrarMedicacao');
  htmlTemplate.listaMedicacoes = listaMedicacoes;
  htmlTemplate.animalId = animalId;
  var htmlOutput = htmlTemplate.evaluate()
      .setWidth(400)
      .setHeight(400);
  ui.showModalDialog(htmlOutput, 'Administrar Medicação');
}


function registrarAdministracaoMedicacao(animalId, dataHora, medicacao, dose, via, observacoes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaHistorico = ss.getSheetByName(ABA_HISTORICO);
  var responsavel = Session.getActiveUser().getEmail();
  var novaLinha = [animalId, dataHora, "Medicação (Admin)", medicacao, dose, via, responsavel, observacoes || ""];
  abaHistorico.appendRow(novaLinha);
  SpreadsheetApp.getUi().alert('Medicação "' + medicacao + '" administrada para o animal com ID: ' + animalId + ' em ' + Utilities.formatDate(new Date(dataHora), Session.getTimeZone(), "dd/MM/yyyy HH:mm"));
}




function solicitarIdParaInternacao() {
  solicitarIdComAmbulatorio('Registrar Internação', 'registrarInternacao');
}

function solicitarIdParaAlta() {
  solicitarIdComAmbulatorio('Registrar Alta', 'registrarAlta');
}

function solicitarIdParaConsulta() {
  var ui = SpreadsheetApp.getUi();
  var htmlOutput = HtmlService.createHtmlOutputFromFile('FormularioConsulta')
      .setWidth(400) // Ajuste a largura conforme necessário
      .setHeight(350); // Ajuste a altura conforme necessário
  ui.showModalDialog(htmlOutput, 'Registrar Consulta');
}

function registrarConsultaDoFormulario(animalId, ambulatorio, observacoes) {
  registrarEventoHistorico(animalId, "Consulta", "", "", "", "Ambulatório: " + ambulatorio + (observacoes ? " | " + observacoes : ""));
  SpreadsheetApp.getUi().alert('Consulta registrada para o animal com ID: ' + animalId);
}


function solicitarIdParaAcao(titulo, funcao) {
  var ui = SpreadsheetApp.getUi();
  var resultado = ui.prompt(
      titulo,
      'Digite o ID do animal:',
      ui.ButtonSet.OK_CANCEL);
  if (resultado.getSelectedButton() == ui.Button.OK) {
    var idAnimal = resultado.getResponseText();
    if (idAnimal) {
      this[funcao](idAnimal); // Chama a função dinamicamente pelo nome
    }
  }
}

function solicitarIdComAmbulatorio(titulo, funcao) {
  var ui = SpreadsheetApp.getUi();
  var resultadoId = ui.prompt(
      titulo,
      'Digite o ID do animal:',
      ui.ButtonSet.OK_CANCEL);
  if (resultadoId.getSelectedButton() == ui.Button.OK) {
    var idAnimal = resultadoId.getResponseText();
    if (idAnimal) {
      var resultadoAmbulatorio = ui.prompt(
          titulo,
          'Selecione o ambulatório (1 ou 2):',
          ui.ButtonSet.OK_CANCEL);
      if (resultadoAmbulatorio.getSelectedButton() == ui.Button.OK) {
        var ambulatorio = resultadoAmbulatorio.getResponseText();
        if (ambulatorio === '1' || ambulatorio === '2') {
          this[funcao](idAnimal, ambulatorio);
        } else if (ambulatorio) {
          ui.alert('Aviso', 'Ambulatório inválido. Digite 1 ou 2.', ui.ButtonSet.OK);
        } else {
          this[funcao](idAnimal, ""); // Sem ambulatório
        }
      }
    }
  }
}