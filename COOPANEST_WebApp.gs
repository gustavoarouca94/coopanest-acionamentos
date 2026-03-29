// ============================================================
// COOPANEST-PE — Web App + Sistema de Alertas de Atestado
// ============================================================
// INSTALAÇÃO:
// 1. Cole no Apps Script da planilha COOPANEST — Acionamentos 2026
// 2. Implantar > Nova implantação > Web App
//    - Executar como: Eu mesmo
//    - Quem pode acessar: Qualquer pessoa
// 3. Copie a URL gerada e cole no index.html em APPS_SCRIPT_URL
// 4. Execute configurarGatilhosTemporais() UMA VEZ para ativar alertas
// ============================================================

const CFG = {
  EMAIL_COORD_EXECUTIVO: "coordenador@coopanestpe.com.br",
  EMAIL_COMISSAO:        "atestados@coopanestpe.com.br",
  EMAIL_COMITE:          "comite@coopanestpe.com.br",
  FUSO:                  "America/Recife",
  ABA_ACIONAMENTOS:      "Acionamentos",

  // Índices das colunas (base 1)
  COL_PROTOCOLO:       1,
  COL_DATAHORA:        2,
  COL_NOME:            3,
  COL_EMAIL:           4,
  COL_HOSPITAL:        5,
  COL_DATA_PLANTAO:    6,
  COL_TURNO:           7,
  COL_MOTIVO:          8,
  COL_DURACAO:         9,
  COL_EXIGE_REP:       10,
  COL_DATA_REP:        11,
  COL_TURNO_REP:       12,
  COL_TEM_ATESTADO:    13,
  COL_STATUS_ATESTADO: 14,
  COL_OBS:             15,
  COL_CATEGORIA:       16,
  COL_PRAZO_48H:       17,
  COL_ALERTA_ENVIADO:  18,
  COL_ESCALADO:        19,
};

// ─────────────────────────────────────────────
// RECEBE GET — busca protocolo para preview
// ─────────────────────────────────────────────
function doGet(e) {
  const action    = e.parameter.action || "";
  const protocolo = (e.parameter.protocolo || "").toUpperCase();

  if (action === "buscarProtocolo" && protocolo) {
    const resultado = buscarDadosProtocolo(protocolo);
    return jsonResponse(resultado);
  }

  if (action === "listarAcionamentos") {
    return jsonResponse(listarAcionamentos());
  }

  return jsonResponse({ ok: true, msg: "COOPANEST API online" });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function buscarDadosProtocolo(protocolo) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba) return { encontrado: false };

  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][CFG.COL_PROTOCOLO - 1] === protocolo) {
      const prazo = dados[i][CFG.COL_PRAZO_48H - 1];
      return {
        encontrado:     true,
        dados: {
          protocolo,
          nome:           dados[i][CFG.COL_NOME - 1],
          hospital:       dados[i][CFG.COL_HOSPITAL - 1],
          data:           dados[i][CFG.COL_DATA_PLANTAO - 1],
          turno:          dados[i][CFG.COL_TURNO - 1],
          prazo:          prazo || null,
          statusAtestado: dados[i][CFG.COL_STATUS_ATESTADO - 1]
        }
      };
    }
  }
  return { encontrado: false };
}

// ─────────────────────────────────────────────
// RECEBE ATESTADO POSTERIOR (enviado pela página enviar-atestado.html)
// ─────────────────────────────────────────────
function receberAtestadoPosterior(d) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba) return { ok: false, erro: "Aba não encontrada" };

  const dados = aba.getDataRange().getValues();
  let linhaEncontrada = -1;
  let nomeOriginal    = "";

  for (let i = 1; i < dados.length; i++) {
    if (dados[i][CFG.COL_PROTOCOLO - 1] === d.protocolo) {
      linhaEncontrada = i + 1;
      nomeOriginal    = dados[i][CFG.COL_NOME - 1];
      break;
    }
  }

  if (linhaEncontrada === -1) return { ok: false, erro: "Protocolo não encontrado" };

  // Salva o arquivo no Drive
  let pasta = obterOuCriarPasta("COOPANEST — Atestados Médicos");
  const blob = Utilities.newBlob(
    Utilities.base64Decode(d.base64Arquivo),
    d.tipoArquivo,
    `${d.protocolo}_${d.nomeArquivo}`
  );
  const arquivo = pasta.createFile(blob);
  arquivo.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
  const linkArquivo = arquivo.getUrl();

  // Atualiza a planilha
  aba.getRange(linhaEncontrada, CFG.COL_TEM_ATESTADO).setValue("SIM");
  aba.getRange(linhaEncontrada, CFG.COL_STATUS_ATESTADO).setValue("Recebido — posterior");
  aba.getRange(linhaEncontrada, CFG.COL_PRAZO_48H).setValue("—");
  aba.getRange(linhaEncontrada, CFG.COL_ALERTA_ENVIADO).setValue("SIM");
  aba.getRange(linhaEncontrada, 1, 1, aba.getLastColumn()).setBackground("#f0fff7");

  // E-mail para a Comissão com o arquivo
  MailApp.sendEmail({
    to: CFG.EMAIL_COMISSAO,
    subject: `📎 Atestado recebido (posterior) | ${nomeOriginal} | ${d.protocolo}`,
    body:
      `A Comissão de Plantões recebeu o atestado médico enviado posteriormente ao acionamento.\n\n` +
      `Protocolo:   ${d.protocolo}\n` +
      `Plantonista: ${nomeOriginal}\n` +
      `Enviado por: ${d.nome} (${d.email})\n` +
      `Recebido em: ${d.dataHora}\n` +
      `Arquivo:     ${d.nomeArquivo}\n` +
      (d.obs ? `Observações: ${d.obs}\n` : "") +
      `\nLink do arquivo: ${linkArquivo}\n\n` +
      `Os alertas automáticos foram interrompidos para este protocolo.\n\n` +
      `Sistema de Acionamentos — COOPANEST-PE`,
    attachments: [arquivo.getAs(d.tipoArquivo)]
  });

  // Confirmação para o plantonista
  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `✅ COOPANEST — Atestado recebido | Protocolo ${d.protocolo}`,
      body:
        `Olá, ${d.nome},\n\n` +
        `Seu atestado médico foi recebido e vinculado ao protocolo ${d.protocolo}.\n\n` +
        `Arquivo: ${d.nomeArquivo}\n` +
        `Recebido em: ${d.dataHora}\n\n` +
        `A Comissão de Plantões analisará o documento e entrará em contato caso necessário.\n\n` +
        `Atenciosamente,\nSistema de Acionamentos — COOPANEST-PE`
    });
  }

  return { ok: true, protocolo: d.protocolo };
}

// ─────────────────────────────────────────────
// RECEBE POST DO FORMULÁRIO WEB (atualizado para rotear ações)
// ─────────────────────────────────────────────
function doPost(e) {
  try {
    const dados = JSON.parse(e.postData.contents);

    // Roteamento por ação
    if (dados.action === "confirmarAtestado") {
      return jsonResponse(confirmarAtestadoComissao(dados.protocolo));
    }

    if (dados.action === "confirmarReposicao") {
      return jsonResponse(confirmarReposicaoComissao(dados.protocolo));
    }

    if (dados.action === "receberAtestado") {
      const resultado = receberAtestadoPosterior(dados);
      return ContentService
        .createTextOutput(JSON.stringify(resultado))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (dados.action === "registrarTurno") {
      const resultado = registrarTurno(dados);
      return ContentService
        .createTextOutput(JSON.stringify(resultado))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Acionamento padrão (sem action = formulário principal)
    gravarNaPlanilha(dados);
    enviarEmailsIniciais(dados);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, protocolo: dados.protocolo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Erro doPost: " + err);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, erro: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function obterOuCriarPasta(nome) {
  const pastas = DriveApp.getFoldersByName(nome);
  return pastas.hasNext() ? pastas.next() : DriveApp.createFolder(nome);
}

// ─────────────────────────────────────────────
// GRAVA NA PLANILHA
// ─────────────────────────────────────────────
function gravarNaPlanilha(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let aba  = ss.getSheetByName(CFG.ABA_ACIONAMENTOS);

  if (!aba) {
    aba = ss.insertSheet(CFG.ABA_ACIONAMENTOS);
    const cab = [
      "Protocolo","Data/Hora","Nome","E-mail","Hospital",
      "Data Plantão","Turno","Motivo","Duração Atestado",
      "Exige Reposição","Data Reposição","Turno Reposição",
      "Atestado Enviado?","Status Atestado","Observações",
      "Categoria","Prazo 24h (ISO)","Alerta Enviado","Escalado Comitê"
    ];
    aba.getRange(1,1,1,cab.length).setValues([cab])
      .setFontWeight("bold")
      .setBackground("#0a5c3a")
      .setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  const isAtestado = d.motivo === "Doença/Atestado Médico";
  const exigeRep   = d.duracaoAtestado === "lt72";
  const prazo24h   = isAtestado && !d.temAtestado
    ? new Date(new Date().getTime() + 24 * 60 * 60 * 1000).toISOString()
    : "";

  const turnoRepLabel = d.turnoReposicao === "diurno"  ? "Diurno (07h–19h)"
    : d.turnoReposicao === "noturno" ? "Noturno (19h–07h)" : "—";

  const linha = [
    d.protocolo, d.dataHora, d.nome, d.email, d.hospital,
    d.dataPlantao ? formatarData(d.dataPlantao) : "",
    d.turno, d.motivo,
    d.duracaoAtestado === "lt72" ? "< 72h" : d.duracaoAtestado === "gte72" ? "≥ 72h" : "N/A",
    exigeRep ? "SIM" : (MOTIVOS_SEM_REPOSICAO.includes(d.motivo) ? "Não" : "Verificar"),
    d.dataReposicao ? formatarData(d.dataReposicao) : "—",
    turnoRepLabel,
    d.temAtestado ? "SIM" : "NÃO",
    isAtestado ? (d.temAtestado ? "Recebido" : "Pendente") : "N/A",
    d.obs || "—",
    classificarMotivo(d.motivo),
    prazo24h,
    "Não", "Não"
  ];

  const prox = aba.getLastRow() + 1;
  aba.getRange(prox, 1, 1, linha.length).setValues([linha]);

  if (isAtestado && !d.temAtestado)     aba.getRange(prox,1,1,linha.length).setBackground("#fff4ee");
  else if (isAtestado && d.temAtestado) aba.getRange(prox,1,1,linha.length).setBackground("#f0fff7");
  else if (d.motivo === "Afastamento disciplinar") aba.getRange(prox,1,1,linha.length).setBackground("#fef2f2");
}

// ─────────────────────────────────────────────
// E-MAILS INICIAIS
// ─────────────────────────────────────────────
function enviarEmailsIniciais(d) {
  const isAtestado    = d.motivo === "Doença/Atestado Médico";
  const isDisc        = d.motivo === "Afastamento disciplinar";
  const exigeRep      = d.duracaoAtestado === "lt72";
  const turnoRepLabel = d.turnoReposicao === "diurno" ? "Diurno (07h–19h)"
    : d.turnoReposicao === "noturno" ? "Noturno (19h–07h)" : "—";
  const prazoFmt      = formatarDataHoraISO(new Date(new Date().getTime() + 24*60*60*1000));

  const resumo =
    `Protocolo:     ${d.protocolo}\n` +
    `Registrado em: ${d.dataHora}\n` +
    `────────────────────────────────────\n` +
    `Plantonista:   ${d.nome} (${d.email})\n` +
    `Hospital:      ${d.hospital}\n` +
    `Plantão:       ${formatarData(d.dataPlantao)} — ${d.turno}\n` +
    `Motivo:        ${d.motivo}\n` +
    (isAtestado ? `Duração:       ${d.duracaoAtestado === "lt72" ? "< 72h" : "≥ 72h"}\n` : "") +
    `Exige rep.:    ${exigeRep ? "SIM" : "Não"}\n` +
    (exigeRep ? `Reposição:     ${formatarData(d.dataReposicao)} — ${turnoRepLabel}\n` : "") +
    (isAtestado ? `Atestado:      ${d.temAtestado ? "Enviado no ato" : "PENDENTE"}\n` : "") +
    (d.obs ? `Observações:   ${d.obs}\n` : "") +
    `────────────────────────────────────\n` +
    `COOPANEST-PE — Sistema de Acionamentos`;

  // Coordenador executivo
  MailApp.sendEmail({
    to: CFG.EMAIL_COORD_EXECUTIVO,
    subject: `🔔 Acionamento ${d.protocolo} | ${d.nome} | ${formatarData(d.dataPlantao)} ${d.turno}`,
    body: resumo
  });

  // Comissão (atestado médico ou disciplinar)
  if (isAtestado || isDisc) {
    const tipo = isDisc ? "🔴 AFASTAMENTO DISCIPLINAR" : "📋 ATESTADO MÉDICO";
    MailApp.sendEmail({
      to: CFG.EMAIL_COMISSAO,
      subject: `${tipo} | ${d.nome} | ${formatarData(d.dataPlantao)} | ${d.protocolo}`,
      body: `Acionamento para atenção da Comissão de Plantões.\n\n` + resumo +
        (isAtestado && !d.temAtestado
          ? `\n\n⚠️ Atestado NÃO foi enviado no acionamento.\n` +
            `Prazo para recebimento: ${prazoFmt} (24h a partir do acionamento).\n` +
            `Caso não recebido no prazo, o sistema escalará ao Comitê de Integridade automaticamente.`
          : "")
    });
  }

  // Aviso ao próprio plantonista sobre pendência
  if (isAtestado && !d.temAtestado && d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `⏱ COOPANEST — Atestado pendente | Protocolo ${d.protocolo}`,
      body:
        `Olá, ${d.nome},\n\n` +
        `Seu acionamento foi registrado (protocolo ${d.protocolo}).\n\n` +
        `O atestado médico do plantão de ${formatarData(d.dataPlantao)} não foi anexado.\n\n` +
        `⚠️ Você tem até ${prazoFmt} para encaminhar o documento para:\n` +
        `${CFG.EMAIL_COMISSAO}\n\n` +
        `Assunto: Atestado — ${d.protocolo}\n\n` +
        `Após esse prazo, sem recebimento, o caso será encaminhado ao Comitê de Integridade e Ética.\n\n` +
        `Atenciosamente,\nSistema de Acionamentos — COOPANEST-PE`
    });
  }
}

// ─────────────────────────────────────────────
// VERIFICAÇÃO PERIÓDICA (a cada hora via gatilho)
// ─────────────────────────────────────────────
function verificarPendenciasAtestado() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const aba = ss.getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba || aba.getLastRow() < 2) return;

  const dados = aba.getDataRange().getValues();
  const agora = new Date();

  for (let i = 1; i < dados.length; i++) {
    const l = dados[i];
    const motivo        = l[CFG.COL_MOTIVO - 1];
    const temAtestado   = l[CFG.COL_TEM_ATESTADO - 1];
    const prazo24hStr   = l[CFG.COL_PRAZO_48H - 1];
    const alertaEnviado = l[CFG.COL_ALERTA_ENVIADO - 1];
    const escalado      = l[CFG.COL_ESCALADO - 1];
    const email         = l[CFG.COL_EMAIL - 1];
    const nome          = l[CFG.COL_NOME - 1];
    const protocolo     = l[CFG.COL_PROTOCOLO - 1];
    const dataPlantao   = l[CFG.COL_DATA_PLANTAO - 1];

    if (motivo !== "Doença/Atestado Médico") continue;
    if (temAtestado === "SIM") continue;
    if (!prazo24hStr || prazo24hStr === "—") continue;

    const prazo = new Date(prazo24hStr);
    const d = { nome, email, protocolo, dataPlantao };

    if (agora > prazo) {
      // Prazo vencido — escala se ainda não escalou
      if (escalado !== "SIM") {
        escalarParaComite(d);
        aba.getRange(i+1, CFG.COL_ESCALADO).setValue("SIM");
        aba.getRange(i+1, CFG.COL_STATUS_ATESTADO).setValue("Vencido — escalado");
        aba.getRange(i+1, 1, 1, aba.getLastColumn()).setBackground("#fce4e4");
      }
    } else {
      // Dentro do prazo — lembrete se faltam ≤ 12h e ainda não enviou
      const horasRestantes = (prazo - agora) / 1000 / 3600;
      if (horasRestantes <= 12 && alertaEnviado !== "SIM") {
        enviarLembreteAtestado(d, Math.round(horasRestantes));
        aba.getRange(i+1, CFG.COL_ALERTA_ENVIADO).setValue("SIM");
      }
    }
  }
}

// ─────────────────────────────────────────────
// LEMBRETE 12H ANTES DO VENCIMENTO
// ─────────────────────────────────────────────
function enviarLembreteAtestado(d, horas) {
  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `⚠️ COOPANEST — Últimas ${horas}h para envio do atestado | ${d.protocolo}`,
      body:
        `Olá, ${d.nome},\n\n` +
        `Restam aproximadamente ${horas} hora(s) para enviar o atestado médico ` +
        `do protocolo ${d.protocolo} (plantão ${d.dataPlantao}).\n\n` +
        `Encaminhe para: ${CFG.EMAIL_COMISSAO}\n` +
        `Assunto: Atestado — ${d.protocolo}\n\n` +
        `Após o prazo, o caso será encaminhado ao Comitê de Integridade e Ética.\n\n` +
        `Atenciosamente,\nSistema de Acionamentos — COOPANEST-PE`
    });
  }

  MailApp.sendEmail({
    to: CFG.EMAIL_COMISSAO,
    subject: `⚠️ Prazo vencendo em ${horas}h | ${d.nome} | ${d.protocolo}`,
    body:
      `A Comissão está sendo notificada: prazo de atestado vencendo em ~${horas}h.\n\n` +
      `Plantonista: ${d.nome}\nProtocolo: ${d.protocolo}\nPlantão: ${d.dataPlantao}\n\n` +
      `Se o atestado não for recebido, o caso será escalado automaticamente ao Comitê.`
  });
}

// ─────────────────────────────────────────────
// ESCALAÇÃO AO COMITÊ DE INTEGRIDADE E ÉTICA
// ─────────────────────────────────────────────
function escalarParaComite(d) {
  const agora = Utilities.formatDate(new Date(), CFG.FUSO, "dd/MM/yyyy 'às' HH:mm");

  MailApp.sendEmail({
    to: CFG.EMAIL_COMITE,
    cc: CFG.EMAIL_COMISSAO,
    subject: `🔴 ESCALAÇÃO INSTITUCIONAL | Atestado não recebido | ${d.nome} | ${d.protocolo}`,
    body:
      `Prezado Comitê de Integridade e Ética,\n\n` +
      `O sistema de acionamentos identificou um caso para avaliação institucional.\n\n` +
      `SITUAÇÃO: Acionamento por Doença/Atestado Médico sem envio do atestado ` +
      `comprobatório no prazo de 24 horas, mesmo após notificações automáticas.\n\n` +
      `──────────────────────────────────────\n` +
      `Plantonista:  ${d.nome}\n` +
      `E-mail:       ${d.email || "não informado"}\n` +
      `Protocolo:    ${d.protocolo}\n` +
      `Data plantão: ${d.dataPlantao}\n` +
      `Escalado em:  ${agora}\n` +
      `──────────────────────────────────────\n\n` +
      `Recomenda-se avaliação quanto às medidas institucionais previstas no regulamento da COOPANEST-PE.\n\n` +
      `Sistema de Acionamentos — COOPANEST-PE`
  });

  if (d.email) {
    MailApp.sendEmail({
      to: d.email,
      subject: `🔴 COOPANEST — Prazo encerrado | Caso encaminhado ao Comitê | ${d.protocolo}`,
      body:
        `Olá, ${d.nome},\n\n` +
        `O prazo de 24 horas para envio do atestado do protocolo ${d.protocolo} foi encerrado ` +
        `sem que o documento fosse recebido.\n\n` +
        `Seu caso foi encaminhado ao Comitê de Integridade e Ética para avaliação.\n\n` +
        `Se já enviou o atestado ou tem alguma justificativa, entre em contato imediatamente:\n` +
        `${CFG.EMAIL_COMISSAO}\n\n` +
        `Atenciosamente,\nSistema de Acionamentos — COOPANEST-PE`
    });
  }

  Logger.log("Escalado ao Comitê: " + d.protocolo + " — " + d.nome);
}

// ─────────────────────────────────────────────
// REGISTRAR TURNO COMPLETO (Módulo 2)
// ─────────────────────────────────────────────
function registrarTurno(d) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let aba   = ss.getSheetByName("Turnos");

  if (!aba) {
    aba = ss.insertSheet("Turnos");
    const cab = [
      "Data/Hora Fechamento","Data Turno","Tipo Turno","Coordenador","Reservas Escalados",
      "Nº Acionamento","Protocolo Vinculado","Tipo Registro","Plantonista Saiu",
      "Hospital","Motivo","Reserva Acionado","Status","Horário Acionamento",
      "Observações Acionamento","Observações Gerais do Turno"
    ];
    aba.getRange(1,1,1,cab.length).setValues([cab])
      .setFontWeight("bold").setBackground("#073d28").setFontColor("#ffffff");
    aba.setFrozenRows(1);
  }

  const obsTurno = d.obsTurno || "—";

  if (d.acionamentos.length === 0) {
    // Turno sem acionamentos — registra uma linha resumo
    aba.appendRow([
      d.dataHoraFecho, d.turno.data, d.turno.tipo, d.turno.coord,
      d.turno.reservas, 0, "—", "—", "—", "—", "—", "—", "—", "—", "—", obsTurno
    ]);
  } else {
    d.acionamentos.forEach((a, i) => {
      const linha = [
        d.dataHoraFecho, d.turno.data, d.turno.tipo, d.turno.coord,
        d.turno.reservas, i + 1,
        a.protocolo || "—",
        a.tipo === "formulario" ? "Formulário" : "Manual",
        a.plantonistaSaiu, a.hospital, a.motivo, a.reservaAcionado,
        a.status === "confirmado" ? "Confirmado" : "Recusado",
        a.horario, a.obs || "—", i === 0 ? obsTurno : "—"
      ];
      aba.appendRow(linha);

      // Destaque por status
      const ul = aba.getLastRow();
      if (a.status === "recusado") {
        aba.getRange(ul, 1, 1, linha.length).setBackground("#fef2f2");
      } else if (a.tipo === "formulario") {
        aba.getRange(ul, 1, 1, linha.length).setBackground("#f0fff7");
      }
    });
  }

  // Notifica Coordenador Executivo com resumo
  const resumoLinhas = d.acionamentos.map((a, i) =>
    `${i+1}. [${a.status === "confirmado" ? "✓" : "✕"}] ${a.plantonistaSaiu} → ${a.reservaAcionado} | ${a.hospital} | ${a.motivo}`
  ).join("\n");

  MailApp.sendEmail({
    to: CFG.EMAIL_COORD_EXECUTIVO,
    subject: `📋 Relatório de Turno | ${d.turno.tipo} | ${d.turno.data} | ${d.turno.coord}`,
    body:
      `Turno encerrado e relatório enviado.\n\n` +
      `────────────────────────────────────\n` +
      `Coordenador: ${d.turno.coord}\n` +
      `Data/Turno:  ${d.turno.data} — ${d.turno.tipo}\n` +
      `Fechamento:  ${d.dataHoraFecho}\n` +
      `Reservas escalados: ${d.turno.reservas}\n` +
      `────────────────────────────────────\n` +
      `Total de acionamentos: ${d.totalAcionamentos}\n` +
      `Confirmados: ${d.totalConfirmados} | Recusados: ${d.totalRecusados}\n` +
      `Via formulário: ${d.totalFormulario} | Manuais: ${d.totalManual}\n` +
      `────────────────────────────────────\n` +
      (d.acionamentos.length > 0 ? resumoLinhas + "\n" : "Nenhum acionamento neste turno.\n") +
      `────────────────────────────────────\n` +
      (d.obsTurno ? `Observações: ${d.obsTurno}\n` : "") +
      `\nSistema de Acionamentos — COOPANEST-PE`
  });

  return { ok: true };
}

// ─────────────────────────────────────────────
// LISTAR ACIONAMENTOS (para Comissão e Dashboard)
// ─────────────────────────────────────────────
function listarAcionamentos() {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba || aba.getLastRow() < 2) return { ok: true, dados: [] };

  const linhas = aba.getDataRange().getValues();
  const dados  = [];
  for (let i = 1; i < linhas.length; i++) {
    const l = linhas[i];
    dados.push({
      protocolo:      l[CFG.COL_PROTOCOLO - 1],
      nome:           l[CFG.COL_NOME - 1],
      email:          l[CFG.COL_EMAIL - 1],
      hospital:       l[CFG.COL_HOSPITAL - 1],
      dataPlantao:    l[CFG.COL_DATA_PLANTAO - 1],
      turno:          l[CFG.COL_TURNO - 1],
      motivo:         l[CFG.COL_MOTIVO - 1],
      duracao:        l[CFG.COL_DURACAO - 1],
      exigeRep:       l[CFG.COL_EXIGE_REP - 1] === "SIM",
      dataRep:        l[CFG.COL_DATA_REP - 1],
      turnoRep:       l[CFG.COL_TURNO_REP - 1],
      statusAtestado: l[CFG.COL_STATUS_ATESTADO - 1],
      statusRep:      l[CFG.COL_EXIGE_REP - 1] === "SIM" ? (l[CFG.COL_DATA_REP - 1] ? "Pendente" : "N/A") : "N/A",
      prazo24h:       l[CFG.COL_PRAZO_48H - 1] || "",
      obs:            l[CFG.COL_OBS - 1],
      categoria:      l[CFG.COL_CATEGORIA - 1],
    });
  }
  return { ok: true, dados };
}

// ─────────────────────────────────────────────
// CONFIRMAR ATESTADO (Comissão)
// ─────────────────────────────────────────────
function confirmarAtestadoComissao(protocolo) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba) return { ok: false };
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][CFG.COL_PROTOCOLO - 1] === protocolo) {
      aba.getRange(i+1, CFG.COL_TEM_ATESTADO).setValue("SIM");
      aba.getRange(i+1, CFG.COL_STATUS_ATESTADO).setValue("Recebido");
      aba.getRange(i+1, CFG.COL_PRAZO_48H).setValue("—");
      aba.getRange(i+1, 1, 1, aba.getLastColumn()).setBackground("#f0fff7");
      const email = dados[i][CFG.COL_EMAIL - 1];
      const nome  = dados[i][CFG.COL_NOME - 1];
      if (email) MailApp.sendEmail({
        to: email,
        subject: `✅ COOPANEST — Atestado validado | ${protocolo}`,
        body: `Olá, ${nome},\n\nSeu atestado médico (protocolo ${protocolo}) foi validado pela Comissão de Plantões.\n\nComissão de Plantões — COOPANEST-PE`
      });
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─────────────────────────────────────────────
// CONFIRMAR REPOSIÇÃO CUMPRIDA (Comissão)
// ─────────────────────────────────────────────
function confirmarReposicaoComissao(protocolo) {
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba) return { ok: false };
  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][CFG.COL_PROTOCOLO - 1] === protocolo) {
      let cabecalho = aba.getRange(1,1,1,aba.getLastColumn()).getValues()[0];
      let colRep = cabecalho.indexOf("Reposição Cumprida") + 1;
      if (colRep === 0) {
        colRep = aba.getLastColumn() + 1;
        aba.getRange(1,colRep).setValue("Reposição Cumprida").setFontWeight("bold").setBackground("#0a5c3a").setFontColor("#ffffff");
      }
      aba.getRange(i+1, colRep).setValue("SIM");
      const email   = dados[i][CFG.COL_EMAIL - 1];
      const nome    = dados[i][CFG.COL_NOME - 1];
      const dataRep = dados[i][CFG.COL_DATA_REP - 1];
      if (email) MailApp.sendEmail({
        to: email,
        subject: `✅ COOPANEST — Reposição confirmada | ${protocolo}`,
        body: `Olá, ${nome},\n\nSua reposição (protocolo ${protocolo}, data: ${dataRep}) foi confirmada pela Comissão.\n\nComissão de Plantões — COOPANEST-PE`
      });
      return { ok: true };
    }
  }
  return { ok: false };
}

// ─────────────────────────────────────────────
function marcarAtestadoRecebido() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt(
    "Marcar atestado como recebido",
    "Informe o número do protocolo (ex: AC-20260320-143052):",
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const protocolo = resp.getResponseText().trim().toUpperCase();
  const aba = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.ABA_ACIONAMENTOS);
  if (!aba) { ui.alert("Aba não encontrada."); return; }

  const dados = aba.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][CFG.COL_PROTOCOLO - 1] === protocolo) {
      aba.getRange(i+1, CFG.COL_TEM_ATESTADO).setValue("SIM");
      aba.getRange(i+1, CFG.COL_STATUS_ATESTADO).setValue("Recebido");
      aba.getRange(i+1, CFG.COL_PRAZO_48H).setValue("—");
      aba.getRange(i+1, 1, 1, aba.getLastColumn()).setBackground("#f0fff7");
      ui.alert("✅ Protocolo " + protocolo + " marcado como atestado recebido.");
      return;
    }
  }
  ui.alert("Protocolo " + protocolo + " não encontrado.");
}

// ─────────────────────────────────────────────
// MENU NA PLANILHA
// ─────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("COOPANEST")
    .addItem("✅ Marcar atestado como recebido", "marcarAtestadoRecebido")
    .addItem("🔍 Verificar pendências agora",    "verificarPendenciasAtestado")
    .addSeparator()
    .addItem("⚙️ Ativar alertas automáticos",   "configurarGatilhosTemporais")
    .addToUi();
}

// ─────────────────────────────────────────────
// CONFIGURAR GATILHO HORÁRIO
// Execute UMA VEZ após publicar o Web App
// ─────────────────────────────────────────────
function configurarGatilhosTemporais() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("verificarPendenciasAtestado")
    .timeBased().everyHours(1).create();
  Logger.log("✅ Gatilho horário ativado.");
  SpreadsheetApp.getUi().alert(
    "✅ Sistema de alertas ativado!\n\nVerificação de pendências: a cada hora, automaticamente."
  );
}

// ─────────────────────────────────────────────
// AUXILIARES
// ─────────────────────────────────────────────
const MOTIVOS_SEM_REPOSICAO = [
  "Licença-maternidade","Licença-paternidade",
  "Óbito — cônjuge/pai/mãe/filho/irmão","Óbito — avô/avó/neto/sogro",
  "Casamento do plantonista","Intimação judicial/policial/posse",
  "Convocação CA/CF/Comitê de Conformidade","Saída definitiva de cooperado",
  "Nova vaga contratual","Afastamento disciplinar","Férias de funcionário do hospital"
];

function formatarData(str) {
  if (!str) return "—";
  const p = String(str).split("-");
  return p.length === 3 ? `${p[2]}/${p[1]}/${p[0]}` : str;
}

function formatarDataHoraISO(d) {
  const pad = x => String(x).padStart(2,"0");
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} às ${pad(d.getHours())}:${pad(d.getMinutes())}h`;
}

function classificarMotivo(motivo) {
  const mapa = {
    "Doença/Atestado Médico":"Saúde — AM","Internamento hospitalar":"Saúde — Internamento",
    "Cirurgia não estética":"Saúde — Cirurgia","Afastamento temporário por saúde":"Saúde — Afastamento",
    "Licença-maternidade":"Licença Legal","Licença-paternidade":"Licença Legal",
    "Doença grave de familiar":"Familiar","Acompanhar cirurgia de familiar":"Familiar",
    "Óbito — cônjuge/pai/mãe/filho/irmão":"Óbito 1º grau","Óbito — avô/avó/neto/sogro":"Óbito 2º grau",
    "Casamento do plantonista":"Evento Pessoal","Episódio social violento (com B.O.)":"Evento Pessoal",
    "Incompatibilidade — novo concurso":"Concurso","Intimação judicial/policial/posse":"Convocação Legal",
    "Dobra não coberta pela escala":"Operacional","Convocação CA/CF/Comitê de Conformidade":"Institucional",
    "Saída definitiva de cooperado":"Desligamento","Nova vaga contratual":"Operacional",
    "Aprimoramento científico (mín. 30 dias)":"Científico",
    "Devolução de plantão voluntário (aviso prévio)":"Voluntário",
    "Afastamento disciplinar":"Disciplinar","Férias de funcionário do hospital":"Operacional"
  };
  return mapa[motivo] || "Outros";
}
