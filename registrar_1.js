"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var path = require('path');
const fs = require('fs');
let listautuacao = [];
var instancia_1;
let dist_dif = [];
let lista_pfn = [];
let lista_opcoes = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `Índice do Registro - ${dia}-${mes}-${ano}.xlsx`;

const juizo_1inst = ['', '1ª Instância']

const classes_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (informações do débito pendente)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)',
    'Execução Fiscal (FNDE)', 'Cumprimento de Sentença', 'Ação Trabalhista', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cumprimento de Sentença contra a Fazenda Pública', 'Arrolamento', 
    'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Carta Precatória', 'Cautelar', 'Cautelar Fiscal', 'Consignação em Pagamento', 'Desapropriação', 'Embargos de Terceiro', 
    'Execução de Título Extrajudicial', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Exibição de Documento ou Coisa', 'Falência', 'Habilitação', 'Inventário', 'Outras', 'Procedimento Comum',    
    'Procedimento do Juizado Especial Cível', 'Protesto', 'Petição', 'Incidente de Desconsideração de Personalidade Jurídica', 'Representação', 'Restauração de Autos', 'Restituição de Coisa ou Dinheiro na Falência', 
    'Reintegração / Manutenção de Posse', 'Reclamação', 'Recuperação Judicial', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Usucapião', 'Ação de Exigir Contas'];

const classes_defesa = ['Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Mandado de Segurança', 'Cumprimento de Sentença', 'Mandado de Segurança Coletivo', 'Cumprimento de Sentença contra a Fazenda Pública', 
    'Cumprimento Provisório de Sentença', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Execução de Título Extrajudicial', 'Execução de Título Extrajudicial contra a Fazenda Pública', 'Petição', 
    'Outras', 'Embargos de Terceiro', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cautelar', 'Cautelar Fiscal', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Restauração de Autos',  
    'Monitória', 'Notificação', 'Habilitação', 'Habeas Data', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Ação Penal', 'Ação Popular', 'Alvará/Outros Procedimentos Jurisdição Voluntária',
    'Habeas Corpus', 'Liquidação por Arbitramento', 'Liquidação Provisória por Arbitramento', 'Liquidação Provisória pelo Procedimento Comum', 'Liquidação pelo Procedimento Comum', 'Liquidação Provisória de Sentença', 
    'Reintegração / Manutenção de Posse', 'Retificação de Registro de Imóvel', 'Restituição de Coisas Apreendidas', 'Exibição de Documento ou Coisa', 'Produção Antecipada da Prova', 'Procedimento Sumário', 
    'Protesto', 'Representação', 'Oposição', 'Depósito', 'Restituição de Coisa ou Dinheiro na Falência', 'Ação de Demarcação', 'Ação de Exigir Contas', 'Consignação em Pagamento', 'Depósito da Lei 8.866/94', 
    'Desapropriação', 'Despejo', 'Dissolução e Liquidação de Sociedade', 'Mandado de Injunção', 'Recuperação Judicial', 'Embargos à Adjudicação', 'Embargos à Arrematação', 'Impugnação de Assistência Judiciária', 
    'Incidente de Suspeição', 'Impugnação de Crédito', 'Impugnação ao Pedido de Assist Litiscon Simples', 'Impugnação Ao Cumprimento de Sentença', 'Incidente de Impedimento', 'Incidente de Falsidade', 
    'Incidente de Desconsideração de Personalidade Jurídica', 'Pedido de Quebra Sigilo de Dados e/ou Telefônico', 'Insolvência Requerida pelo Devedor ou pelo Espólio', 'Dissolução e Liquidação de Sociedade', 
    'Dissolução Parcial de Sociedade', 'Exceção de Incompetência', 'Exceção de Litispendência', 'Falência', 'Embargos Infringentes na Execução Fiscal', 'Impugnação ao Valor da Causa', 'Carta Precatória', 
    'Carta Rogatória', 'Carta de Sentença', 'Carta de Ordem', 'Arrolamento', 'Inventário', 'Usucapião', 'Ação Trabalhista', 'Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FNDE)', 
    'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];

     //################# FUNÇÃO PARA ROLAR A TELA PARA BAIXO
let rolar_tela = async (page) => {
    await page.evaluate(() => { //rolar a tela para baixo
      const heightPage = document.body.scrollHeight;
      window.scrollTo(0 , heightPage);
    }); 
};

let writeexcel_lista = async (arq) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Relatório");
    worksheet.columns = [
        {header: 'Status', key: 'status', width: 10},
        {header: 'Triador', key: 'triador', width: 30},
        {header: 'Tipo Petição', key: 'tipo', width: 25},
        {header: 'Descrição', key: 'descricao', width: 30},
        {header: 'Processo', key: 'processo', width: 30}
    ];
    for (let lin = 1; lin < 11; lin++) {worksheet.getColumn(lin).font = {name: 'Times New Roman', size: 10};}
    worksheet.getRow(1).font = {name: 'Calibri', size: 11, bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listautuacao.length; i++){
        let lista = await listautuacao[i].dados.map(el => el.numero).join(',');
        worksheet.addRow({status: listautuacao[i].status, triador: listautuacao[i].triador, tipo: listautuacao[i].tipo_peticao,  descricao: listautuacao[i].descricao, processo: lista});
    }  
    await wbook.xlsx.writeFile(arq);
}


let dados_inicio = async (page) => {
    var filePath = path.resolve(__dirname, 'Dados.xlsx');
    if (fs.existsSync(filePath)) { 
        var wb = new Excel.Workbook();
        await wb.xlsx.readFile(filePath);
        let worksheet = wb.getWorksheet('SAJ');
        let linha_ind = await worksheet.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
        await page.waitForSelector('input[id="frmLogin:username"]');
        await page.type('input[id="frmLogin:username"]', linha_ind[1]);
        await page.type('input[id="frmLogin:password"]', linha_ind[2]);
        await page.click('button[id="frmLogin:entrar"]');
    }
};

let inclui_dados_procurador_distribuicao = async (page, tria) => {
    await page.focus('div[id="formManifestacao:signatarios:selectOneMenu"]'); //Combobox do signatário recebe o foco
    await console.log('RECEBEU O FOCO');
    const indice_signatario = await indice_opcao(page, 'div[id$="signatarios:selectOneMenu_panel"] ul li', tria)-1;//Lista em um array todas as opções do campo signatário
    await console.log('ACHOU O INDICE: '+indice_signatario);
    await page.waitForFunction(() => document.querySelector('div[id="formManifestacao:signatarios:selectOneMenu"]').focus);//Aguarda o foco
    await console.log('ACHADO E COM FOCO');
    await page.click('select[name="formManifestacao:signatarios:selectOneMenu_input"]');
    await console.log('CLICADO');
    await page.waitForSelector('div[id$="signatarios:selectOneMenu_panel"] ul li', {visible: true});
    await console.log('VENDO');
    await page.click(`div[id$="signatarios:selectOneMenu_panel"] ul li:nth-child(${indice_signatario})`);
    await console.log('CLICOU');
    //await select_evaluete(page, `div[id*="procurador:selectOneMenu_panel"] > ul > li:nth-child(${indice_signatario})`, 'select[name$="procurador:selectOneMenu_input"]') //Clica no nome do Procurador signatário
    await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${tria}')]`);//Aguarda o campo signatário ser alterado
    await console.log('PASSOU');
    await new Promise(r => setTimeout(r, 12500));
};

let nova_linha = async (page, linha_anterior) => {
    await page.waitForSelector(`tbody[id="formManifestacao:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(5)`);
    let img_coluna; 
    for (var i = 1; i <= 3500; i++) {
        img_coluna = await indicadores_coluna(page, linha_anterior);
        if (img_coluna > 0) {
            i= 4000;
        }
    }
    if (img_coluna > 0) {
        await page.waitForSelector(`tbody[id="formManifestacao:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(5) div[class="ui-dt-c"] div[id*="pgIndProc"] div[style="display:inline-block"]`, {visible: true});//Aguarda a coluna 5 formar (tornar visível todos os elementos)
    }
}

let indicadores_coluna = async (page, linha_anterior) => {
    let painel_caixas_disponiveis = await page.evaluate((linha_anterior)=>{
        let allcaixas_span = Array.from(document.querySelectorAll(`tbody[id="formManifestacao:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(5) div[class="ui-dt-c"] div[id*="pgIndProc"] div`))
        let allcaixas_a_array = allcaixas_span.map(i=>{return i.id}) // retorna o id de todas as jurisdições que constam no painel 
        return allcaixas_a_array
    }, linha_anterior) // encerra localizar NOME e ID das caixas disponíveis
    return painel_caixas_disponiveis.length
}

let indice_opcao = async (page, seletor_opcoes, valor) => {
    const opcoes = await page.evaluate((seletor_opcoes) => Array.from(document.querySelectorAll(seletor_opcoes)).map((el)=>{return el.innerText}), seletor_opcoes);
    const indice = await opcoes.findIndex(element => element === valor);
    await console.log('SOMANDO 1 ______________________________ '+opcoes[indice]+' Indice: '+ indice);
    await console.log('SEM SOMAR______________________________ '+opcoes[indice-1]+' Indice: '+ indice-1)
    return indice
};

let click_botao = async (page, texto) => {
    const botao = await page.$x(`//span[contains(text(), '${texto}')]`); //Procura o botão pelo texto, dependendo do argumento 
    botao[0].click(); //Clicar no botão limpar ou voltar, dependendo do argumento 
};

let slavar_indice = async (arq, registro) => {
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Índice_Autuação");
    worksheet.columns = [
        {header: 'Registro', key: 'autuacao', width: 25}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    await worksheet.addRow({autuacao: registro});
    await wbook.xlsx.writeFile(arq);
};

let apagar_indices = async () => {
    let filePath = await path.resolve(__dirname, nome_arquivo_excel);
    let filePath_lista_registro = await path.resolve(__dirname, 'Lista para Registro.xlsx');
    if (fs.existsSync(filePath)) { 
        try {
            fs.unlinkSync(filePath)
          } catch(err) {console.error(err)}
    }
    if (fs.existsSync(filePath_lista_registro)) { 
        try {
            fs.unlinkSync(filePath_lista_registro)
          } catch(err) {console.error(err)}
    }
};

const click_evaluete = async (page, seletor_acao, seletor_evento) => {
    await new Promise(r => setTimeout(r, 500));
    await page.evaluate((seletor_acao, seletor_evento) => {
        document.querySelector(seletor_acao).selected = true;
        document.querySelector(seletor_acao).click();
        element = document.querySelector(seletor_evento);
        var event = new Event('change', { bubbles: true });
        event.simulated=true;
        element.dispatchEvent(event);
    }, seletor_acao, seletor_evento);
    await new Promise(r => setTimeout(r, 1000)); 
}

let indices_pesquisa = async (arg_acao, tipo) => {
    let indice;
    for (let i = 0; i < arg_acao.length; i++) {
        if (tipo.includes(arg_acao[i])) {
            indice = i;
            i = arg_acao.length + 1;
        }
    }
    if (indice === undefined) {
        indice = -1;
    }
    return indice
};

let pagina_autuacao = async (page) => {
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/tarefas/registrarAtuacaoProcessual.jsf?', {waitUntil: 'networkidle0'}); //AMBIENTE DE PRODUÇÃO
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/pages/tarefas/registrarAtuacaoProcessual.jsf?', {waitUntil: 'networkidle0'}); //AMBIENTE DE TESTE
}

let dados_gerais = async (page, idc) => {
    const alteraautuacao = await page.$('table[id$="tiposManifestacao"] tbody tr td div div input[id*="tiposManifestacao:0"]'); //Seletor de autuação
    await alteraautuacao.click(); //Clica no tipo de autuação           
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    const pfn_abertura = await page.$eval('label[id="formManifestacao:procuradoria:selectOneMenu_label"]', el => el.innerText);
    if (await pfn_abertura !== 'PRFN5 (Sede)') {
        const itens_unidade = await page.evaluate(() => Array.from(document.querySelectorAll('div[id$="procuradoria:selectOneMenu_panel"] ul li')).map((el)=>{return el.innerText}));//Retorna todas as opções de unidade em um array
        let indice_procuradoria = await itens_unidade.findIndex(element => element === 'PRFN5 (Sede)'); //Retorna o índice da Procuradoria (PRFN5 (Sede)) do signatário
        indice_procuradoria = indice_procuradoria+1;
        await page.focus('div[id$="procuradoria:selectOneMenu_panel"]'); //Campo da seleção de procuradoria recebe o foco
        await page.waitForFunction(() => document.querySelector('div[id$="procuradoria:selectOneMenu_panel"]').focus); //Aguarda o foco na Procuradoria do signatário
        await click_evaluete(page, `div[id*="procuradoria:selectOneMenu_panel"] > ul > li:nth-child(${indice_procuradoria})`,'select[name="formManifestacao:procuradoria:selectOneMenu_input"]' ) //Clica na Procuradoria do signatár
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    }
    if (await lista_pfn.length < 1) {
        lista_pfn = await page.evaluate(() => Array.from(document.querySelectorAll('select[id="formManifestacao:signatarios:selectOneMenu_input"] option')).map((el)=>{return el.innerText}));
        lista_opcoes = await page.evaluate(() => Array.from(document.querySelectorAll('select[id="formManifestacao:signatarios:selectOneMenu_input"] option')).map((el)=>{return el.value}));
    }
    const indice = await lista_pfn.findIndex(element => element === listautuacao[idc].triador);
    await page.select('select[id="formManifestacao:signatarios:selectOneMenu_input"]', lista_opcoes[indice]);
    const seletipoautuacao = await page.$('table[id$="tiposManifestacao"] tbody tr td div div input[id*="tiposManifestacao:1"]'); //Seletor de autuação
    await seletipoautuacao.click(); //Clica no tipo de autuação           
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    //await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${listautuacao[idc].triador}')]`);//Aguarda o campo signatário ser alterado
    await page.waitForFunction(`document.querySelector("label[id='formManifestacao:signatarios:selectOneMenu_label']").innerText === "${listautuacao[idc].triador}"`);
    await page.click('tbody tr td div span[id$="autoCompleteSaj"] input[id$="autoCompleteSaj_input"]');//Clica no combo da manifestação
    await page.type('input[id$="formManifestacao:peticao:autoCompleteSaj_hinput"]', listautuacao[idc].tipo_peticao);//Escreve a manifestação
    await page.waitForFunction(`document.querySelector("div[id='formManifestacao:peticao:autoCompleteSaj_panel'] ul li[data-item-label='${listautuacao[idc].tipo_peticao}']")`);//Aguarda somente a manifestação escrita está listada
    await page.click(`div[id='formManifestacao:peticao:autoCompleteSaj_panel'] ul li[data-item-label='${listautuacao[idc].tipo_peticao}']`);//Clica na manifestação
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.$eval('tbody tr td div[id$="inDescricao:input"] input[id$="inDescricao:descricao"]', (el, value) => el.value = value, listautuacao[idc].descricao); //Escrever a descrição no campo
    await page.waitForFunction(`document.querySelector("input[id='formManifestacao:inDescricao:descricao']").value.includes("${listautuacao[idc].descricao}")`);
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    //await new Promise(r => setTimeout(r, 700));
};

let click_modal_escolha_classe = async (page, indice_classe) => {
    await new Promise(r => setTimeout(r, 700));
    await page.evaluate((indice_classe) => {document.querySelector(`#formManifestacao\\:modalListaProcessos\\:tabelaModalProcessosJudiciais_data > tr[data-ri="${indice_classe}"] div div div span`).click()}, indice_classe); // ###### Seleciona/click - checkbox - processo da classe listada no imput (planilha origem)
    await new Promise(r => setTimeout(r, 700));
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForSelector(`tbody[id$="tabelaModalProcessosJudiciais_data"] tr[data-ri="${indice_classe}"] td div div div[class="ui-chkbox-box ui-widget ui-corner-all ui-state-default ui-state-active"] span[class="ui-chkbox-icon ui-icon ui-icon-check"]`); // Aguarda o checkbox ficar selecionado
    await page.evaluate(() => document.querySelector('div div[id*="dialogListaProcessos"] div div button[id*="botaoSelecionarProcessoModal"] span').click());// ###### clica no botão "Selecionar"
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
}

    //#######  SE ABRIR O MODAL DO REGISTRO DE AUTUAÇÃO
async function modal_escolha_classe(page, classe_input) {
    await page.waitForSelector('tbody[id$="tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]', {visible: true})
    //let teste = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]')).map((el)=>{return el.innerText}));
    await page.waitForFunction('document.getElementById("formManifestacao:modalListaProcessos:dialogListaProcessos").style.visibility === "visible"');
    //await page.waitForSelector('div[id="formManifestacao:modalListaProcessos:tabelaModalProcessosJudiciais"]', {visible: true}, {timeout: 120000})
    await page.waitForSelector('tbody[id$="tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]', {visible: true}, {timeout: 120000}) //Aguarda a tabela do modal ficar visivel
    let classes_modal = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]')).map((el)=>{return el.innerText}));//Retorna as classes do listadas no modal em um array
    let indice = await classes_modal.findIndex(element => element === classe_input); //compara as classes listadas no modal (array) com a classe da planilha imput e retorna o indice encontrado
    if (await indice !== -1) {    // ###### Se encontrar a classe pesquisada 
        click_modal_escolha_classe(page, indice)
    } else if (await indice == -1) {
        indice = await indices_pesquisa(instancia_1, classes_modal);
        classe_input = await instancia_1[indice];
        indice = await classes_modal.findIndex(element => element === classe_input);
        if (await indice !== -1) {
            click_modal_escolha_classe(page, indice)
        }  else {    // ###### Se não encontrar a classe pesquisada - clica no botão fechar
                await console.log('A classe do SAJ não coincide com a classe da planilha')
                await page.evaluate(() => document.querySelector('div button[id="formManifestacao:modalListaProcessos:botaoFecharProcessoModal"] span').click()); 
        }
    }
    await page.waitForFunction('document.getElementById("formManifestacao:modalListaProcessos:dialogListaProcessos").style.visibility === "hidden"');
}

let ler_indice_autuacao = async (arquivo) => {
    let id_autuacao;
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, arquivo);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Índice_Autuação'); // Primeira aba do arquivo excel - Planilha
        id_autuacao = await parseInt(sh.getRow(2).getCell(1).text);    
    }   else {
        id_autuacao = await 0;
    }
    return id_autuacao;
};

let finalizar_registro = async (page) => { 
    await rolar_tela(page); 
    await page.waitForSelector('button[id="formManifestacao:btnResumo"] span', {visible: true});     
    await click_botao(page, 'Concluir');
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForXPath("//*[@class='ui-messages-info-summary' and contains(., 'Registro de Atuação Processual realizado com sucesso.')]", {timeout: 120000});
}

let dividir_lote = async (page, array, instancia, indice) => {
    let corte_final;
    let acumulado;
    let r = 0;
    let corte_inicial = 0;
    let repeticao = await (Math.ceil(array.length/200));
    await repeticao === 1 ? corte_final = array.length : corte_final = 200;
    do {
        let lote_processo = await array.slice(corte_inicial, corte_final);
        acumulado = await lote_processo.length;
        await incluir_lote(page, lote_processo, instancia, indice);
        corte_inicial = await corte_final;
        await r++;
        await r == (repeticao-1) ? corte_final = await array.length : corte_final = await (corte_final+200);
        if (await acumulado >= 200 || r==repeticao) {
            await finalizar_registro(page);
            dist_dif = [];
            if (r < repeticao) {await pagina_autuacao(page, indice);}
            acumulado == 0;
        }
    } while (r < repeticao);
    await console.log('___________________________________________________________________________________ \n');
}

let incluir_lote = async (page, array, instancia, ind) => {
    let lista = await array.map(el => el.numero).join(',');
    const lista_acabada = await lista.replace(/,+/g, "\n");
    await console.log(lista_acabada);
    await page.waitForSelector('button[id="formManifestacao:btnIncluir_menuButton"] span', {visible: true});//Aguarda o botão incluir ficar visível
    const selelista = await page.$('button[id="formManifestacao:btnIncluir_menuButton"] span');//Passa o seletor do botão incluir
    await selelista.click();//Click no triangulo do botão Incluir
    await page.waitForSelector('div ul li a[id$="menuIncLote"]', {visible: true});//Aguarda o link do incluir em lote
    await page.waitForFunction('document.getElementById("formManifestacao:btnIncluir_menu").style.display === "block"');
    await page.click('div ul li a[id$="menuIncLote"]');//Clica no link para abrir o form para inserir a lista de processos
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');//Aguarda a form de processamento desaparecer
    await page.waitForFunction('document.getElementById("frmModais:modalProcessos:dialogListaEmBloco").style.visibility === "visible"');//Aguarda ficar visível o formulário de inserção
    await page.evaluate((lista_acabada) => {document.querySelector('tbody tr td textarea[id$="listaNrProcessos"]').value = lista_acabada}, lista_acabada);//Cola a lista de processos no form
    if (await instancia == '2ª Instância') {//Caso os processos sejam de 2ª instância, clica para alterar a instância
        await click_evaluete(page, "#frmModais\\:modalProcessos\\:instancias\\:selectOneMenu_panel > ul > li:nth-child(2)", 'select[id="frmModais:modalProcessos:instancias:selectOneMenu_input"]') //Clica no nome do Procurador signatário
        await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${instancia}')]`);
    }  
    await page.evaluate(() => document.querySelector('div button[id$="botaoConfirmaInclusao"]').click());//Confirma a inclusão
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "visible"');
    await page.waitForSelector('tbody tr td label[id$="erroMsg"]', {visible: true});//Aguarda retorno do resultado da inclusão
    let titulo = page.$eval('tbody tr td label[id$="erroMsg"]', el => el.innerText);//Guarda a mensagem na variável: incluídos direto ou algum que precisa escolher classe
    if (await titulo !== 'Todos os processos foram incluídos com sucesso.') {
        let array_proc_classe_dupla = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="gridConcluirInclusaoBloco_data"] tr td:nth-child(1)')).map((el)=>{return el.innerText}));
        let array_entrada = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(4)')).map((el)=>{return el.innerText}));
        await array_entrada.map((texto, indice) => { //Fazer o map - ver quais processos tem + de uma classe
            if(texto === 'Foram encontrados mais de um processo com este número.' || texto === 'Nenhum processo encontrado na 1ª Instância com esse número.'){
                dist_dif.push(array_proc_classe_dupla[indice].replace(/(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})/g,"\$1-\$2.\$3.\$4.\$5.\$6"));
            } 
        })    
    } 

    const seletor_fechar_resultado_lista = await page.$('button[id*="modalListaEmBlocoConcluir:btnFecharConc"]');
    await seletor_fechar_resultado_lista.click();
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "hidden"');
    await page.waitForSelector('tbody[id="frmModais:modalListaEmBlocoConcluir:gridConcluirInclusaoBloco_data"]', {visible: false});
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "hidden"');
    if (await dist_dif.length > 0) {
        let total_list_autuacoes;
        if (await page.$('tbody[id$="formManifestacao:listaInterna:dataTableListaInterna_data"] tr') !== null) {
            total_list_autuacoes =  await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="formManifestacao:listaInterna:dataTableListaInterna_data"] tr')).length);    
        } else {total_list_autuacoes = 0}
        if (await total_list_autuacoes === 0) {total_list_autuacoes = 1}
        for (let nd = 0; nd < await dist_dif.length; nd++) {
            let id_proc = await listautuacao[ind].dados.findIndex( dados => dados.numero === dist_dif[nd] );
            await console.log(listautuacao[ind].dados[id_proc].numero + ' - ' + listautuacao[ind].dados[id_proc].classe)
            await page.$eval('input[id="formManifestacao:numProcesso"]', (el, value) => el.value = value, dist_dif[nd]); //inclui o número do pj
            await page.waitForFunction(`document.getElementById("formManifestacao:numProcesso").value === "${dist_dif[nd]}"`);
            await click_botao(page, 'Incluir');
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            await modal_escolha_classe(page, listautuacao[ind].dados[id_proc].classe);
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            await page.waitForFunction('document.getElementById("formManifestacao:modalListaProcessos:dialogListaProcessos").style.visibility === "hidden"');//Aguarda o modal da escolha de classe ficar invisível
            await new Promise(r => setTimeout(r, 200));
            //if (total_list_autuacoes < 10) {await nova_linha(page, total_list_autuacoes)}
            //await total_list_autuacoes++;
        } 
    }
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await dados_gerais(page, ind);
};

   //########## LER A PLANILHA COM TRIADORES ############### 
let ler_tabela_triadores = async () => {
    let tabela = [];
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname,'Triadores.xlsx');
    await wb.xlsx.readFile(filePath); 
    let sh = await wb.getWorksheet("Triadores");
    sh.eachRow({ includeEmpty: false }, async (row) => {//ler cada uma das linhas da planilha
        await tabela.push({"triador":row.values[1], "plan": row.values[2], "grupo": row.values[3], "unidade": row.values[4], "dados": ''}) //Salva cada linha em um array
    });
    await tabela.shift(); 
    return  tabela;
};

   //########## LER A PLANILHA COM OS PROCESSO A SEREM PESQUISADOS ############### 
let ler_excel = async () => {
    var wb = await new Excel.Workbook();
    let tabela_triadores = await ler_tabela_triadores();
    let triadores_grupo = await tabela_triadores.filter(function(obj) {
        return (obj.grupo == 'X' || obj.grupo == 'x');   
    });
    
    const plan = await triadores_grupo[0].plan; 
    await plan == 'Aux_EXE' ? instancia_1 = await classes_execucao : instancia_1 = await classes_defesa;
    var filePath = await path.resolve(__dirname,'Distribuição.xlsx');
    let fixo = ['Processo', 'Classe judicial', 'Juízo', 'Setor responsável', 'Procurador', 'Tipo de petição', 'Descrição', 'Providência Apoio'];
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet(plan);
        let linha_ind = await sh.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
        let coluna = new Array();  
        await fixo.forEach(function(elemento) {//Define os indices que serão manipulados - quais colunas
            let id_relacionado = linha_ind.findIndex(element => element === elemento);
            if(id_relacionado !== -1) {coluna.push(id_relacionado)}
        });

        let planilha = new Array();
        await sh.eachRow(function(cell, rowNumber) {
            let triador;
            (sh.getRow(rowNumber).getCell(coluna[7]).text === '' || sh.getRow(rowNumber).getCell(coluna[7]).text === 'Já distribuído no SAJ') ? triador = sh.getRow(rowNumber).getCell(coluna[4]).text : triador = sh.getRow(rowNumber).getCell(coluna[7]).text;
            if (sh.getRow(rowNumber).getCell([coluna[3]]).text.substring(0,7) === 'TRIAGEM' && (sh.getRow(rowNumber).getCell(coluna[5]).text !== 'NÃO REGISTRAR' && sh.getRow(rowNumber).getCell(coluna[5]).text !== '') && triadores_grupo.some(t => t.triador.includes(triador))) {
                planilha.push({"numero": sh.getRow(rowNumber).getCell([coluna[0]]).text, "classe": sh.getRow(rowNumber).getCell([coluna[1]]).text, "instancia": sh.getRow(rowNumber).getCell([coluna[2]]).text, "tipo_peticao": sh.getRow(rowNumber).getCell([coluna[5]]).text, "descricao": sh.getRow(rowNumber).getCell([coluna[6]]).text.trim(), "triador": triador})
            }
        }); 
        
        await planilha.map(function(obj) {
            let id = listautuacao.findIndex( arr => arr.triador === obj.triador && arr.tipo_peticao === obj.tipo_peticao && arr.descricao === obj.descricao && arr.arquivo === obj.arquivo && arr.conteudo === obj.conteudo);
            let processo = {"numero": obj.numero, "classe": obj.classe, "instancia": obj.instancia};
            if (id == -1) {
                listautuacao.push({"triador": obj.triador, "tipo_peticao": obj.tipo_peticao, "descricao": obj.descricao, "dados":[processo], "status": 'Pendente'});    
            } else {
                listautuacao[id].dados.push(processo)
            }
        }); 
    }  
    writeexcel_lista('Lista para Registro.xlsx');
};

ler_excel();

let scrape = async () => {
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        headless: false, args:['--start-maximized'], //torna visível e maximiza a 
        ignoreHTTPSErrors: true,
    });
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf?', {waitUntil: 'networkidle0'});//AMBIENTE DE PRODUÇÃO
    await dados_inicio(page);
    await page.waitForNavigation();    
    let id = await ler_indice_autuacao(nome_arquivo_excel);

    console.log("\n");
    let tria = '';
    await console.log('Total de Registro: '+listautuacao.length);
    await console.log('Índice: '+id);
    for (let m = id; m < await listautuacao.length; m++) {
        if (tria !== await listautuacao[m].triador) {
            tria = listautuacao[m].triador;
            await console.log(' *************************************** '+ tria);
        } else {await console.log('   '+ tria);}
        console.log(`           Tipo da Manifestação: ${listautuacao[m].tipo_peticao}`);
        console.log(`              Descrição: ${listautuacao[m].descricao}`);
        await new Promise(r => setTimeout(r, 500));
        dist_dif = await [];
        await console.log('    ###  Total de Processo da descrição: ' + listautuacao[m].dados.length);
        let array_2inst = await listautuacao[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
            return (!juizo_1inst.includes(obj.instancia));
        });
        let array_1inst = await listautuacao[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 1ª instância
            return (juizo_1inst.includes(obj.instancia));
        }); 
        if (array_1inst.length > 0) { 
            await pagina_autuacao(page);
            //await inclui_dados_procurador_distribuicao(page, listautuacao[m].triador);
            await dividir_lote(page, array_1inst, '1ª Instância', m) //Função para inserir os processos de 1ª instância
        }
        if (array_2inst.length > 0) { 
            await pagina_autuacao(page);
            await dividir_lote(page, array_2inst, '2ª Instância', m);
        }
        let id_salvar = await (m+1);
        await slavar_indice(nome_arquivo_excel, id_salvar)
        listautuacao[m].status = 'Registrado';
        writeexcel_lista('Lista para Registro.xlsx');
        await new Promise(r => setTimeout(r, 1000));
    }
    await new Promise(r => setTimeout(r, 1000));
    await browser.close()
    const result = await `Autuação concluída com sucesso. `;
    await apagar_indices();
    return result
}  

scrape().then((value) => {
   console.log(value)
});