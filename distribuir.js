"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
const fs = require('fs');
var path = require('path');
var listdistribuicao = [];
var instancia_1;
let dist_dif = [];
let dist_prevento = [];
let arr_preventos = new Array();
let data = new Date();
let dia = data.getDate();
let mes = data.getMonth() + 1;
const dia_indice = dia;
const mes_indice = mes;
if (dia < 10) {
    dia = '0'+ dia;
}
//let mes = data.getMonth() + 1;
if (mes < 10) {
    mes = '0'+ mes;
}
const juizo_1inst = ['', '1ª Instância'];
const classes_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (informações do débito pendente)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (FNDE)', 
    'Cumprimento de Sentença', 'Ação Trabalhista', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cumprimento de Sentença contra a Fazenda Pública', 'Arrolamento', 'Ação de Improbidade Administrativa', 
    'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Carta Precatória', 'Cautelar', 'Cautelar Fiscal', 'Consignação em Pagamento', 'Desapropriação', 'Embargos de Terceiro', 'Execução de Título Extrajudicial',  
    'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Exibição de Documento ou Coisa', 'Falência', 'Habilitação', 'Incidente de Desconsideração de Personalidade Jurídica', 'Inventário', 'Outras',  
    'Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Protesto', 'Petição', 'Reclamação', 'Recuperação Judicial', 'Representação', 'Restauração de Autos', 'Restituição de Coisa ou Dinheiro na Falência', 
    'Reintegração / Manutenção de Posse', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Usucapião', 'Ação de Exigir Contas'];

const classes_defesa = ['Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Mandado de Segurança', 'Cumprimento de Sentença', 'Mandado de Segurança Coletivo', 'Cumprimento de Sentença contra a Fazenda Pública', 
    'Cumprimento Provisório de Sentença', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Execução de Título Extrajudicial', 'Execução de Título Extrajudicial contra a Fazenda Pública', 'Petição', 'Outras', 
    'Embargos de Terceiro', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cautelar', 'Cautelar Fiscal', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Restauração de Autos', 'Monitória',  
    'Notificação', 'Habilitação', 'Habeas Data', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Ação Penal', 'Ação Popular', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Habeas Corpus', 
    'Liquidação por Arbitramento', 'Liquidação Provisória por Arbitramento', 'Liquidação Provisória pelo Procedimento Comum',  'Liquidação pelo Procedimento Comum', 'Liquidação Provisória de Sentença', 
    'Reintegração / Manutenção de Posse', 'Retificação de Registro de Imóvel', 'Restituição de Coisas Apreendidas', 'Exibição de Documento ou Coisa', 'Produção Antecipada da Prova', 'Procedimento Sumário', 
    'Protesto', 'Representação', 'Oposição', 'Depósito', 'Restituição de Coisa ou Dinheiro na Falência', 'Ação de Demarcação', 'Ação de Exigir Contas', 'Consignação em Pagamento', 'Depósito da Lei 8.866/94', 'Desapropriação',
    'Despejo', 'Dissolução e Liquidação de Sociedade', 'Mandado de Injunção', 'Recuperação Judicial', 'Embargos à Adjudicação', 'Embargos à Arrematação', 'Impugnação de Assistência Judiciária', 'Incidente de Suspeição', 
    'Impugnação de Crédito', 'Impugnação ao Pedido de Assist Litiscon Simples', 'Impugnação Ao Cumprimento de Sentença', 'Incidente de Impedimento', 'Incidente de Falsidade', 'Incidente de Desconsideração de Personalidade Jurídica', 
    'Pedido de Quebra Sigilo de Dados e/ou Telefônico', 'Insolvência Requerida pelo Devedor ou pelo Espólio', 'Dissolução e Liquidação de Sociedade', 'Dissolução Parcial de Sociedade', 'Exceção de Incompetência', 
    'Exceção de Litispendência', 'Falência', 'Embargos Infringentes na Execução Fiscal',  'Impugnação ao Valor da Causa', 'Carta Precatória', 'Carta Rogatória', 'Carta de Sentença', 'Carta de Ordem', 'Arrolamento', 'Inventário',  
    'Usucapião', 'Ação Trabalhista', 'Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FNDE)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];


let data_distribuicao = dia + '/' + mes + '/' + data.getFullYear();
const nome_arquivo_excel = `Índice Distribuição - ${dia_indice}-${mes_indice}-${data.getFullYear()}.xlsx`;

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

let indice_opcao = async (page, seletor_opcoes, valor) => {
    const opcoes = await page.evaluate((seletor_opcoes) => Array.from(document.querySelectorAll(seletor_opcoes)).map((el)=>{return el.innerText}), seletor_opcoes);//Retorna uma array os valores do componente combobox 
    const indice = await opcoes.findIndex(element => element === valor)+1;
    return indice
};

     //################# FUNÇÃO PARA ROLAR A TELA PARA BAIXO
let rolar_tela = async (page) => {
    await page.evaluate(() => { //rolar a tela para baixo
      const heightPage = document.body.scrollHeight;
      window.scrollTo(0 , heightPage);
    }); 
};

        //################# FUNÇÃO PARA VERIFICAR SE UM ELEMENTO EXISTE
let verifica_se_elemento_existe = async (page, caminho) => {
    const elemento = await page.evaluate((caminho) => {
        const el = document.querySelector(caminho);
        //let exist = el.length != 0 ? true : false;
        return el
    }, caminho);
    let exist = await elemento == null ? false : true;
    return exist
};

        //################# FUNÇÃO PARA CLICK EM BOTÃO - PELO TEXTO DO BOTÃO 
let click_botao = async (page, texto) => {
    const botao = await page.$x(`//span[contains(text(), '${texto}')]`); //Procura o botão pelo texto, dependendo do argumento 
    botao[0].click(); //Clicar no botão limpar ou voltar, dependendo do argumento 
};

let nova_linha = async (page, linha_anterior) => {
    await page.waitForSelector(`tbody[id="frmDistProc:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(6)`);
    let img_coluna; 
    for (var i = 1; i <= 3500; i++) {
        img_coluna = await indicadores_coluna(page, linha_anterior);
        if (img_coluna > 0) {
            i= 4000;
        }
    }
    if (img_coluna > 0) {
        await page.waitForSelector (`tbody[id="frmDistProc:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(6) div[class="ui-dt-c"] div div[style="display:inline-block"]`, {visible: true});//Aguarda a coluna 6 formar (tornar visível todos os elementos)
    }
}

let aguarda_data = async (page, d_linha) => {   //função para aguardar todos os processos em uma só página                                                               
    let total_linhas = d_linha+1;
    do {
        let arr_coluna_7 = await page.evaluate(() => Array.from(document.querySelectorAll(`tbody[id="frmDistProc:listaInterna:dataTableListaInterna_data"] tr td:nth-child(7) div span[id*="txtDtEnt"]`)).map((el)=>{return el.innerText}));
        var arraySemVazios = arr_coluna_7.filter(function (i) {//Filtra elementos excluindo os de valor vazio ('')
            return i;
        });
    } while (arraySemVazios.length !== total_linhas); //Repete enquanto o total de elementos diferente do total de linhas, ou seja, enquanto tiver elemento com valor vazio                                                                                                                                                                            
}

let indicadores_coluna = async (page, linha_anterior) => {
    let painel_caixas_disponiveis = await page.evaluate((linha_anterior)=>{
        let allcaixas_span = Array.from(document.querySelectorAll(`tbody[id="frmDistProc:listaInterna:dataTableListaInterna_data"] tr:nth-child(${linha_anterior}) td:nth-child(6) div[class="ui-dt-c"] div[id*="pgIndProc"] div`))
        let allcaixas_a_array = allcaixas_span.map(i=>{return i.id}) // retorna o id de todas as jurisdições que constam no painel 
        return allcaixas_a_array
    }, linha_anterior) // encerra localizar NOME e ID das caixas disponíveis
    return painel_caixas_disponiveis.length
}


let inclui_dados_procurador_distribuicao = async (page, tria) => {
    await page.focus('div[id$="procurador:selectOneMenu_panel"]'); //Combobox do signatário recebe o foco
    const indice_signatario = await indice_opcao(page, 'div[id$="procurador:selectOneMenu_panel"] ul li', tria);//Lista em um array todas as opções do campo signatário
    await page.waitForFunction(() => document.querySelector('div[id$="procurador:selectOneMenu_panel"]').focus);//Aguarda o foco
    await page.click('select[name$="procurador:selectOneMenu_input"]');
    await page.waitForSelector('div[id*="procurador:selectOneMenu_panel"] > ul > li', {visible: true});
    await page.click(`div[id*="procurador:selectOneMenu_panel"] > ul > li:nth-child(${indice_signatario})`);
    //await select_evaluete(page, `div[id*="procurador:selectOneMenu_panel"] > ul > li:nth-child(${indice_signatario})`, 'select[name$="procurador:selectOneMenu_input"]') //Clica no nome do Procurador signatário
    await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${tria}')]`);//Aguarda o campo signatário ser alterado
    await new Promise(r => setTimeout(r, 500));
};

let salva_indice_distribuicao = async (arq, id_triador) => {
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Índice_Triador");
    worksheet.columns = [
        {header: 'Triador', key: 'triador', width: 25}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    await worksheet.addRow({triador: id_triador});
    await wbook.xlsx.writeFile(arq);
};

let apagar_indices_distribuicao = async (arq) => {
    var filePath = await path.resolve(__dirname, arq);
    if (await fs.existsSync(filePath)) { 
        try {
            fs.unlinkSync(filePath);
        } catch(err) {console.error(err)}
    }
};

let ler_indice_distribuicao = async (arq) => {
    let id_distribuicao;
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, arq);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Índice_Triador'); // Primeira aba do arquivo excel - Planilha
        id_distribuicao = await parseInt(sh.getRow(2).getCell(1).text);    
    }   else {  
        id_distribuicao = await 0;
    }
    return id_distribuicao;
};

   //########## LER A PLANILHA COM TRIADORES ############### 
let ler_tabela_triadores = async () => {
    let tabela = [];
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname,'Triadores.xlsx');
    await wb.xlsx.readFile(filePath); 
    let sh = await wb.getWorksheet("Triadores");
    sh.eachRow({ includeEmpty: false }, async (row) => {//ler cada uma das linhas da planilha
        await tabela.push({"triador":row.values[1], "plan": row.values[2], "grupo": row.values[3], "unidade": row.values[4], "dados": []}) //Salva cada linha em um array
    });
    await tabela.shift(); 
    return  tabela;
};

   //########## LER A PLANILHA COM OS PROCESSO A SEREM DISTRIBUÍDOS ############### 
let ler_excel = async () => {
    var wb = await new Excel.Workbook();
    let tabela_triadores = await ler_tabela_triadores();
    let triadores_grupo = await tabela_triadores.filter(function(obj) {
        return (obj.grupo == 'X' || obj.grupo == 'x');   
    });
    
    const plan = await triadores_grupo[0].plan; 
    await plan == 'Aux_EXE' ? instancia_1 = await classes_execucao : instancia_1 = await classes_defesa;
    let fixo = ['Processo', 'Classe judicial', 'Juízo', 'Setor responsável', 'Procurador', 'Providência Apoio'];
    var filePath = await path.resolve(__dirname, 'Distribuição.xlsx');
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet(plan); // nome aba do arquivo excel - Planilha
        let linha_ind = await sh.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
        let coluna = new Array();  
        await fixo.forEach(function(elemento) {//Define os indices que serão manipulados - quais colunas
            let id_relacionado = linha_ind.findIndex(element => element === elemento);
            if(id_relacionado !== -1) {coluna.push(id_relacionado)}
        });
        await sh.eachRow(function(cell, rowNumber) {
            if (sh.getRow(rowNumber).getCell([coluna[3]]).text !== '' && sh.getRow(rowNumber).getCell([coluna[3]]).text.substring(0,7) === 'TRIAGEM' && triadores_grupo.some(t => t.triador.includes(sh.getRow(rowNumber).getCell([coluna[4]]).text))) {
                let triador;
                (sh.getRow(rowNumber).getCell(coluna[5]).text === '' || sh.getRow(rowNumber).getCell(coluna[5]).text === 'Já distribuído no SAJ') ? triador = sh.getRow(rowNumber).getCell(coluna[4]).text : triador = sh.getRow(rowNumber).getCell(coluna[5]).text;
                let processo = {"numero": sh.getRow(rowNumber).getCell([coluna[0]]).text, "classe": sh.getRow(rowNumber).getCell([coluna[1]]).text, "instancia": sh.getRow(rowNumber).getCell([coluna[2]]).text};
                let id = tabela_triadores.findIndex( arr => arr.triador === triador);
                tabela_triadores[id].dados.push(processo)
            }
        });
        
        listdistribuicao = tabela_triadores.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
            return (obj.dados.length > 0);
        });             
    }  
     
};

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

    //####### SE ABRIR FORMULÁRIO COM PREVENTO DEPOIS DE INSERIR O LOTE - SE O PROCESSO JÁ ESTÁ COM OUTRO PROCURADOR - FECHA PARA ESCOLHA UM A UM
async function modal_prevento_lote(page) {
    /*await page.waitForSelector('div button[id*="botaoFecharProcessosOutrosUsuariosModal"] span', {visible: true});
    await page.waitForSelector('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td', {visible: true});
    const barra_pagina = await verifica_se_elemento_existe(page, 'div[id*="tabelaModalProcessosJudiciaisOutrosUsuarios"] td[class="ui-paginator ui-paginator-bottom ui-widget-header"]')//Verifica se existe barra de paginação
    let arr_prev_temp = await new Array();
    arr_prev_temp = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText}));
    await console.log(arr_prev_temp.length + '\n');
    if (barra_pagina == true) { //Se tiver mais de uma página de processos
        const lista_pag = await page.evaluate(() => Array.from(document.querySelectorAll('#frmDistProc\\:mdListProcUsuariosDiferentes\\:tabelaModalProcessosJudiciaisOutrosUsuarios_paginator_bottom > span.ui-paginator-pages > span')).map((el)=>{return el.innerText}));
        let fim_repeticao = await lista_pag[lista_pag.length-1];
        for (let pag = 1; pag < await fim_repeticao; pag++) {
            await page.evaluate(() => document.querySelector('tfoot tr td[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_paginator_bottom"] span[class="ui-paginator-next ui-state-default ui-corner-all"] span[class="ui-icon ui-icon-seek-next"]').click());// Clica no botão next - próxima página
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            let outros_prevento = await new Array();
            outros_prevento = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText}));
            arr_prev_temp = await arr_prev_temp.concat(outros_prevento);
            await console.log(arr_prev_temp.length + '\n');
        }
    }
    for (const proc_temp of arr_prev_temp) {
        const p = await proc_temp.replace(/[\.,\-]/g, "");
        await arr_preventos.push(p);
    }
    //arr_preventos = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText}));
    await page.evaluate(() => document.querySelector('div button[id*="botaoFecharProcessosOutrosUsuariosModal"] span').click());
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');*/
    let pag_repeticao;
    let estado_botao = await '';
    await page.waitForSelector('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td', {visible: true});
    const barra_pagina = await verifica_se_elemento_existe(page, 'div[id*="tabelaModalProcessosJudiciaisOutrosUsuarios"] td[class="ui-paginator ui-paginator-bottom ui-widget-header"]')//Verifica se existe barra de paginação
    let orgao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(5)')).map((el)=>{return el.innerText}));
    let col_setor = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(4)')).map((el)=>{return el.innerText}));
    if (barra_pagina == false && orgao.length > 0){
        for (let n = 0; n < await col_setor.length; n++) {
            if ((col_setor[n] !== 'DIDE2 (Procuradores)') && ((orgao[n] !== 'CASTJ' && orgao[n] !== 'CASTF'))) {  // && ((orgao[n] !== 'CASTJ' && orgao[n] !== 'CASTF'))) {
                await page.evaluate((n) => {document.querySelector(`tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr[data-ri="${n}"] td div div span`).click()}, n);//Clica no check box da linha
                await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
                //dist_prevento = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText}));
                estado_botao = await 'Selecionado';
            }
        }
    } else if (barra_pagina == true) {
        const paginas = await page.evaluate(() => {
            const paginacao = document.querySelector('tfoot tr td[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_paginator_bottom"] span[class="ui-paginator-current"]').innerText; //Verifica a quantidade de páginas do modal
            return paginacao   
        });
        await new Promise(r => setTimeout(r, 1000));
        pag_repeticao = await paginas.substring(6, paginas.length-1); //O número final de páginas para fazer a repetição
        let cont_linha = await 0;
        for (let pag = 0; pag < pag_repeticao; pag++) {
            orgao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(5)')).map((el)=>{return el.innerText}));//Verifica se tem algum de instância diferente
            col_setor = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(4)')).map((el)=>{return el.innerText}));//Verifica se tem algum de segunda instância
            await new Promise(r => setTimeout(r, 500));
            for (let s = 0; s < await col_setor.length; s++){
                if ((col_setor[s] !== 'DIDE2 (Procuradores)') && ((orgao[s] !== 'CASTJ' && orgao[s] !== 'CASTF'))) {
                    await page.evaluate((cont_linha) => {document.querySelector(`tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr[data-ri="${cont_linha}"] td div div span`).click()}, cont_linha); //Clica no checkbox da linha
                    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
                    await new Promise(r => setTimeout(r, 500)); 
                    estado_botao = await 'Selecionado'; 
                } 
                await cont_linha++;
            }
            //let outros_prevento = await new Array();
            //outros_prevento = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText}));
            //dist_prevento = await dist_prevento.concat(outros_prevento);
            let rep = await (pag_repeticao-1)
            if (pag !== rep) {//condição para clicar no botão e ir para próxima pagina - caso exista
                await page.evaluate(() => document.querySelector('tfoot tr td[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_paginator_bottom"] span[class="ui-paginator-next ui-state-default ui-corner-all"] span[class="ui-icon ui-icon-seek-next"]').click());// Clica no botão next - próxima página
                await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            }  
            await new Promise(r => setTimeout(r, 500));           
        }     
    }
    await new Promise(r => setTimeout(r, 1000)); 
    if (estado_botao == 'Selecionado') {
        //await page.waitForXPath(`//*[@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only' and contains(., 'Selecionar')]`); //Aguarda o botão selecionar ficar habilitado
        await page.evaluate(() => document.querySelector('button[id*="botaoSelecionarProcessosOutrosUsuariosModal"] span').click());
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
        await new Promise(r => setTimeout(r, 500)); 
    } else {
        await page.evaluate(() => document.querySelector('div button[id*="botaoFecharProcessosOutrosUsuariosModal"] span').click());
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    }
}

    //####### SE ABRIR FORMULÁRIO COM PREVENTO QUANDO DA INSERÇÃO INDIVIDUAL - SE O PROCESSO JÁ ESTÁ COM OUTRO PROCURADOR
async function modal_prevento_individual(page) {
    await page.waitForSelector('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr td', {visible: true});
    await page.evaluate(() => {document.querySelector(`tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr[data-ri="0"] td div div span`).click()});//Clica no check box da linha
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForSelector('tbody[id*="tabelaModalProcessosJudiciaisOutrosUsuarios_data"] tr[data-ri="0"] td div div div span[class="ui-chkbox-icon ui-icon ui-icon-check"]');
    await new Promise(r => setTimeout(r, 200));
    let estado_botao = await 'Selecionado';
    if (estado_botao === 'Selecionado') {
        await page.waitForXPath(`//*[@class='ui-button ui-widget ui-state-default ui-corner-all ui-button-text-only' and contains(., 'Selecionar')]`); //Aguarda o botão selecionar ficar habilitado
        await page.evaluate(() => document.querySelector('button[id*="botaoSelecionarProcessosOutrosUsuariosModal"] span').click());
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
        await new Promise(r => setTimeout(r, 200)); 
    } else {
        await page.evaluate(() => document.querySelector('div button[id*="botaoFecharProcessosOutrosUsuariosModal"] span').click());
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    }
    //await page.waitForTimeout(500); // Aguarda 0,5 segundo antes de verificar novamente
}

let click_modal_escolha_classe = async (page, indice_classe) => {
    await new Promise(r => setTimeout(r, 700));
    await page.evaluate((indice_classe) => {document.querySelector(`tbody[id$="tabelaModalProcessosJudiciais_data"] tr[data-ri="${indice_classe}"] td div div div span`).click()}, indice_classe); // ###### Seleciona/click - checkbox - processo da classe listada no imput (planilha origem)
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForSelector(`tbody[id$="tabelaModalProcessosJudiciais_data"] tr[data-ri="${indice_classe}"] td div div div[class="ui-chkbox-box ui-widget ui-corner-all ui-state-default ui-state-active"] span[class="ui-chkbox-icon ui-icon ui-icon-check"]`); // Aguarda o checkbox ficar selecionado
    await page.evaluate(() => document.querySelector('div button[id="frmDistProc:mdListProc:botaoSelecionarProcessoModal"] span').click());// ###### clica no botão "Selecionar"
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmDistProc:mdListProc:dialogListaProcessos").style.visibility === "hidden"');
    await page.waitForFunction(`document.getElementById("frmDistProc:numProc").value === ""`);
}

async function modal_escolha_classe(page, classe_input) { 
    await page.waitForSelector('tbody[id="frmDistProc:mdListProc:tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]', {visible: true}, {timeout: 120000}) //Aguarda a tabela do modal ficar visivel
    let classes_modal = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="frmDistProc:mdListProc:tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]')).map((el)=>{return el.innerText}));//Retorna as classes do listadas no modal em um array
    let indice = await classes_modal.findIndex(element => element === classe_input); //compara as classes listadas no modal (array) com a classe da planilha imput e retorna o indice encontrado
    if (await indice !== -1) {    // ###### Se encontrar a classe pesquisada 
        click_modal_escolha_classe(page, indice);
    } else if (await indice == -1) {
        indice = await indices_pesquisa(instancia_1, classes_modal);
        classe_input = await instancia_1[indice];
        indice = await classes_modal.findIndex(element => element === classe_input);
        if (await indice !== -1) {
            click_modal_escolha_classe(page, indice);
        }  else {    // ###### Se não encontrar a classe pesquisada - clica no botão fechar
            await console.log('A classe do SAJ não coincide com a classe da planilha')
            await page.evaluate(() => document.querySelector('#formManifestacao\\:modalListaProcessos\\:botaoFecharProcessoModal > span').click()); 
        }
    }             
}

let pagina_distribuicao = async (page, unidade) => {
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/pages/tarefas/distribuirProcessoProcurador.jsf?');//AMBIENTE DE TESTE
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/tarefas/distribuirProcessoProcurador.jsf?', {waitUntil: 'networkidle0'}),//AMBIENTE DE PRODUÇÃO
    await page.waitForSelector('label[id$="procuradoria:selectOneMenu_label"]', {visible: true});
    let procuradoria = await page.$$eval('label[id*="procuradoria:selectOneMenu_label"]', anchors => { return anchors.map(anchor => anchor.textContent)})
    if (await procuradoria[0] !== 'PRFN5 (Sede)') {
        const indice = await indice_opcao(page, 'div[id*="procuradoria:selectOneMenu_panel"] > ul > li', 'PRFN5 (Sede)');
        await page.click('select[name$="procuradoria:selectOneMenu_input"]');
        await page.waitForSelector('div[id*="procuradoria:selectOneMenu_panel"] > ul > li', {visible: true});
        await page.click(`div[id*="procuradoria:selectOneMenu_panel"] > ul > li:nth-child(${indice})`);
        await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${unidade}')]`); //Aguarda o campo da Procuradoria ser alterado
        await new Promise(r => setTimeout(r, 1000));
    }
};

let incluir_lote = async (page, array, instancia) => {
    let lista = await array.map(el => el.numero).join(',');
    const lista_acabada = await lista.replace(/,+/g, "\n");
    await page.waitForSelector('div[id="frmDistProc:btnInc"] button[id="frmDistProc:btnInc_menuButton"] span');
    await page.click('div[id="frmDistProc:btnInc"] button[id="frmDistProc:btnInc_menuButton"] span', {visible: true});
    //await page.click('tbody tr td div div[id$="btnInc"] button[id$="btnInc_menuButton"] span', {visible: true});
    await page.waitForSelector('div ul li a[id$="menuIncLote"]', {visible: true});
    await page.click('div ul li a[id$="menuIncLote"]');
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmModais:modalProcessos:dialogListaEmBloco").style.visibility === "visible"');//Aguarda o modal de inserir processo em bloco ficar visível
    await page.evaluate((lista_acabada) => {document.querySelector('tbody tr td textarea[id$="listaNrProcessos"]').value = lista_acabada}, lista_acabada);//Inclui os processos na lista
    if (await instancia == '2ª Instância') {
        await page.click('select[id="frmModais:modalProcessos:instancias:selectOneMenu_input"]');
        await page.waitForSelector("#frmModais\\:modalProcessos\\:instancias\\:selectOneMenu_panel > ul > li");
        await page.click("#frmModais\\:modalProcessos\\:instancias\\:selectOneMenu_panel > ul > li:nth-child(2)");
        await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${instancia}')]`);
    } 
    await page.click('button[id*="modalProcessos:botaoConfirmaInclusao"]') ;
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await resposta_lote(page, array);
}

let resposta_lote = async (page, array) => {
    const modal_existe = await verifica_se_elemento_existe(page, 'div[id*="frmDistProc:mdListProcUsuariosDiferentes:dialogListaProcessosOutrosUsuarios"]');
    if (await modal_existe) {//Se o modal prevento estiver visível
        await modal_prevento_lote(page); //Confirma a distribuição ou não e fecha o formulário modal
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    }
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "visible"'); 
    //await page.waitForSelector('div[id="frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir"] table tfoot tr td[colspan="4"]', {visible: true}); //Aguarda o botão do modal - incluir em lote - ficar habilitado
    const titulo = await page.evaluate(() => document.querySelector('tbody tr td label[id$="erroMsg"]').innerText);//Variável recebe o conteúdo do resultado 
    //let todos_distribuidos = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(1)')).map((el)=>{return el.innerText})); // Array com todos os processos que estão na lista, inclusive os com classe duplicada
    //let todos_classe = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(2)')).map((el)=>{return el.innerText})); // Array com todos os processos que estão na lista, inclusive os com classe duplicada
    if (await titulo !== 'Todos os processos foram incluídos com sucesso.') {
        let array_proc_classe_dupla = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(1)')).map((el)=>{return el.innerText}));
        let array_entrada = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(4)')).map((el)=>{return el.innerText}));
        await array_entrada.map((texto, indice) => { //Fazer o map - ver quais processos tem + de uma classe
            if(texto === 'Foram encontrados mais de um processo com este número.'|| texto === 'Nenhum processo encontrado na 1ª Instância com esse número.'){
                dist_dif.push(array_proc_classe_dupla[indice]);
            }  
        }); 
    } 
    /*if (await dist_prevento.length > 0) {// Testa as classes da resposta para os preventos - se entraram na lista direto
        await dist_prevento.map(function(item) {
            const p = item.replace(/[\.,\-]/g, "");
            let indice = todos_distribuidos.findIndex(element => element === p);
            if (indice !== -1) {
                console.log('Indice: '+indice);
                console.log(p+' - '+todos_classe[indice]);
                let r = array.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
                    return (obj.numero === item);
                });
                console.log(r);
            }
        });
    }*/
    await page.click('div[class="containerBotao"] button[id*="btnFecharConc"] span');
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "hidden"');//Aguarda o modal do resultado das classes ficar invisível
}

let distribuir = async (page, listdistribuicao_triador) => {
    let quant_proc_lista = await page.$eval('tfoot tr td[colspan="8"]', elem => elem.innerText.substring(20,elem.innerText.length));
    await console.log(`${quant_proc_lista} processos incluídos direto` );
    if (await dist_dif.length > 0) {//Para os processos com + de uma classe cadastrada
        await console.log(`${dist_dif.length} processos para fazer seleção`);
        for (let nd = 0; nd < await dist_dif.length; nd++) {
            let numero_pj = await dist_dif[nd];
            let processos_formatado = await numero_pj.replace(/(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})/g,"\$1-\$2.\$3.\$4.\$5.\$6");
            let id = await listdistribuicao_triador.dados.findIndex( dados => dados.numero === processos_formatado );
            await console.log(listdistribuicao_triador.dados[id].numero + ' - ' + listdistribuicao_triador.dados[id].classe);
            await page.$eval('input[id="frmDistProc:numProc"]', (el, value) => el.value = value, numero_pj); //inclui o número do pj
            await page.waitForFunction(`document.getElementById("frmDistProc:numProc").value === "${numero_pj}"`);
            await page.keyboard.press('Enter');
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            await page.waitForFunction('document.getElementById("frmDistProc:mdListProc:dialogListaProcessos").style.visibility === "visible"'); //Aguarda o modal da escolha de classe ficar visível
            await modal_escolha_classe(page, listdistribuicao_triador.dados[id].classe); // Função que escolhe a classe
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"'); 
            await page.waitForFunction('document.getElementById("frmDistProc:mdListProc:dialogListaProcessos").style.visibility === "hidden"');//Aguarda o modal da escolha de classe ficar invisível
            await new Promise(r => setTimeout(r, 1000));
            let modal_prevento = await verifica_se_elemento_existe(page, 'div[id="frmDistProc:mdListProcUsuariosDiferentes:dialogListaProcessosOutrosUsuarios"]');//Verifica se o modal prevento está visível - se o processo já está com outro servidor/procurador
            if (await modal_prevento) {//Se o modal prevento estiver visível
                await modal_prevento_individual(page); //Confirma a distribuição ou não e fecha o formulário modal
                await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            }
            await quant_proc_lista++;
            await page.focus('input[id="frmDistProc:numProc"]');
        }
    }  
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await rolar_tela(page);
    await page.waitForSelector('tfoot tr td[colspan="8"]', {visible: true});
    let total_pj = await page.$eval('tfoot tr td[colspan="8"]', elem => elem.innerText.substring(20,elem.innerText.length));
    await page.waitForSelector('button[class="ui-datepicker-trigger ui-button ui-widget ui-state-default ui-corner-all ui-button-icon-only"] span[class="ui-button-icon-left ui-icon ui-icon-calendar"]', {visible: true}); //Aguarda o campo data de entrada ficar habilitado
    await inclui_dados_procurador_distribuicao(page, listdistribuicao_triador.triador);
    await page.click('span[id="frmDistProc:dtEnt:calendar"] button[class="ui-datepicker-trigger ui-button ui-widget ui-state-default ui-corner-all ui-button-icon-only"] span[class="ui-button-icon-left ui-icon ui-icon-calendar"]');
    await page.waitForSelector('table[class="ui-datepicker-calendar"] tbody tr td[class*="ui-datepicker-today"] a', {visible: true});
    await page.click('table[class="ui-datepicker-calendar"] tbody tr td[class*="ui-datepicker-today"] a');
    await page.waitForFunction(`document.querySelector("input[id='frmDistProc:dtEnt:calendar_input']").value.includes("${data_distribuicao}")`);
    let final_linha = await (total_pj - 1);
    await aguarda_data(page, final_linha);
    total_pj == listdistribuicao_triador.dados.length ? await console.log(`Todos os ${listdistribuicao_triador.dados.length} processos foram distribuídos`) :  await console.log('Existe processo(s) que não foi distribuído(s)')
    await click_botao(page, 'Concluir');
    await page.waitForXPath("//*[@class='ui-messages-info-summary' and contains(., 'Execução de distribuição de processos concluída.')]");
    await page.waitForXPath(`//*[@class='texto-bold' and contains(., '${listdistribuicao_triador.triador}')]`);
    await console.log("\n");
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/pages/principal.jsf?');//AMBIENTE DE TESTE
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/principal.jsf?');//AMBIENTE DE PRODUÇÃO
    await new Promise(r => setTimeout(r, 1000));
};

ler_excel();

let scrape = async () => {

    const browser = await puppeteer.launch({ //cria uma instância do navegador
        executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        headless: false, args: ['--start-maximized'],//torna visível 
        ignoreHTTPSErrors: true,
    });
    
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close(); 
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/login.jsf', {timeout: 120000});//AMBIENTE DE TESTE
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf', {waitUntil: 'networkidle0'}); //AMBIENTE DE PRODUÇÃO
    await dados_inicio(page);
    await page.waitForNavigation();
    let id = await ler_indice_distribuicao(nome_arquivo_excel);
    await console.log('_____________________________________________________ ')
    for (let m = id; m < await listdistribuicao.length; m++) {
        dist_dif = await [];
        await console.log(`${listdistribuicao[m].triador} - ${listdistribuicao[m].dados.length} processos para distribuição`)
        await pagina_distribuicao(page, listdistribuicao[m].unidade);
        let array_2inst = await listdistribuicao[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
            return (!juizo_1inst.includes(obj.instancia));
        });
        let array_1inst = await listdistribuicao[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 1ª instância
            return (juizo_1inst.includes(obj.instancia));
        });
        if (array_1inst.length > 0) { await incluir_lote(page, array_1inst, '1ª Instância') }//Função para inserir os processos de 1ª instância
        if (array_2inst.length > 0) { await incluir_lote(page, array_2inst, '2ª Instância') }//Função para inserir os processos de 2ª instância 
        await distribuir(page, listdistribuicao[m]);
        await salva_indice_distribuicao(nome_arquivo_excel, (m+1));
        array_2inst = await [];
        array_1inst = await[];  
    }
    await apagar_indices_distribuicao(nome_arquivo_excel);
    await new Promise(r => setTimeout(r, 3000));
    await browser.close()
    let result = `Total: ` + listdistribuicao.length;
    return result
}  

scrape().then((value) => {
   console.log(value)
});