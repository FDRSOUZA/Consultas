"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
const fs = require('fs');
var path = require('path');
const { Console } = require('console');
var instancia_1;
let dist_dif = [];
let listabaixa = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `Índices Baixa - ${dia}-${mes}-${ano}.xlsx`;

const juizo_1inst = ['', '1ª Instância']

const classes_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (informações do débito pendente)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)',
    'Execução Fiscal (FNDE)', 'Cumprimento de Sentença', 'Ação Trabalhista', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cumprimento de Sentença contra a Fazenda Pública', 'Arrolamento', 
    'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Carta Precatória', 'Cautelar', 'Cautelar Fiscal', 'Consignação em Pagamento', 'Desapropriação',  
    'Embargos de Terceiro', 'Execução de Título Extrajudicial', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Exibição de Documento ou Coisa', 'Falência', 'Habilitação',   
    'Incidente de Desconsideração de Personalidade Jurídica', 'Inventário', 'Outras', 'Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Protesto',  'Petição', 'Reclamação', 'Recuperação Judicial', 
    'Representação', 'Restauração de Autos', 'Restituição de Coisa ou Dinheiro na Falência', 'Reintegração / Manutenção de Posse', 'Tutela Antecipada Antecedente',  'Tutela de Cautelar Antecedente', 'Usucapião'];

const classes_defesa = ['Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Mandado de Segurança', 'Cumprimento de Sentença', 'Mandado de Segurança Coletivo', 'Cumprimento de Sentença contra a Fazenda Pública', 
    'Cumprimento Provisório de Sentença', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Execução de Título Extrajudicial', 'Execução de Título Extrajudicial contra a Fazenda Pública', 'Petição', 'Outras',
    'Embargos de Terceiro', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cautelar', 'Cautelar Fiscal', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Restauração de Autos', 'Monitória',
    'Notificação', 'Habilitação', 'Habeas Data', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Ação Penal', 'Ação Popular', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Habeas Corpus', 
    'Liquidação por Arbitramento', 'Liquidação Provisória por Arbitramento', 'Liquidação Provisória pelo Procedimento Comum', 'Liquidação pelo Procedimento Comum', 'Liquidação Provisória de Sentença', 
    'Reintegração / Manutenção de Posse', 'Retificação de Registro de Imóvel', 'Restituição de Coisas Apreendidas', 'Exibição de Documento ou Coisa', 'Produção Antecipada da Prova', 'Procedimento Sumário', 
    'Protesto', 'Representação', 'Oposição', 'Depósito', 'Restituição de Coisa ou Dinheiro na Falência', 'Ação de Demarcação', 'Ação de Exigir Contas', 'Consignação em Pagamento', 'Depósito da Lei 8.866/94',  
    'Desapropriação', 'Despejo', 'Dissolução e Liquidação de Sociedade', 'Mandado de Injunção', 'Recuperação Judicial', 'Embargos à Adjudicação', 'Embargos à Arrematação', 'Impugnação de Assistência Judiciária', 
    'Incidente de Suspeição', 'Impugnação de Crédito', 'Impugnação ao Pedido de Assist Litiscon Simples', 'Impugnação Ao Cumprimento de Sentença', 'Incidente de Impedimento', 'Incidente de Falsidade', 
    'Incidente de Desconsideração de Personalidade Jurídica', 'Pedido de Quebra Sigilo de Dados e/ou Telefônico', 'Insolvência Requerida pelo Devedor ou pelo Espólio', 'Dissolução e Liquidação de Sociedade', 
    'Dissolução Parcial de Sociedade', 'Exceção de Incompetência', 'Exceção de Litispendência', 'Falência', 'Embargos Infringentes na Execução Fiscal', 'Impugnação ao Valor da Causa', 'Carta Precatória', 
    'Carta Rogatória',  'Carta de Sentença', 'Carta de Ordem', 'Arrolamento', 'Inventário', 'Usucapião', 'Ação Trabalhista', 'Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FNDE)', 
    'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];

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
    await page.waitForSelector(`tbody[id="frmBaixa:dtTb_data"] tr:nth-child(${linha_anterior}) td:nth-child(6)`);
    let objeto_baixa = await page.$eval(`tbody[id="frmBaixa:dtTb_data"] tr:nth-child(${linha_anterior}) td:nth-child(6) div span`, el => el.textContent); 
    let tipo_objeto = await objeto_baixa.trim();
    if (await tipo_objeto === 'Autos') {
        let img_coluna; 
        for (var i = 1; i <= 3500; i++) {
            img_coluna = await indicadores_coluna_5(page, linha_anterior);
            if (img_coluna > 0) {
                i= 4000}
        }
        if (img_coluna > 0) {
            await page.waitForSelector (`tbody[id="frmBaixa:dtTb_data"] tr:nth-child(${linha_anterior}) td:nth-child(5) div[class="ui-dt-c"] div div[style="display:inline-block"]`, {visible: true});//Aguarda a coluna 5 formar (tornar visível todos os elementos)
        }
    } else {
        await page.waitForSelector (`tbody[id="frmBaixa:dtTb_data"] tr:nth-child(${linha_anterior}) td:nth-child(5) div[class="ui-dt-c"] div div[style="display:inline-block"]`, {visible: true});//Aguarda a coluna 5 formar (tornar visível todos os elementos)
    } 
}


let indicadores_coluna_5 = async (page, linha_anterior) => {
    let painel_caixas_disponiveis = await page.evaluate((linha_anterior)=>{
        let allcaixas_span = Array.from(document.querySelectorAll(`tbody[id="frmBaixa:dtTb_data"] tr:nth-child(${linha_anterior}) td:nth-child(5) div[class="ui-dt-c"] div[id*="pgIndProc"] div`))
        let allcaixas_a_array = allcaixas_span.map(i=>{return i.id}) // retorna o id de todas as jurisdições que constam no painel 
        return allcaixas_a_array
    }, linha_anterior) ;
    return painel_caixas_disponiveis.length 
}


let salva_indice_baixa = async (arq, id_triador) => {
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Índice_Triador");
    worksheet.columns = [
        {header: 'Triador', key: 'triador', width: 25}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    await worksheet.addRow({triador: id_triador});
    await wbook.xlsx.writeFile(arq);
};

let apagar_indices_baixa = async (arq) => {
    var filePath = await path.resolve(__dirname, arq);
    if (fs.existsSync(filePath)) { 
        try {
            fs.unlinkSync(filePath)
          } catch(err) {console.error(err)}
    }
};

let ler_indice_baixa = async (arq) => {
    let id_baixa;
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, arq);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Índice_Triador'); // Primeira aba do arquivo excel - Planilha
        id_baixa = await parseInt(sh.getRow(2).getCell(1).text);    
    }   else {
        id_baixa = await 0;
    }
    return id_baixa;
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
        //await tabela.push({"triador":row.values[1], "plan": row.values[2], "grupo": row.values[3], "unidade": row.values[4], "dados": ''}) //Salva cada linha em um array
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
    const fixo = ['Processo', 'Classe judicial', 'Juízo', 'Setor responsável', 'Procurador', 'Tipo de petição', 'Providência Apoio'];
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
        let planilha_lida = new Array();

        await sh.eachRow(function(cell, rowNumber) {
            let triador;
            (sh.getRow(rowNumber).getCell(coluna[6]).text === '' || sh.getRow(rowNumber).getCell(coluna[6]).text === 'Já distribuído no SAJ') ? triador = sh.getRow(rowNumber).getCell(coluna[4]).text : triador = sh.getRow(rowNumber).getCell(coluna[6]).text;
            if (sh.getRow(rowNumber).getCell([coluna[3]]).text !== '' && sh.getRow(rowNumber).getCell([coluna[3]]).text.substring(0,7) === 'TRIAGEM' && (sh.getRow(rowNumber).getCell(coluna[5]).text !== '') && triadores_grupo.some(t => t.triador.includes(triador))) {    
                //planilha_lida.push({"numero": sh.getRow(rowNumber).getCell([coluna[0]]).text, "classe": sh.getRow(rowNumber).getCell([coluna[1]]).text, "instancia": sh.getRow(rowNumber).getCell([coluna[2]]).text, "triador": triador})
                let processo = {"numero": sh.getRow(rowNumber).getCell([coluna[0]]).text, "classe": sh.getRow(rowNumber).getCell([coluna[1]]).text, "instancia": sh.getRow(rowNumber).getCell([coluna[2]]).text};
                let id = triadores_grupo.findIndex( arr => arr.triador === triador);
                triadores_grupo[id].dados.push(processo)
            }
        });

        listabaixa = await triadores_grupo.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
            return (obj.dados.length > 0);
        });              
        
        console.log("\n");
        await listabaixa.map(function(obj) {
            console.log(`Triador: ${obj.triador} - ${(obj.dados.length)} processos`) 
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

        //################# FUNÇÃO PARA INSERIR VALOR E DISPARAR O EVENTO NO CAMPO SELECT 
const select_evaluete = async (page, seletor_acao, seletor_evento) => {
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

let click_modal_escolha_classe = async (page, indice_classe) => {
    await new Promise(r => setTimeout(r, 700));
    await page.evaluate((indice_classe) => {document.querySelector(`tbody[id*="tabelaModalProcessosJudiciais_data"] tr[data-ri="${indice_classe}"] td div div div span`).click()}, indice_classe); // ###### Seleciona/click - checkbox - processo da classe listada no imput (planilha origem)
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await new Promise(r => setTimeout(r, 700));
    //await page.waitForFunction('document.getElementById("frmBaixa:modalListaProcessos:tabelaModalProcessosJudiciais_data")).ariaSelected = "true"');
    //await page.waitForSelector(`tbody[id*="tabelaModalProcessosJudiciais_data"] tr[data-ri="${indice_classe}"] td div div div[class="ui-chkbox-box ui-widget ui-corner-all ui-state-default ui-state-active"] span[class="ui-chkbox-icon ui-icon ui-icon-check"]`); // Aguarda o checkbox ficar selecionado
    await page.evaluate(() => document.querySelector('div button[id*="botaoSelecionarProcessoModal"] span').click());// ###### clica no botão "Selecionar"
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
}

async function modal_escolha_classe(page, classe_input, num_linha) { 
    await page.waitForFunction('document.getElementById("frmBaixa:modalListaProcessos:dialogListaProcessos").style.visibility === "visible"');
    let classes_modal = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="tabelaModalProcessosJudiciais_data"] tr td:nth-child(3) div[class*="ui-dt-c"]')).map((el)=>{return el.innerText})); //Lista em um array todas as classes do form modal visível
    let indice = await classes_modal.findIndex(element => element === classe_input); //compara as classes listadas no modal (array) com a classe da planilha imput e retorna o indice encontrado
    if (await indice !== -1) {    // ###### Se encontrar a classe pesquisada 
        click_modal_escolha_classe(page, indice);
    } else if (await indice == -1) {
        indice = await indices_pesquisa(instancia_1, classes_modal);
        classe_input = await instancia_1[indice];
        indice = await classes_modal.findIndex(element => element === classe_input);
        if (await indice !== -1) {
            click_modal_escolha_classe(page, indice);
        }
    }
    //await nova_linha(page, num_linha);         
}


let pagina_baixa = async (page) => {
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/pages/mesaTrabalho/mesaTrabalho.jsf?');
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/mesaTrabalho/mesaTrabalho.jsf?');
    await page.waitForSelector('button[id$="mbExecutarTarefa_button"]'); 
    await click_botao(page, 'Executar Tarefa');
    await page.waitForSelector('div[id*="mbExecutarTarefa_menu"] ul li a[id*="menuBaixarProcessos"]', {visible: true});
    await page.click('div[id*="mbExecutarTarefa_menu"] ul li a[id*="menuBaixarProcessos"] span')
    await new Promise(r => setTimeout(r, 1000));
};

let dividir_lote = async (page, array, instancia) => {
    let corte_final;
    let r = 0;
    let corte_inicial = 0;
    let repeticao = await (Math.ceil(array.length/200));
    await repeticao === 1 ? corte_final = array.length : corte_final = 200;
    do {
        let lote_processo = await array.slice(corte_inicial, corte_final);
        await incluir_lote(page, lote_processo, instancia);
        corte_inicial = await corte_final;
        await r++;
        await r == (repeticao-1) ? corte_final = await array.length : corte_final = await (corte_final+200);
    } while (r < repeticao)
}

let incluir_lote = async (page, array, instancia) => {
    let lista = await array.map(el => el.numero).join(',');
    const lista_acabada = await lista.replace(/,+/g, "\n");
    //await console.log(lista_acabada);
    await page.waitForSelector('div[id="frmBaixa:incluir"] button[id="frmBaixa:incluir_menuButton"] span', {visible: true});
    await page.click('div[id="frmBaixa:incluir"] button[id="frmBaixa:incluir_menuButton"] span');
    await page.waitForSelector('div ul li a[id$="menuIncLote"]', {visible: true});
    await page.click('div ul li a[id$="menuIncLote"]');
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await page.waitForFunction('document.getElementById("frmModais:modalProcessos:dialogListaEmBloco").style.visibility === "visible"');//Aguarda o modal de inserir processo em bloco ficar visível
    await page.evaluate((lista_acabada) => {document.querySelector('tbody tr td textarea[id$="listaNrProcessos"]').value = lista_acabada}, lista_acabada);//Inclui os processos na lista
    if (await instancia == '2ª Instância') {
        await select_evaluete(page, "#frmModais\\:modalProcessos\\:instancias\\:selectOneMenu_panel > ul > li:nth-child(2)", 'select[id="frmModais:modalProcessos:instancias:selectOneMenu_input"]') //Clica no nome do Procurador signatário
        await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., '${instancia}')]`);
    } 
    await page.click('button[id*="modalProcessos:botaoConfirmaInclusao"]') ;
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
    await resposta_lote(page);
}

let resposta_lote = async (page) => {
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "visible"');
    await page.waitForSelector('label[id="frmModais:modalListaEmBlocoConcluir:erroMsg"]', {visible: true}); //Aguarda o botão do modal - incluir em lote - ficar habilitado
    const titulo = await page.evaluate(() => document.querySelector('tbody tr td label[id$="erroMsg"]').innerText);//Variável recebe o conteúdo do resultado
    if (await titulo !== 'Todos os processos foram incluídos com sucesso.') {
        let array_proc_classe_dupla = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(1)')).map((el)=>{return el.innerText}));
        let array_entrada = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id$="gridConcluirInclusaoBloco_data"] tr td:nth-child(4)')).map((el)=>{return el.innerText}));
        await array_entrada.map((texto, indice) => { //Fazer o map - ver quais processos tem + de uma classe
            if(texto === 'Foram encontrados mais de um processo com este número.'|| texto === 'Nenhum processo encontrado na 1ª Instância com esse número.'){
                dist_dif.push(array_proc_classe_dupla[indice]);
            } 
        }) 
    } 
    await page.click('div[class="containerBotao"] button[id*="btnFecharConc"] span');
    await page.waitForFunction('document.getElementById("frmModais:modalListaEmBlocoConcluir:dialogListaEmBlocoConcluir").style.visibility === "hidden"');//Aguarda o modal do resultado das classes ficar invisível
    await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
}

let conferir_baixa = async (page, listabaixa) => {
    let quant_proc_lista = await page.$eval('tfoot tr td[colspan="7"]', elem => elem.innerText.substring(20,elem.innerText.length));
    await console.log(`${quant_proc_lista} processos incluídos direto` );
    if (await dist_dif.length > 0) {//Para os processos com + de uma classe cadastrada
        await console.log('_____________________________');
        for (let nd = 0; nd < await dist_dif.length; nd++) {
            let numero_pj = await dist_dif[nd];
            let processos_formatado = await numero_pj.replace(/(\d{7})(\d{2})(\d{4})(\d{1})(\d{2})(\d{4})/g,"\$1-\$2.\$3.\$4.\$5.\$6");
            let id = await listabaixa.dados.findIndex( dados => dados.numero === processos_formatado );
            await console.log(listabaixa.dados[id].numero + ' - ' + listabaixa.dados[id].classe);
            await page.$eval('input[name$="frmBaixa:numeroProcesso"]', (el, value) => el.value = value, numero_pj); //inclui o número do pj 
            await page.waitForFunction(`document.getElementById("frmBaixa:numeroProcesso").value === "${numero_pj}"`);
            await page.click('button[id="frmBaixa:incluir_button"] span[class="ui-button-text"]');
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            //await page.waitForSelector('div[id*="frmBaixa:modalListaProcessos:dialogListaProcessos"]', {visible: true});
            await modal_escolha_classe(page, listabaixa.dados[id].classe, (quant_proc_lista));
            await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"');
            await quant_proc_lista++;
            await page.waitForXPath(`//*[@class='ui-datatable-footer ui-widget-header' and contains(., 'Total de processos: ${quant_proc_lista}')]`, {timeout: 120000});
            await page.focus('input[name$="frmBaixa:numeroProcesso"]'); //caixa de texto do número do pj recebe o foco
            await page.waitForFunction(() => document.querySelector('input[name$="frmBaixa:numeroProcesso"]').focus);
        }
    }
    quant_proc_lista = await page.$eval('tfoot tr td[colspan="7"]', elem => elem.innerText.substring(20,elem.innerText.length));
    //await nova_linha(page, (quant_proc_lista-1));   
    await new Promise(r => setTimeout(r, 1500));
    let texto_alerta = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id="frmBaixa:dtTb_data"] tr td:nth-child(6) div[class="ui-dt-c"] div span[class="texto-info-red texto-bold"]')).length);
    await rolar_tela(page);
    await new Promise(r => setTimeout(r, 1000));
    await click_botao(page, 'Concluir');
    await new Promise(r => setTimeout(r, 1000));
    let form_cump_decisao;
    let botao_confirma_baixa;
    let selecao_tipo;
    if (texto_alerta > 0) {
        await page.waitForSelector('div[id*="confirmaDlg"]', {visible: true});
        let id_cx_dialogo = await page.evaluate(() => Array.from(document.querySelectorAll('div[id*="confirmaDlg"]')).map((el)=>{return el.id}));
        await page.waitForFunction(`document.getElementById("${id_cx_dialogo[0]}").style.visibility === "visible"`);
        let id_botao = await page.evaluate(() => Array.from(document.querySelectorAll('div[id*="confirmaDlg"] div div button')).map((el)=>{return el.id}));
        let botao_confirm = await page.$(`button[id="${id_botao[0]}"]`);
        await botao_confirm.click();
        await new Promise(r => setTimeout(r, 2000));
        form_cump_decisao = await verifica_se_elemento_existe(page, 'div[id="frmBaixa:mdlMsgAcum"]');
        if (await form_cump_decisao) {
            selecao_tipo = await page.$('table[id="frmBaixa:opcMsgAcum"] tbody tr td div div input[id="frmBaixa:opcMsgAcum:1"]'); //Seletor de autuação
            await selecao_tipo.click(); //Clica no tipo de manter processos na seleção e concluir
            //botao_confirma_baixa = await page.$('div[id="frmBaixa:mdlMsgAcum"] div div button[id="frmBaixa:j_idt183"] span');
            await new Promise(r => setTimeout(r, 1000)); 
            const botao = await page.$x(`//span[contains(text(), 'Confirmar')]`); //Procura o botão pelo texto, dependendo do argumento 
            botao[0].click(); //Clicar no botão limpar ou voltar, dependendo do argumento 
        }
    } else {
        await new Promise(r => setTimeout(r, 2000));
        form_cump_decisao = await verifica_se_elemento_existe(page, 'div[id="frmBaixa:mdlMsgAcum"]');
        if (await form_cump_decisao) {
            selecao_tipo = await page.$('table[id="frmBaixa:opcMsgAcum"] tbody tr td div div input[id="frmBaixa:opcMsgAcum:1"]'); //Seletor de autuação
            await selecao_tipo.click(); //Clica no tipo de manter processos na seleção e concluir
            let lista_id_botton = await page.evaluate(() => Array.from(document.querySelectorAll('div[id="frmBaixa:mdlMsgAcum"] div div button')).map((el)=>{return el.id})); //identifica id do botão concluir
            botao_confirma_baixa = await page.$(`div[id="frmBaixa:mdlMsgAcum"] div div button[id="${lista_id_botton[0]}"] span`);
            await botao_confirma_baixa.click();
        } 
    } 
    await page.waitForSelector(`tbody[id="frmBaixa:dtResultadoBaixa_data"] tr[data-ri="${(quant_proc_lista-1)}"] td div span`);//Aguarda preencher a ultima coluna (status) da ultima linha de tabela baixa
    let status = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="frmBaixa:dtResultadoBaixa_data"] tr td:nth-child(5)')).map((el)=>{return el.innerText}));//Coloca o texto da coluna status em um array para verificar se existe processo com erro
    let id_erro = await status.findIndex(element => element === 'Erro de processamento no SAJ');//Pesquisa no array para verificar se encontra alguma linha de processo com erro
    if (await id_erro !== -1) {
        await console.log('Processo com erro');
    }
    await page.waitForXPath("//*[@class='ui-messages-info-summary' and contains(., 'Baixa executada com sucesso.')]", {timeout: 1200000});
    await new Promise(r => setTimeout(r, 500));
    await console.log("\n");
    //await page.goto('https://hom.saj.pgfn.fazenda.gov.br/saj/pages/principal.jsf?');//AMBIENTE DE TESTE
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/principal.jsf?');//AMBIENTE DE PRODUÇÃO
    await page.waitForFunction(() => document.querySelector('Strong').innerText == 'Sejam bem-vindos ao Sistema de Acompanhamento Judicial - SAJ.');
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
    //await page.waitForSelector('input[name$="formMenus"]', {timeout: 120000}, {visible: true}); //Aguarda a página carregar o menu "Processo"
    await new Promise(r => setTimeout(r, 500));
    let id = await ler_indice_baixa(nome_arquivo_excel);
    await console.log('_____________________________________________________ ')
    for (let m = id; m < await listabaixa.length; m++) {
        dist_dif = await [];
        await console.log(`${listabaixa[m].triador} - ${(listabaixa[m].dados.length)} processos para baixa`);
        let array_2inst = await listabaixa[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
            return (!juizo_1inst.includes(obj.instancia));
        });
        let array_1inst = await listabaixa[m].dados.filter(function(obj) {//Seleciona em um array os dados do processos de 1ª instância
            return (juizo_1inst.includes(obj.instancia));
        });
        await pagina_baixa(page);
        if (array_1inst.length > 0) { await dividir_lote(page, array_1inst, '1ª Instância') }//Função para inserir os processos de 1ª instância
        if (array_2inst.length > 0) { await dividir_lote(page, array_2inst, '2ª Instância') }//Função para inserir os processos de 2ª instância 
        await conferir_baixa(page, listabaixa[m]);
        await salva_indice_baixa(nome_arquivo_excel, (m+1));
        array_2inst = await [];
        array_1inst = await[];  
    }
    await apagar_indices_baixa(nome_arquivo_excel);
    await new Promise(r => setTimeout(r, 3000));
    await browser.close()
    let result = `Total: ` + listabaixa.length;
    return result
}  

scrape().then((value) => {
   console.log(value)
});