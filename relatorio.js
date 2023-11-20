"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var path = require('path');
const fs = require('fs');
let listprocessos = [];
let listconsulta = [];

let classes_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (informações do débito pendente)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (FNDE)', 
    'Cumprimento de Sentença', 'Ação Trabalhista', 'Cumprimento de Sentença contra a Fazenda Pública', 'Arrolamento', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Usucapião', 'Embargos de Terceiro',
    'Carta Precatória', 'Cautelar', 'Cautelar Fiscal', 'Consignação em Pagamento', 'Desapropriação', 'Execução de Título Extrajudicial', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Procedimento Comum',
    'Falência', 'Habilitação', 'Inventário', 'Outras', 'Protesto', 'Petição', 'Reclamação', 'Recuperação Judicial', 'Representação', 'Restauração de Autos', 'Embargos à Execução de Título Extrajudicial',
];

let classes_defesa = ['Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Mandado de Segurança', 'Cumprimento de Sentença', 'Mandado de Segurança Coletivo', 'Cumprimento de Sentença contra a Fazenda Pública', 
    'Cumprimento Provisório de Sentença', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Execução de Título Extrajudicial', 'Execução de Título Extrajudicial contra a Fazenda Pública', 'Petição', 'Outras',  
    'Embargos de Terceiro', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Cautelar', 'Cautelar Fiscal', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Restauração de Autos', 
    'Notificação', 'Habilitação', 'Procedimento Sumário', 'Impugnação ao Valor da Causa', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Ação Penal', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 
    'Liquidação por Arbitramento', 'Liquidação Provisória por Arbitramento', 'Liquidação Provisória pelo Procedimento Comum', 'Liquidação pelo Procedimento Comum', 'Liquidação Provisória de Sentença', 'Monitória',
    'Reintegração / Manutenção de Posse', 'Retificação de Registro de Imóvel', 'Restituição de Coisas Apreendidas', 'Exibição de Documento ou Coisa', 'Produção Antecipada da Prova', 'Procedimento Sumário', 'Protesto', 'Representação', 'Oposição', 'Depósito',
    'Restituição de Coisa ou Dinheiro na Falência', 'Ação de Demarcação', 'Ação de Exigir Contas', 'Consignação em Pagamento', 'Depósito da Lei 8.866/94', 'Desapropriação', 'Despejo', 'Dissolução e Liquidação de Sociedade', 
    'Mandado de Injunção', 'Recuperação Judicial', 'Embargos à Adjudicação', 'Embargos à Arrematação', 'Impugnação de Assistência Judiciária', 'Incidente de Suspeição', 'Impugnação de Crédito', 'Falência', 'Ação Popular',
    'Impugnação ao Pedido de Assist Litiscon Simples', 'Impugnação Ao Cumprimento de Sentença', 'Incidente de Impedimento', 'Incidente de Falsidade', 'Incidente de Desconsideração de Personalidade Jurídica', 'Pedido de Quebra Sigilo de Dados e/ou Telefônico',
    'Insolvência Requerida pelo Devedor ou pelo Espólio', 'Dissolução e Liquidação de Sociedade', 'Dissolução Parcial de Sociedade', 'Exceção de Incompetência', 'Embargos à Execução de Título Extrajudicial', 'Habeas Corpus',
    'Exceção de Litispendência', 'Embargos Infringentes na Execução Fiscal', 'Carta Precatória', 'Carta Rogatória', 'Oposição', 'Carta de Sentença', 'Carta de Ordem', 'Habeas Data'
] 

let tipo_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal (FNDE)', 'Execução Fiscal Previdenciária',
        'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)', 'EXECUÇÃO FISCAL'];

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

let ler_input = async () => {
    var wb = await new Excel.Workbook();
    let tabela_triadores = await ler_tabela_triadores();
    let triadores_grupo = await tabela_triadores.filter(function(obj) {
        return (obj.grupo == 'X' || obj.grupo == 'x');   
    });
    const plan = await triadores_grupo[0].plan;
    let fixo = ['Processo', 'Classe judicial', 'Juízo', 'Setor responsável', 'Procurador', 'Tipo de petição', 'Descrição', 'Providência Apoio'];
    var filePath = await path.resolve(__dirname,'Distribuição.xlsx');
    if (fs.existsSync(filePath)) {
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet(plan);
        let linha_ind = await sh.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
        let coluna = new Array(); 
        await fixo.forEach(function(elemento) {//Define os indices que serão manipulados - quais colunas
            let id_relacionado = linha_ind.findIndex(element => element === elemento);
            if(id_relacionado !== -1) {coluna.push(id_relacionado)}
        });
        await sh.eachRow(function(cell, rowNumber) {
            if (sh.getRow(rowNumber).getCell([coluna[3]]).text !== '' && sh.getRow(rowNumber).getCell([coluna[3]]).text.substring(0,7) === 'TRIAGEM' && triadores_grupo.some(t => t.triador.includes(sh.getRow(rowNumber).getCell([coluna[4]]).text))) {
                let apoio;
                sh.getRow(rowNumber).getCell(coluna[7]).text === '' || sh.getRow(rowNumber).getCell(coluna[5]).text === 'Já distribuído no SAJ' ? apoio = sh.getRow(rowNumber).getCell(coluna[4]).text : apoio = sh.getRow(rowNumber).getCell(coluna[7]).text 
                listprocessos.push({"numero": sh.getRow(rowNumber).getCell([coluna[0]]).text, "classe": sh.getRow(rowNumber).getCell([coluna[1]]).text, "triador": sh.getRow(rowNumber).getCell([coluna[4]]).text, "tipo_peticao": sh.getRow(rowNumber).getCell([coluna[5]]).text, "descricao": sh.getRow(rowNumber).getCell([coluna[6]]).text, "procurador": apoio})
            }
        });
    }
};

let ler_output = async (lista, arquivo) => {
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, arquivo);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Relatório'); // Primeira aba do arquivo excel - Planilha
        let i = 2
        do {
            if (sh.getRow(i).getCell(1).text !== '') {
                await lista.push({numero: sh.getRow(i).getCell(1).value, classe: sh.getRow(i).getCell(2).value, triador: sh.getRow(i).getCell(3).value, tipo_peticao: sh.getRow(i).getCell(4).value, descricao: sh.getRow(i).getCell(5).value, ultimo: sh.getRow(i).getCell(6).value, data_ultimo: sh.getRow(i).getCell(7).value, penultimo: sh.getRow(i).getCell(8).value, data_penultimo: sh.getRow(i).getCell(9).value, procurador: sh.getRow(i).getCell(10).value});
            }
            i++;
        } while (sh.getRow(i).getCell(1).text !== '')
    }
};

let compara_planilhas = async () => {
    let ultima_linha;
    await listconsulta.length > 0 ? ultima_linha = await listconsulta.length : ultima_linha = await 0;
    return ultima_linha;
};

let pesquisa_processo = async (page, numero_pj, classe) => {
    let acoes;
    let id_classe;
    (classe == 'EXECUÇÃO FISCAL' || tipo_execucao.includes(classe)) ? acoes = classes_execucao.concat(classes_defesa) : acoes = classes_defesa.concat(classes_execucao)
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/pesquisarProcessoJudicial.jsf?'); 
    await page.waitForSelector('input[name$="numProcesso"]'); //Aguarda carregar a página da "Consulta"
    await page.focus('input[name$="numProcesso"]'); //caixa de texto do número do pj recebe o foco
    let numero_pj_semformatacao = numero_pj.replace(/[.,-]/g, ''); //retira o "." e "-" do nº do processo
    await page.$eval('input[name$="numProcesso"]', (el, value) => el.value = value, numero_pj_semformatacao);
    await page.keyboard.press('Enter');  
    const navigationPromise = await page.waitForNavigation();
    let cadastrado = await page.$$eval('tbody[id*="dtTable_data"] tr td div[class="ui-dt-c"]', anchors => { return anchors.map(anchor => anchor.textContent)});
    await navigationPromise;
    let corte = 5;
    let dados_processo = [];
    let sub_lista = [];
    if (await cadastrado.length > 0) {
        let painel_caixas_disponiveis = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id*="dtTable_data"] tr td div[class="ui-dt-c"] a')).map(i=>{return i.id}));
        for (var i = 0; i < await cadastrado.length; i = i + corte) {
            let dados_lista = {"p": cadastrado[i], "c": cadastrado[i+1], "m": cadastrado[i+2]}
            await sub_lista.push(dados_lista);
        }
        if (await sub_lista.length == 1) {
            await dados_processo.push({"num": sub_lista[0].p, "classe": sub_lista[0].c, "link":'tbody[id*="dtTable_data"] tr td div a[id*="frmPesquisaProcessoJudicial:listaProcessosJudiciais:dtTable:0:"]', "indice": 0})
        } else if (await sub_lista.length > 1) {
            id_classe = await sub_lista.findIndex(element => element.p.substring(0,25) == numero_pj && element.c === classe); 
            if (await id_classe !== -1) {
                await dados_processo.push({"num": sub_lista[id_classe].p, "classe": sub_lista[id_classe].c, "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[id_classe]}"]`, "indice": id_classe})
            } else {
                let elementos_classe = await Object.keys(sub_lista).map(function (key) {//Coloca os valores da classe em um array 
                    return sub_lista[key].c;
                });
                for (let id_sub = 0; id_sub < await elementos_classe.length; id_sub++) {
                    if (await acoes.includes(elementos_classe[id_sub]) && sub_lista[id_sub].p.substring(0,25) == numero_pj) {
                        await dados_processo.push({"num": sub_lista[id_sub].p, "classe": sub_lista[id_sub].c, "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[id_sub]}"]`, "indice": id_sub})
                        id_sub = await elementos_classe.length + 1;
                    }
                }
            }
        }
        if (await dados_processo.length == 0) {
            let num = await page.$eval('tbody[id*="dtTable_data"] tr td div a[id*="frmPesquisaProcessoJudicial:listaProcessosJudiciais:dtTable:0:"]', element => element.textContent);
            let temp_c = await page.$eval('tbody[id*="dtTable_data"] tr td:nth-child(2) ', element => element.textContent);
            let so_num = await num.substring(0,25);
            await dados_processo.push({"num": so_num, "classe": temp_c, "link":'tbody[id*="dtTable_data"] tr td div a[id*="frmPesquisaProcessoJudicial:listaProcessosJudiciais:dtTable:0:"]', "indice": 0})
        }
        await page.waitForSelector(dados_processo[0].link);
        await page.click(dados_processo[0].link);
        await page.waitForSelector('img[id="graphicImageAguarde"]', {visible: false});
        await page.waitForSelector('div[id$="pnDetail_header"]', {visible: true});
    }
    return dados_processo;  
};

let dados_gerais = async (page, col) => {
    let mesa_procur;
    let procurador = ''
    let dados = [];
    let ultimo = '';
    let penultimo = '';
    let dt_ultimo = '';
    let dt_penultimo = '';
    let coluna_1 = await page.$$eval(`table[id*=":${col}:pgDadosBasicos"] tbody tr td[class*="coluna1"]`, anchors => { return anchors.map(anchor => anchor.textContent)});
    let coluna_2 = await page.$$eval(`table[id*=":${col}:pgDadosBasicos"] tbody tr td[class*="coluna2"]`, anchors => { return anchors.map(anchor => anchor.textContent)});
    const i_mesa = await coluna_1.findIndex(element => element === 'Processo na mesa de trabalho de:');// procura, no array da coluna 1, o índice, para verificar se está na mesa do procurador classe para pesquisa no array da coluna 2
    await i_mesa === -1 ? mesa_procur = await '' : mesa_procur = await coluna_2[i_mesa].split(" - ",3);
    if (mesa_procur !== '') {procurador = mesa_procur[1]}
    await page.waitForSelector('tbody[id*="manifestacaoTable"] tr td div[class*="ui-dt-c"]', { timeout: 120000 });
    let manifestacao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="manifestacaoTable_data"] tr td:nth-child(3n+1)')).map((el)=>{return el.innerText}));
    if (await manifestacao.length > 2) {
        ultimo = manifestacao[0];
        dt_ultimo = manifestacao[1]; 
        penultimo = manifestacao[2];
        dt_penultimo = manifestacao[3];
    } else if (await manifestacao.length == 2) {
        ultimo = manifestacao[0];
        dt_ultimo = manifestacao[1];
    }
    await dados.push({"procur": procurador, "ultimo": ultimo, "dt_ultimo": dt_ultimo, "penultimo": penultimo, "dt_penultimo": dt_penultimo});
    return dados;
};

let writeexcel = async (arq) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Relatório");
    worksheet.columns = [
        {header: 'Processo', key: 'processo', width: 25},
        {header: 'Classe judicial', key: 'classe', width: 32},
        {header: 'Triador', key: 'triador', width: 30},
        {header: 'Tipo Petição', key: 'tipo', width: 25},
        {header: 'Descrição', key: 'descricao', width: 30},
        {header: 'Último Registro', key: 'ultimo', width: 45},
        {header: 'Data último', key: 'data_ultimo', width: 10},
        {header: 'Penúltimo Registro', key: 'penultimo', width: 45},
        {header: 'Data penultimo', key: 'data_pen', width: 10},
        {header: 'Procurador', key: 'procurador', width: 30}
    ];
    for (let lin = 1; lin < 11; lin++) {worksheet.getColumn(lin).font = {name: 'Times New Roman', size: 10};}
    worksheet.getRow(1).font = {name: 'Calibri', size: 11, bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsulta.length; i++){
        worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, triador: listconsulta[i].triador, tipo: listconsulta[i].tipo_peticao,  descricao: listconsulta[i].descricao, ultimo: listconsulta[i].ultimo, data_ultimo: listconsulta[i].data_ultimo, penultimo: listconsulta[i].penultimo, data_pen: listconsulta[i].data_penultimo, procurador: listconsulta[i].procurador});
    }  
    await wbook.xlsx.writeFile(arq);
}

let scrape = async () => {
    let id_classe = '';
    let linha_excel;
    await console.log('Lendo arquivo excel ' + "\n");
    ler_output(listconsulta, 'OUTPUT - Noélia.xlsx');
    await ler_input();
    linha_excel = await compara_planilhas()
    await console.log('Lidos ' + (listprocessos.length - linha_excel) + ' Processos para pesquisa' + "\n");
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
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf', {waitUntil: 'networkidle0'}); //Ambiente de produção
    await dados_inicio(page); 
    await page.waitForSelector('input[name$="formMenus"]');
    let parametro;
    for (let i = linha_excel; i < await listprocessos.length; i++){
        id_classe = await -1;
        let classe_acao = '';
        let processo_cadastrado = await pesquisa_processo(page, listprocessos[i].numero, listprocessos[i].classe); //inserir processo para consulta e verifica se cadastrado
        if (processo_cadastrado.length > 0) {
            classe_acao = await processo_cadastrado[0].classe;
            parametro = await processo_cadastrado[0].classe.toUpperCase() + ' ' + processo_cadastrado[0].num;
            let detalhes = await page.evaluate(()=> Array.from(document.querySelectorAll('div[id*="pnDetail_header"]')).map(i=>{return i.innerText}));
            id_classe = await detalhes.findIndex(element => element === parametro);
            let parametros = await dados_gerais(page, id_classe, processo_cadastrado[0].classe);
            await listconsulta.push({numero: listprocessos[i].numero, classe: classe_acao, triador: listprocessos[i].triador, tipo_peticao: listprocessos[i].tipo_peticao, descricao: listprocessos[i].descricao, ultimo: parametros[0].ultimo, data_ultimo: parametros[0].dt_ultimo, penultimo: parametros[0].penultimo, data_penultimo: parametros[0].dt_penultimo, procurador: parametros[0].procur});
        } else {await listconsulta.push({numero: listprocessos[i].numero, classe: listprocessos[i].classe, triador: listprocessos[i].triador, tipo_peticao: listprocessos[i].tipo_peticao, descricao: listprocessos[i].descricao, ultimo: '', data_ultimo: '', penultimo: '', data_penultimo: '', procurador: ''});}
        await console.log(`Linha ${(i+1)}: ${listprocessos[i].numero} - ${classe_acao}`);
        await writeexcel('OUTPUT - Noélia.xlsx');
    }
    let result = 'Total de Processo pesquisados - Execução: ' + listprocessos.length;
    browser.close();
    return result
};  

scrape().then((value) => {
   console.log(value)
});