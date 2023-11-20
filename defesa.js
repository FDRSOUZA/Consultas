"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var path = require('path');
const downloadPath = path.resolve(__dirname);
const fs = require('fs');
let listprocessos = [];
let listconsulta = [];
let listrelatorio = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `OUTPUT - Defesa - ${dia}-${mes}-${ano}.xlsx`;

const instancia_1 = ['Procedimento do Juizado Especial Cível', 'Procedimento Comum', 'Mandado de Segurança', 'Mandado de Segurança Coletivo', 'Cumprimento de Sentença',
                'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento Provisório de Sentença', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Recurso (JEF)', 'Outras', 
                'Agravo de Instrumento', 'Mandado de Segurança Coletivo','Petição', 'Outras', 'Embargos de Terceiro', 'Embargos à Execução', 
                'Embargos à Execução Fiscal', 'Cautelar', 'Cautelar Fiscal', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Restauração de Autos', 'Monitória', 'Notificação', 
                'Habilitação', 'Habeas Data', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Ação Penal', 'Ação Popular', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 
                'Habeas Corpus', 'Liquidação por Arbitramento', 'Liquidação Provisória por Arbitramento', 'Liquidação Provisória pelo Procedimento Comum', 'Liquidação pelo Procedimento Comum', 
                'Liquidação Provisória de Sentença', 'Reintegração / Manutenção de Posse', 'Retificação de Registro de Imóvel', 'Restituição de Coisas Apreendidas', 'Exibição de Documento ou Coisa', 
                'Produção Antecipada da Prova', 'Procedimento Sumário', 'Protesto', 'Representação', 'Oposição', 'Depósito', 'Restituição de Coisa ou Dinheiro na Falência', 'Ação de Demarcação', 
                'Ação de Exigir Contas', 'Consignação em Pagamento', 'Depósito da Lei 8.866/94', 'Desapropriação', 'Despejo', 'Dissolução e Liquidação de Sociedade', 'Mandado de Injunção', 'Recuperação Judicial',
                'Embargos à Adjudicação', 'Embargos à Arrematação', 'Impugnação de Assistência Judiciária', 'Incidente de Suspeição', 'Impugnação de Crédito', 'Impugnação ao Pedido de Assist Litiscon Simples', 
                'Impugnação Ao Cumprimento de Sentença', 'Incidente de Impedimento', 'Incidente de Falsidade', 'Incidente de Desconsideração de Personalidade Jurídica', 'Pedido de Quebra Sigilo de Dados e/ou Telefônico', 
                'Insolvência Requerida pelo Devedor ou pelo Espólio', 'Dissolução e Liquidação de Sociedade', 'Dissolução Parcial de Sociedade', 'Exceção de Incompetência', 'Exceção de Litispendência', 'Falência', 
                'Embargos Infringentes na Execução Fiscal', 'Impugnação ao Valor da Causa',  'Carta Precatória', 'Carta Rogatória', 'Carta de Sentença', 'Carta de Ordem', 'Arrolamento', 'Inventário', 'Usucapião']
                 
                //'Ação Trabalhista', 'Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FNDE)', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];

const tipo_outras_origem = ['AÇÃO POPULAR', 'AÇÃO CIVIL PÚBLICA', 'AÇÃO CIVIL DE IMPROBIDADE ADMINISTRATIVA', 'OUTROS PROCEDIMENTOS DE JURISDIÇÃO VOLUNTÁRIA', 'AÇÃO CIVIL COLETIVA', 'CARTA PRECATÓRIA CÍVEL',  
        'CAUTELAR FISCAL', 'CAUTELAR INOMINADA', 'CUMPRIMENTO DE SENTENÇA CONTRA A FAZENDA PÚBLICA', 'CumSenFP', 'CUMPRIMENTO DE SENTENÇA', 'CUMPRIMENTO PROVISÓRIO DE SENTENÇA', 
        'EXECUÇÃO DE TÍTULO EXTRAJUDICIAL CONTRA A FAZENDA PÚBLICA', 'EMBARGOS À EXECUÇÃO', 'EMBARGOS À EXECUÇÃO FISCAL', 'EMBARGOS DE TERCEIRO CÍVEL', 'EXECUÇÃO CONTRA A FAZENDA PÚBLICA', 
        'HABEAS DATA', 'HABILITAÇÃO', 'INCIDENTE DE DESCONSIDERAÇÃO DE PERSONALIDADE JURÍDICA', 'LIQUIDAÇÃO DE SENTENÇA PELO PROCEDIMENTO COMUM', 'MS', 'MSCiv', 'MANDADO DE SEGURANÇA CÍVEL', 
        'MANDADO DE SEGURANÇA COLETIVO', 'EXECUÇÃO DE TÍTULO JUDICIAL - CEJUSC', 'PROCEDIMENTO COMUM CÍVEL', 'PROCEDIMENTO DO JUIZADO ESPECIAL CÍVEL', 'ProOrd',                   
        'RESTAURAÇÃO DE AUTOS', 'TUTELA CAUTELAR ANTECEDENTE', 'TUTELA ANTECIPADA ANTECEDENTE', 'USUCAPIÃO', 'LIQUIDAÇÃO POR ARBITRAMENTO', 'AÇÃO PENAL - PROCEDIMENTO ORDINÁRIO',  
        'ProceComCiv', 'CumSenFaz', 'Cumprimento de sentença', 'PROTESTO', 'PETIÇÃO CÍVEL', 'MONITÓRIA', 'LIQUIDAÇÃO PROVISÓRIA POR ARBITRAMENTO', 'PRODUÇÃO ANTECIPADA DA PROVA',  
        'RECLAMAÇÃO PRÉ-PROCESSUAL', 'REINTEGRAÇÃO / MANUTENÇÃO DE POSSE', 'EXIBIÇÃO DE DOCUMENTO OU COISA CÍVEL', 'OUTROS PROCEDIMENTOS DE JURISDIÇÃO VOLUNTÁRIA'];

const tipo_outras_relacionadas = ['Ação Popular', 'Ação Civil Pública', 'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Outras', 'Carta Precatória', 'Cautelar Fiscal', 
        'Cautelar', 'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento de Sentença', 'Cumprimento Provisório de Sentença',  
        'Execução de Título Extrajudicial contra a Fazenda Pública', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Execução contra a Fazenda Pública (art. 730, CPC/73)',  
        'Habeas Data', 'Habilitação', 'Incidente de Desconsideração de Personalidade Jurídica', 'Cumprimento de Sentença contra a Fazenda Pública', 'Mandado de Segurança', 'Mandado de Segurança', 
        'Mandado de Segurança', 'Mandado de Segurança Coletivo', 'Outras', 'Procedimento Comum', 'Procedimento do Juizado Especial Cível', 'Procedimento Comum', 'Restauração de Autos', 
        'Tutela de Cautelar Antecedente', 'Tutela Antecipada Antecedente', 'Usucapião', 'Liquidação por Arbitramento', 'Ação Penal', 'Procedimento Comum',  
        'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento de Sentença', 'Protesto', 'Petição', 'Monitória', 'Liquidação Provisória por Arbitramento', 'Produção Antecipada da Prova', 'Outras', 
        'Reintegração / Manutenção de Posse', 'Exibição de Documento ou Coisa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Alvará/Outros Procedimentos Jurisdição Voluntária'];

const tipo_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal (FNDE)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (informações do débito pendente)'];
const tipo_instacia = ['STJ', 'STF', 'TRT', 'TURMA RECURSAL']; 

let dados_inicio = async (page) => {
    var filePath = path.resolve(__dirname, 'Dados.xlsx');
    if (fs.existsSync(filePath)) { 
        var wb = new Excel.Workbook();
        await wb.xlsx.readFile(filePath);
        let worksheet = await wb.getWorksheet('SAJ');
        let linha_ind = await worksheet.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
        await page.waitForSelector('input[id="frmLogin:username"]');
        await page.type('input[id="frmLogin:username"]', linha_ind[1]);
        await page.type('input[id="frmLogin:password"]', linha_ind[2]);
        await page.click('button[id="frmLogin:entrar"]');
    }
};

let ler_relatorio = async () => {
    const classes_2 = ['Apelação', 'Apelação / Remessa Necessária', 'Agravo de Petição', 'Agravo de Instrumento em Agravo de Petição', 'Agravo Interno', 'Agravo de Instrumento',
        'Agravo de Instrumento em Recurso de Revista', 'Agravo de Instrumento em Recurso de Revista', 'Agravo de Instrumento em Recurso Extraordinário', 'Agravo em Recurso Especial',
        'Agravo Regimental', 'Remessa Necessária', 'Recurso Ordinário (outros)', 'Recurso (JEF)',  
    ];
    const fixo = ['Número do processo CNJ', 'Classe SAJ', 'Procurador da última distribuição', 'Distribuído no SAJ', 'Instância', 'Juízo', 'Matéria', 'Instância', 'Valor da causa'];
    var wb = new Excel.Workbook();
    var filePath = path.resolve(__dirname, 'RelatorioDadosDoProcesso.xlsx');
    await wb.xlsx.readFile(filePath);
    let worksheet = wb.getWorksheet('Relatórios');
    let planilha_lida = new Array(); 
    worksheet.eachRow({ includeEmpty: false }, async (row) => {//ler cada uma das linhas da planilha Relatorio
        const container = await{...row.values};   
        await planilha_lida.push(container) //Salva cada linha em um array
    }) 
    
    let linha_ind = await worksheet.getRow(1).values;//Salva os dados da 1ª linha da planilha para selecionar as colunas que interessa
    await planilha_lida.shift()//Exclcui os valores da 1ª linha da planilha
    let coluna = []; 
    await fixo.forEach(function(elemento) {//Define os indices que serão manipulados - quais colunas
        let id_relacionado = linha_ind.findIndex(element => element === elemento);
        if(id_relacionado !== -1) {coluna.push(id_relacionado)}
    });
    let planilha_alterada = new Array();
    planilha_alterada = await planilha_lida.map(obj => {//Converte os dados lidos da planilha em um array de objetos definindo as keys
        let classe;
        classe = obj[coluna[1]].slice(obj[coluna[1]].indexOf('-')+2, obj[coluna[1]].length);
        let prevento; 
        obj[coluna[3]] === 'Sim' ? prevento = obj[coluna[2]] : prevento = '';
        let materia;
        obj[coluna[6]] !== ' - ' ? materia = obj[coluna[6]] : materia = '';
        return {
            numero: obj[coluna[0]],
            classe: classe,
            prevento: prevento,
            materia: materia,
            bruta: materia, // Matéria sem tratamento
            quant: '',
            valor_causa: obj[coluna[8]],
            juizo: obj[coluna[4]],
            registro: '',
            instancia: obj[coluna[7]],
        }
    });
    let planilha_sem_superiores = await planilha_alterada.filter(function(obj) { //Exclui as classes que não são da triagem execução
        return (!tipo_instacia.includes((obj.instancia))); 
    });
    
    planilha_lida = await planilha_sem_superiores.filter(function(obj) { //Exclui as classes que não são da triagem execução
        return (!classes_2.includes(obj.classe)) 
    });
    await planilha_lida.map(function(elemento) {//Define os indices que serão manipulados - quais colunas
        if(elemento.materia !== ' - ') {
            let materia_bruta = elemento.materia.split("\n").filter(function (i) {//o novo array recebe os dados sem linhas em branco e o filtro é para tirar os elementos com valor vazio
                return i;
            }); 
            let temp = [];
            materia_bruta.map(function(bruta) {
                let mb = bruta.slice(0, bruta.indexOf(' '));
                temp.push(mb); 
            });
            elemento.materia = temp.join("; ");
        } //else {elemento.materia = '';}
    });
    return planilha_lida;          
};

let aguarda_download = async () => { //Verifica se o arquivo já terminou de ser baixado.
    await console.log('Aguardando download do arquivo ...');
    var filePath = path.resolve(__dirname, 'RelatorioDadosDoProcesso.xlsx');
    let sit = ''
    do {
        await fs.existsSync(filePath) ? sit = 'Baixou' : sit = '';
    } while (sit == '')
    await console.log('Arquivo baixado com sucesso!')
}

//########## LER A PLANILHA COM OS PROCESSO A SEREM PESQUISADOS ############### 
let ler_input = async () => {
    let sheets = [];
    let wb = new Excel.Workbook();
    var filePath = path.resolve(__dirname, 'Defesa.xlsx');
    await wb.xlsx.readFile(filePath);
    await wb.eachSheet(function (worksheet) {
        sheets.push(worksheet.name); //Coloca o nome das abas da planilha em um array
    });
    await sheets.map(item => {
        let worksheet = wb.getWorksheet(item); 
        let acao;
        worksheet.eachRow({ includeEmpty: false }, async (row) => {//ler cada uma das linhas da planilha
            let id_relacionado = tipo_outras_origem.findIndex(element => element === row.values[2]);
            id_relacionado !== -1 ? acao = tipo_outras_relacionadas[id_relacionado] : acao = row.values[2];
            await listprocessos.push({"numero":row.values[1], "classe": acao}) //Salva cada linha em um array
        }) 
    });
    //await listprocessos.shift();
    await console.log('Arquivo input lido' + "\n");
};

let pesquisa_processo = async (page, numero_pj, classe) => {
    let autuacao;
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/pesquisarProcessoJudicial.jsf?'); 
    await page.waitForSelector('input[name$="numProcesso"]'); //Aguarda carregar a página da "Consulta"
    await page.focus('input[name$="numProcesso"]'); //caixa de texto do número do pj recebe o foco
    let numero_pj_semformatacao = numero_pj.replace(/[.,-]/g, ''); //retira o "." e "-" do nº do processo
    await page.$eval('input[name$="numProcesso"]', (el, value) => el.value = value, numero_pj_semformatacao);
    await page.keyboard.press('Enter');  
    const navigationPromise = await page.waitForNavigation();
    let cadastrado = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(2)', anchors => { return anchors.map(anchor => anchor.textContent)});
    await navigationPromise;
    let dados_processo = [];
    if (await cadastrado.length > 0) {
        let painel_caixas_disponiveis = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id*="dtTable_data"] tr td div[class="ui-dt-c"] a')).map(i=>{return i.id}));
        let numeros = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(1)', anchors => { return anchors.map(anchor => anchor.textContent)});
        if (cadastrado.some(el => instancia_1.includes(el)) || cadastrado.some(el => tipo_execucao.includes(el))) { //Se tiver uma classe de execução ou de 1º instância
            let id_classe = await cadastrado.findIndex(element => element === classe); 
            if (await id_classe !== -1) {
                dados_processo.push({"num": numeros[id_classe], "classe": cadastrado[id_classe], "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[id_classe]}"]`, "indice": id_classe, "quant": cadastrado.length})
            } 
            if (await dados_processo.length > 0) {
                await page.waitForSelector(dados_processo[0].link);
                await page.click(dados_processo[0].link);
                await page.waitForSelector('img[id="graphicImageAguarde"]', {visible: false});
                await page.waitForSelector('div[id$="pnDetail_header"]', {visible: true});
                await page.waitForSelector('tbody[id*="manifestacaoTable_data"] tr td div[class="ui-dt-c"]', {visible: true});
                let autuacoes_processuais = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id*="manifestacaoTable_data"] tr td div[class="ui-dt-c"]')).map(i=>{return i.innerText}));
                if (await autuacoes_processuais.length > 1) {autuacao = await (autuacoes_processuais[0] + '; ' + autuacoes_processuais[1] + '; ' + autuacoes_processuais[3])}
            }
        }
    }
    return autuacao;  
};

let pesquisa_planilha = async () => {
    let acoes;
    //#######  Filtra os processos que estão apenas na planilha input, não estão no relatório do SAJ - logo, NÃO ESTÃO CADASTRADOS
    let nao_cadastrados = await listprocessos.filter(function(obj) {
        return (!listrelatorio.some(t => t.numero.includes(obj.numero)));     
    });
    
    //#######  Salva os processos que estão apenas na planilha input, não estão no relatório do SAJ - logo, NÃO ESTÃO CADASTRADOS
    await nao_cadastrados.forEach(function(obj) {//Define os indices que serão manipulados - quais colunas
        listconsulta.push({numero: obj.numero, classe: obj.classe, prevento: '', juizo: '', registro:'', cadastrado: 'não', pesquisa: 'Finalizado'})
    });

    //#######  Filtra os processos que estão nas 2 planilha - input e relatório do SAJ - logo, ESTÃO CADASTRADOS - Separando somente do de Execução Fiscal
    let cadastrados = await listprocessos.filter(function(obj) {
        return (listrelatorio.some(t => t.numero.includes(obj.numero))) 
    });

    await cadastrados.map(item => {//PROCESSOS CADASTRADOS E CLASSE DIFERENTE DE EXECUÇÃO FISCAL
        acoes = instancia_1.concat(tipo_execucao);
        let elementos = listrelatorio.filter(x => x.numero === item.numero);//Filtra pelo número do processo
        let indice = elementos.findIndex(element => element.classe === item.classe);//Pesquisa para ver a classe coincidente
        if (indice !== -1) { //Se encontrou salva no array de consulta
            listconsulta.push({numero: elementos[indice].numero, classe: elementos[indice].classe, prevento: elementos[indice].prevento, materia: elementos[indice].materia, bruta: elementos[indice].bruta, juizo: elementos[indice].juizo, registro: elementos[indice].registro, valor_causa: elementos[indice].valor_causa, cadastrado: 'sim', pesquisa: 'Pendente'}); 
        } else {//Se não encontrou procura a classe
            let elementos_classe = Object.keys(elementos).map(function (key) {//Coloca os valores da classe em um array 
                return elementos[key].classe;
            });
            let classe_saj;
            acoes.some(el => {//Pesquisa o nome da classe no array
                if(elementos_classe.includes(el)) {
                    classe_saj = el;
                }
            });
            indice = elementos_classe.findIndex(element => element === classe_saj);
            if (indice == -1) {
                indice = 0;
            }
            listconsulta.push({numero: elementos[indice].numero, classe: elementos[indice].classe, prevento: elementos[indice].prevento, materia: elementos[indice].materia, bruta: elementos[indice].bruta, juizo: elementos[indice].juizo, registro: elementos[indice].registro, valor_causa: elementos[indice].valor_causa, cadastrado: 'sim', pesquisa: 'Pendente'});
        }
    }); 
    await writeexcel(nome_arquivo_excel, 'Provisória');
};

let apagar_relatorio = async () => {
    var filePath = await path.resolve(__dirname, 'RelatorioDadosDoProcesso.xlsx');
    if (fs.existsSync(filePath)) { 
        try {
            fs.unlinkSync(filePath)
          } catch(err) {console.error(err)}
    }
};

let dividir_lote = async (page) => {
    let corte_final;
    let r = 0;
    let corte_inicial = 0;
    let repeticao = await (Math.ceil(listprocessos.length/5000));
    await repeticao === 1 ? corte_final = listprocessos.length : corte_final = 5000;
    let ids_limpar = await page.evaluate(()=> Array.from(document.querySelectorAll('table[style="padding-bottom: 5px"] tbody tr td button')).map((el)=>{return el.id}));//lista as ids do botão
    let botao_limpar = await page.evaluate(()=> Array.from(document.querySelectorAll('table[style="padding-bottom: 5px"] tbody tr td')).map((el)=>{return el.innerText}));//lista as ids das tarefas do painel    await repeticao === 1 ? corte_final = listprocessos.length : corte_final = 5000;
    let id_limpar = await botao_limpar.findIndex(element => element === 'LIMPAR');
    do {
        let lote_processo = await listprocessos.slice(corte_inicial, corte_final);
        let lista = await lote_processo.map(el => el.numero).join(',');
        const lista_acabada = await lista.replace(/,+/g, "\n");
        await console.log('Quantidade de processos no lote: '+lote_processo.length);
        await page.evaluate((lista_acabada) => {document.querySelector('tbody tr td textarea[id="formRelatorioDadosDoProcesso:listaNrCpfCnpj"]').value = lista_acabada}, lista_acabada);//Cola a lista de processos no form
        await page.click('button[id="formRelatorioDadosDoProcesso:btnExportar"] span');//Clica no botão exibir Exportar
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"'); //Aguarda o form de "Aguarde o processamento" ficar invisível
        await page.waitForFunction('document.getElementById("formRelatorioDadosDoProcesso:dialogGerarRelatorio").style.visibility === "visible"'); //Aguarda o form de "Aguarde o processamento" ficar invisível
        await page.click('table[id="formRelatorioDadosDoProcesso:smTipoRecebimentoRelatorio"] tbody tr td div div input[name="formRelatorioDadosDoProcesso:smTipoRecebimentoRelatorio"]');
        await page._client.send('Page.setDownloadBehavior', {
        //await page._client().send('Page.setDownloadBehavior', {
            behavior: 'allow',
            downloadPath: downloadPath 
        });
        let ids_ok = await page.evaluate(()=> Array.from(document.querySelectorAll('div[id="formRelatorioDadosDoProcesso:dialogGerarRelatorio"] div div[class="containerBotao centraliza"] button')).map((el)=>{return el.id}));//lista as ids do botão
        let botao_ok = await page.evaluate(()=> Array.from(document.querySelectorAll('div[id="formRelatorioDadosDoProcesso:dialogGerarRelatorio"] div div[class="containerBotao centraliza"] button')).map((el)=>{return el.innerText}));//lista as ids das tarefas do painel
        let id_ok = await botao_ok.findIndex(element => element === 'OK');
        await page.click(`button[id="${ids_ok[id_ok]}"] span`);
        await aguarda_download();
        await new Promise((r) => setTimeout(r, 500));
        listrelatorio = await ler_relatorio();
        await page.waitForFunction('document.getElementById("statusAguarde").style.visibility === "hidden"'); //Aguarda o form de "Aguarde o processamento" ficar invisível
        await new Promise((r) => setTimeout(r, 500));
        await page.click(`button[id="${ids_limpar[id_limpar]}"] span`);
        await new Promise((r) => setTimeout(r, 2000));
        await apagar_relatorio(); 
        corte_inicial = await corte_final;
        await r++;
        await r == (repeticao-1) ? corte_final = await listprocessos.length : corte_final = await (corte_final+5000);   
    } while (r < repeticao);
    await console.log('___________________________________________________________________________________ \n');
}

let writeexcel = async (nome_arquivo, coluna) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Auxiliar");
    worksheet.columns = [
        {header: 'Processo', key: 'processo', width: 25},
        {header: 'Classe judicial', key: 'classe', width: 35},
        {header: 'Procurador prevento', key: 'procurador', width: 40},
        {header: 'Matéria SAJ', key: 'materia', width: 35},
        {header: 'Matéria Bruta', key: 'num', width: 50},
        {header: 'Juízo', key: 'juizo', width: 15},
        {header: 'Atuacao', key: 'autuacao', width: 35},
        {header: 'Valor da Causa', key: 'valor', width: 15},
        {header: '', key: 'finalizado', width: 10}
        //{header: 'Finalizado', key: 'finalizado', width: 10}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsulta.length; i++){
        if (coluna == 'Provisória') {
            worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento,  materia: listconsulta[i].materia, num: listconsulta[i].bruta, juizo: listconsulta[i].juizo, autuacao: listconsulta[i].registro, valor: listconsulta[i].valor_causa, finalizado: listconsulta[i].pesquisa}); //loop para escrever o nome dos procuradores e processos na planilha exportada
        } else {
            worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento,  materia: listconsulta[i].materia, num: listconsulta[i].bruta, juizo: listconsulta[i].juizo, autuacao: listconsulta[i].registro, valor: listconsulta[i].valor_causa, finalizado: ''}); //loop para escrever o nome dos procuradores e processos na planilha exportada
        }    
    }  
    await wbook.xlsx.writeFile(nome_arquivo);
}

let ler_output = async (arq) => {
    let wb = new Excel.Workbook();
    let filePath = path.resolve(__dirname, arq);
    if (fs.existsSync(filePath)) {
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Auxiliar'); // Primeira aba do arquivo excel - Planilha
        for (let i = 2; i <= await sh.rowCount; i++) { //Começa a ler da linha 2
            await listconsulta.push({numero: sh.getRow(i).getCell(1).value, classe: sh.getRow(i).getCell(2).value, prevento: sh.getRow(i).getCell(3).value, materia: sh.getRow(i).getCell(4).value, acoes: sh.getRow(i).getCell(5).value, juizo: sh.getRow(i).getCell(6).value, registro: sh.getRow(i).getCell(7).value, pesquisa: sh.getRow(i).getCell(8).value});
        }
    }
};

let scrape = async () => {
    await ler_output(nome_arquivo_excel);
    await console.log('Lendo arquivo excel ' + "\n");
    await ler_input();
    //await console.log('Lida Planilha Relatório de Dados de Processo');
    await console.log('Total de ' + (listprocessos.length) + ' Processos lidos no Input' + "\n");
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        //executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        headless: false, args:['--start-maximized',  '--disable-features=site-per-process'], //torna visível e maximiza a 
        ignoreHTTPSErrors: true,
    });
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf'); //Ambiente de produção
    //await dados_inicio(page);
    await page.waitForNavigation();
    await page.waitForSelector('input[name$="formMenus"]');
    if (await listconsulta.length == 0) {
        await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/relatorio/relatorioDadosDoProcesso.jsf?'); 
        await dividir_lote(page);
        await pesquisa_planilha();
    }  
    let array_registro = await listconsulta.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
        return obj.pesquisa === 'Pendente';
    });
    await console.log('Selecionados ' + (array_registro.length) + ' Processos para Pesquisa de um total de '+ listconsulta.length);
    if (await array_registro.length > 0) {
        await console.log('Pesquisa dos registros de autuação');
        for (let i = 0; i < await array_registro.length; i++) {
            let id_el = await listconsulta.findIndex(x => x.numero === array_registro[i].numero);
            let dados_autuacao = await pesquisa_processo (page, array_registro[i].numero, array_registro[i].classe);
            if (await dados_autuacao !== '') {
                listconsulta[id_el].pesquisa = await 'Finalizado';
                listconsulta[id_el].registro = await dados_autuacao;
                await console.log(`linha ${(id_el+1)} / ${listconsulta.length} : ${array_registro[i].numero}`)
            }
            await writeexcel(nome_arquivo_excel, 'Provisória');
        }    
    }
    await writeexcel(nome_arquivo_excel, 'Final');
    let result = await 'Pesquisa Concluída ';
    browser.close();
    return result
};    

scrape().then((value) => {
   console.log(value)
});