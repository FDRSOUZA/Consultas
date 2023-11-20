"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var path = require('path');
//const downloadsFolder = require('downloads-folder');
const downloadPath = path.resolve(__dirname);
const fs = require('fs');
let listprocessos = [];
let listconsulta = [];
let listrelatorio = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `OUTPUT - Execução - ${dia}-${mes}-${ano}.xlsx`;
//console.log(downloadsFolder());

const tipo_execucao = ['Execução Fiscal (SIDA)', 'Execução Fiscal Previdenciária', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)',
    'Execução Fiscal (informações do débito pendente)', 'Execução Fiscal (FNDE)'];

const tipo_outras_origem = ['CARTA PRECATÓRIA CÍVEL', 'CAUTELAR FISCAL', 'CUMPRIMENTO DE SENTENÇA', 'CUMPRIMENTO DE SENTENÇA CONTRA A FAZENDA PÚBLICA', 'Cumprimento de sentença', 'CumSen', 'DESAPROPRIAÇÃO', 
    'EMBARGOS À EXECUÇÃO', 'EMBARGOS À EXECUÇÃO FISCAL', 'EMBARGOS DE TERCEIRO CÍVEL', 'EMBARGOS DE TERCEIRO', 'EXECUÇÃO DE TÍTULO EXTRAJUDICIAL', 'INCIDENTE DE DESCONSIDERAÇÃO DE PERSONALIDADE JURÍDICA', 
    'RESTAURAÇÃO DE AUTOS', 'EE', 'ALIENAÇÃO JUDICIAL DE BENS', 'CumSenFaz', 'EXECUÇÃO CONTRA A FAZENDA PÚBLICA', 'ArrCom', 'Arrolamento Comum','ARROLAMENTO SUMÁRIO', 'Arrolamento Sumário', 'ACIA', 'Alvará Judicial - Lei 6858/80', 
    'ETIJ', 'ETCiv', 'ET', 'PJEC', 'ResAutCiv', 'CartOrdCiv', 'CartPrecCiv', 'CumSen', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro Cível', 'Execução Contra a Fazenda Pública', 'FESEMEPP', 
    'Falencia', 'Invent', 'Inventário', 'PetCiv', 'ProceComCiv', 'Procedimento Comum Cível', 'RecJud', 'TutAntAnt', 'Usucap', 'Usucapião', 'USUCAPIÃO', 'CauFis', 'TutCautAnt', 'Falência de Empresários, Sociedades Empresáriais, Microempresas e Empresas de Pequeno Porte',   
    'Desapr', 'CONSIGNAÇÃO EM PAGAMENTO', 'ConPag', 'HabCre', 'HTE', 'Demarcação / Divisão', 'ACPCiv', 'Oposic', 'EEFis', 'PCE', 'Sobrepartilha', 'ECFP', 'MSCiv', 'ArrSum', 'OPJV', 'Rp', 'APEl', 'Execução Fiscal', 
    'ExFis', 'ExTiEx', 'CumPrSe', 'RelFal', 'Monito', 'ExcInc', 'RtMtPosse'];

const tipo_outras_relacionadas = ['Carta Precatória', 'Cautelar Fiscal', 'Cumprimento de Sentença', 'Cumprimento de Sentença contra a Fazenda Pública', 'Cumprimento de Sentença', 'Cumprimento de Sentença',
    'Desapropriação', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Execução de Título Extrajudicial', 'Incidente de Desconsideração de Personalidade Jurídica', 
    'Restauração de Autos', 'Embargos à Execução', 'Outras', 'Cumprimento de Sentença contra a Fazenda Pública', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Arrolamento', 'Arrolamento', 'Arrolamento', 
    'Arrolamento', 'Ação de Improbidade Administrativa', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Embargos de Terceiro', 'Procedimento do Juizado Especial Cível', 
    'Restauração de Autos', 'Carta de Ordem', 'Carta Precatória', 'Cumprimento de Sentença', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Execução contra a Fazenda Pública (art. 730, CPC/73)',
    'Falência', 'Falência', 'Inventário', 'Inventário', 'Petição', 'Procedimento Comum', 'Procedimento Comum', 'Recuperação Judicial', 'Tutela Antecipada Antecedente', 'Usucapião', 'Usucapião', 'Usucapião', 'Cautelar', 
    'Tutela Antecipada Antecedente', 'Falência', 'Desapropriação', 'Consignação em Pagamento', 'Consignação em Pagamento', 'Habilitação', 'Habilitação', 'Reintegração / Manutenção de Posse', 'Ação Civil Pública', 
    'Oposição', 'Embargos à Execução Fiscal', 'Outras', 'Inventário', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Mandado de Segurança', 'Arrolamento', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 
    'Representação', 'Ação Penal', 'EXECUÇÃO FISCAL', 'EXECUÇÃO FISCAL', 'EXECUÇÃO FISCAL', 'Cumprimento Provisório de Sentença', 'Falência', 'Monitória', 'Exceção de Incompetência', 'Reintegração / Manutenção de Posse' ];

const instancia_1 = ['Ação Trabalhista', 'Arrolamento', 'Procedimento Comum', 'Cumprimento de Sentença', 'Cumprimento de Sentença contra a Fazenda Pública', 'Carta Precatória',  
    'Usucapião', 'Consignação em Pagamento', 'Desapropriação', 'Embargos à Execução', 'Embargos à Execução Fiscal', 'Embargos de Terceiro', 'Alvará/Outros Procedimentos Jurisdição Voluntária', 'Ação de Improbidade Administrativa', 
    'Execução de Título Extrajudicial', 'Execução contra a Fazenda Pública (art. 730, CPC/73)', 'Exibição de Documento ou Coisa', 'Falência', 'Habilitação', 'Incidente de Desconsideração de Personalidade Jurídica',   
    'Inventário', 'Outras', 'Procedimento do Juizado Especial Cível', 'Protesto', 'Petição', 'Reclamação', 'Recuperação Judicial', 'Representação', 'Restauração de Autos', 'Cautelar', 'Cautelar Fiscal',
    'Restituição de Coisa ou Dinheiro na Falência', 'Reintegração / Manutenção de Posse', 'Tutela Antecipada Antecedente', 'Tutela de Cautelar Antecedente', 'Oposição', 'Ação Penal',  
    'Cumprimento Provisório de Sentença', 'Embargos à Execução de Título Extrajudicial', 'Monitória', 'Exceção de Incompetência'];


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
            
let separar_cnpj_cpf = async (conteudo_celula) => {
    let cnpj_cpf = '';
    function somentenumeros (value) {
        let x = value.replace(/[.,/,-]/g, '');
        if (x.match("[0-9]*")) {
            if (x.length === 11 || x.length === 14 && x !== '00394460021653') {
                return x;
            }
        }
    }

    let numeros = await conteudo_celula.filter(somentenumeros);
    
    if (numeros.length > 0) {
        let obj_final = await numeros.reduce(function(valorAcumulador, valorArray) {
            const tipo = valorArray.length === 18 ? 'cnpj' : 'cpf'
            valorAcumulador[tipo].push(valorArray)
            return valorAcumulador;
        }, {cnpj: [], cpf: []});
        if (await obj_final.cnpj.length > 0) {
            cnpj_cpf = await obj_final.cnpj[0]
        } else if (await obj_final.cpf.length > 0) {
            cnpj_cpf = await obj_final.cpf[0]
        }
    }
    return cnpj_cpf;
};

let ler_relatorio = async () => {
    const classes_2 = ['Apelação', 'Apelação / Remessa Necessária', 'Agravo de Petição', 'Agravo de Instrumento em Agravo de Petição', 'Agravo Interno', 'Agravo de Instrumento',
        'Agravo de Instrumento em Recurso de Revista', 'Agravo de Instrumento em Recurso de Revista', 'Agravo de Instrumento em Recurso Extraordinário', 
        'Agravo Regimental', 'Mandado de Segurança', 'Mandado de Segurança Coletivo', 'Remessa Necessária', 'Recurso Ordinário (outros)', 'Recurso (JEF)', 'Procedimento do Juizado Especial Cível', 
    ];
    const tipo_instacia = ['TRT', 'STJ', 'STF', 'TURMA RECURSAL'];
    const fixo = ['Número do processo CNJ', 'Classe SAJ', 'Polo passivo', 'Instância', 'Valor da causa', 'Procurador da última distribuição', 'Distribuído no SAJ', 'Localidade', 'Juízo', 'Inscrições SIDA processo principal', 'Somatório inscrições SIDA (processo principal e vinculados)', 'Controle de prescrição intercorrente', 'Rating grupo', 'Demanda analytics', 'Processos vinculados']
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
        let processos_vinculados;
        classe = obj[coluna[1]].slice(obj[coluna[1]].indexOf('-')+2, obj[coluna[1]].length);
        let prevento; 
        obj[coluna[6]] === 'Sim' ? prevento = obj[coluna[5]] : prevento = '';
        let prescricao;
        obj[coluna[11]] !== ' - ' ? prescricao = obj[coluna[11]] : prescricao = '';
        let inscricoes;
        obj[coluna[9]] !== ' - ' ? inscricoes = obj[coluna[9]].replace(/\|/g,";").split(" ; ").join("; ") : inscricoes = '';
        let somatorio = ''
        if (tipo_execucao.includes(classe) && obj[coluna[10]] !== ' - ') {
            obj[coluna[10]] !== '0,00' ? somatorio = obj[coluna[10]].replace(/\./g, "") : somatorio = obj[coluna[10]];
        } else if (!tipo_execucao.includes(classe) && obj[coluna[4]] !== ' - ') {
            obj[coluna[4]] == '0' ? somatorio = '0,00' : somatorio = String(obj[coluna[4]]).replace('.',',');
        }
        let parte = '';
        obj[coluna[2]] !== ' - ' ? parte = obj[coluna[2]].replace(/\n/g," - ").split(" - ") : parte = '';
        let rating_grupo 
        obj[coluna[12]] !== ' - ' ? rating_grupo =  obj[coluna[12]].replace(/\n/g," ").replace(/\;/g,"; ") : rating_grupo = '';
        let analytics;
        obj[coluna[13]] !== ' - ' ? analytics = obj[coluna[13]].replace(/\n/g," ") : analytics = '';
        obj[coluna[14]] !== ' - ' ? processos_vinculados = obj[coluna[14]] : processos_vinculados = '';
        return {
            numero: obj[coluna[0]],
            classe: classe,
            prevento: prevento,
            cda: inscricoes,
            prescricao: prescricao,
            valor: somatorio,
            parte: parte, 
            juizo: obj[coluna[3]],
            instancia: obj[coluna[3]], 
            registro: '',
            grupo: rating_grupo,
            demanda: analytics,
            vinculados: processos_vinculados,
        }
    });
    let planilha_sem_superiores = await planilha_alterada.filter(function(obj) { //Exclui as classes que não são da triagem execução
        return (!tipo_instacia.includes((obj.instancia))); 
    });
    planilha_lida = await planilha_sem_superiores.filter(function(obj) { //Exclui as classes que não são da triagem execução
        return (!classes_2.includes(obj.classe)) 
    });
    
    return await planilha_lida;          
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
    var filePath = path.resolve(__dirname, 'Execução.xlsx');
    let wb = new Excel.Workbook();
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
            } else if (await id_classe == -1 && cadastrado.length == 1) {
                dados_processo.push({"num": numeros[0], "classe": cadastrado[0], "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[0]}"]`, "indice": 0, "quant": cadastrado.length})
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
        listconsulta.push({numero: obj.numero, classe: obj.classe, prevento: '', cda: '', prescricao: '', valor: '', parte: '', juizo: '', registro:'', grupo: '', demanda: '', vinculados: '', cadastrado: 'não', pesquisa: 'Finalizado'})
    });
    //#######  Filtra os processos que estão nas 2 planilha - input e relatório do SAJ - logo, ESTÃO CADASTRADOS - Separando somente do de Execução Fiscal
    let cadastrados = await listprocessos.filter(function(obj) {
        return (listrelatorio.some(t => t.numero.includes(obj.numero)) && obj.classe == 'EXECUÇÃO FISCAL') 
    });

    await cadastrados.map(item => {//PROCESSOS CADASTRADOS DE EXECUÇÃO FISCAL
        acoes = tipo_execucao.concat(instancia_1);
        let elementos = listrelatorio.filter(x => x.numero === item.numero);//Filtra pelo número do processo
        let elementos_classe = Object.keys(elementos).map(function (key) {//Coloca os valores da classe em um array 
            return elementos[key].classe;
        });
        let classe_saj;
        for (let i = 0; i < elementos_classe.length; i++) {
            if (acoes.includes(elementos_classe[i])) {
                classe_saj = elementos_classe[i];
                i = elementos_classe.length + 1;
            }
        }
        
        let indice = elementos_classe.findIndex(element => element === classe_saj);
        if (indice == -1) {//Se a descrição da classe não existir no array acoes
            indice = 0;
        }
        let sit;
        elementos[indice].classe === 'Execução Fiscal (SIDA)' ? sit = 'Finalizado' : sit = 'Pendente'
        listconsulta.push({numero: elementos[indice].numero, classe: elementos[indice].classe, prevento: elementos[indice].prevento, cda: elementos[indice].cda, prescricao: elementos[indice].prescricao, valor: elementos[indice].valor, parte: elementos[indice].parte, juizo: elementos[indice].juizo, registro: elementos[indice].registro, grupo: elementos[indice].grupo, demanda: elementos[indice].demanda, vinculados: elementos[indice].vinculados, cadastrado: 'sim', pesquisa: sit}); 
    });

    cadastrados = await listprocessos.filter(function(obj) {
        return (listrelatorio.some(t => t.numero.includes(obj.numero)) && obj.classe !== 'EXECUÇÃO FISCAL') 
    });
    
    await cadastrados.map(item => {//PROCESSOS CADASTRADOS E CLASSE DIFERENTE DE EXECUÇÃO FISCAL
        acoes = instancia_1.concat(tipo_execucao);
        let elementos = listrelatorio.filter(x => x.numero === item.numero);//Filtra pelo número do processo
        let indice = elementos.findIndex(element => element.classe === item.classe);
        if (indice !== -1) {
            listconsulta.push({numero: elementos[indice].numero, classe: elementos[indice].classe, prevento: elementos[indice].prevento, cda: elementos[indice].cda, prescricao: elementos[indice].prescricao, valor: elementos[indice].valor, parte: elementos[indice].parte, juizo: elementos[indice].juizo, registro: elementos[indice].registro, grupo: elementos[indice].grupo, demanda: elementos[indice].demanda, vinculados: elementos[indice].vinculados, cadastrado: 'sim', pesquisa: 'Pendente'}); 
        } else {
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
            listconsulta.push({numero: elementos[indice].numero, classe: elementos[indice].classe, prevento: elementos[indice].prevento, cda: elementos[indice].cda, prescricao: elementos[indice].prescricao, valor: elementos[indice].valor, parte: elementos[indice].parte, juizo: elementos[indice].juizo, registro: elementos[indice].registro, grupo: elementos[indice].grupo, demanda: elementos[indice].demanda, vinculados: elementos[indice].vinculados, cadastrado: 'sim', pesquisa: 'Pendente'});
        }
    }); 
    for (let i = 0; i < await listconsulta.length; i++) {
        if (listconsulta[i].parte !== '') {
            let cnpj_cpf = await listconsulta[i].parte;
            let consultado = await separar_cnpj_cpf(cnpj_cpf);
            listconsulta[i].parte = await consultado
        }
    }
    //await writeexcel(nome_arquivo_excel, 'Provisória');
};

let apagar_relatorio = async () => {
    var filePath = await path.resolve(__dirname, 'RelatorioDadosDoProcesso.xlsx');
    if (fs.existsSync(filePath)) { 
        try {
            fs.unlinkSync(filePath)
          } catch(err) {console.error(err)}
    }
};

let cda_outros_tipos = async (page) => {
    let acumulador
    let inscricao = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id*="inscricao"] tr td:nth-child(1) div[class="ui-dt-c"]')).map((el)=>{return el.innerText}));
    (await inscricao.length == 1 && (inscricao[0].substring(0,3) === 'Não')) ? acumulador = '' : acumulador = await inscricao.join("; ");
    return acumulador
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
        //await page._client.send('Page.setDownloadBehavior', {
        await page._client().send('Page.setDownloadBehavior', {
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
        {header: 'Classe Judicial', key: 'classe', width: 35},
        {header: 'Procurador prevento', key: 'procurador', width: 40},
        {header: 'CDA / DEBCAD / NDFG / FNDE', key: 'cda_debcad', width: 80},
        {header: 'Controle_presc', key: 'controle_presc', width: 15},
        {header: 'Valor Atualizado', key: 'valor', width: 17},
        {header: 'CPF/CNPJ polo passivo', key: 'polo', width: 25},
        {header: 'Juízo', key: 'juizo', width: 20},
        //{header: 'última autuação', key: 'auto', width: 20},
        {header: 'Rating Grupo', key: 'rating_grupo', width: 20},
        {header: 'Demanda Analytics', key: 'demanda_analytics', width: 20},
        {header: 'Processos Vinculados', key: 'vinculados', width: 20},
        {header: 'Finalizado', key: 'finalizado', width: 10}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsulta.length; i++){
        if (coluna == 'Provisória') {
            worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento, cda_debcad: listconsulta[i].cda, controle_presc: listconsulta[i].prescricao, valor: listconsulta[i].valor, polo: listconsulta[i].parte, juizo: listconsulta[i].juizo, rating_grupo: listconsulta[i].grupo, demanda_analytics: listconsulta[i].demanda, vinculados: listconsulta[i].vinculados, finalizado: listconsulta[i].pesquisa}); //loop para escrever o nome dos procuradores e processos na planilha exportada
        } else {
            worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento, cda_debcad: listconsulta[i].cda, controle_presc: listconsulta[i].prescricao, valor: listconsulta[i].valor, polo: listconsulta[i].parte, juizo: listconsulta[i].juizo, rating_grupo: listconsulta[i].grupo, demanda_analytics: listconsulta[i].demanda, vinculados: listconsulta[i].vinculados, finalizado: ''}); //loop para escrever o nome dos procuradores e processos na planilha exportada
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
            listconsulta.push({numero: sh.getRow(i).getCell(1).value, classe: sh.getRow(i).getCell(2).value, prevento: sh.getRow(i).getCell(3).value, cda: sh.getRow(i).getCell(4).value, prescricao: sh.getRow(i).getCell(5).value, valor: sh.getRow(i).getCell(6).value, parte: sh.getRow(i).getCell(7).value, juizo: sh.getRow(i).getCell(8).value, registro: sh.getRow(i).getCell(9).value, grupo: sh.getRow(i).getCell(10).value, demanda: sh.getRow(i).getCell(11).value, vinculados: sh.getRow(i).getCell(12).value, pesquisa: sh.getRow(i).getCell(13).value});
        }
    }
};

let scrape = async () => {
    await ler_output(nome_arquivo_excel);
    await console.log('Lendo arquivo excel ' + "\n");
    await ler_input();
    await console.log('Total de ' + (listprocessos.length) + ' Processos lidos no Input' + "\n"); 
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        headless: false, args:['--start-maximized',  '--disable-features=site-per-process'], //torna visível e maximiza a 
        ignoreHTTPSErrors: true,
    });
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf', {waitUntil: 'networkidle0'}); //Ambiente de produção
    await dados_inicio(page);
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
    await writeexcel(nome_arquivo_excel, 'Provisória');
    if (await array_registro.length > 0) {
        await console.log('Pesquisa dos registros de autuação');
        for (let i = 0; i < await array_registro.length; i++) {
            let id_el = await listconsulta.findIndex(x => x.numero === array_registro[i].numero); 
            //let dados_autuacao = await pesquisa_processo (page, array_registro[i].numero, array_registro[i].classe);
            await pesquisa_processo (page, array_registro[i].numero, array_registro[i].classe);           
            if (await (['Execução Fiscal Previdenciária', 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)', 'Execução Fiscal (FNDE)'].includes(array_registro[i].classe))){
                await page.waitForSelector('tbody[id*="inscricao"] tr td:nth-child(1) div[class="ui-dt-c"]', {visible: true}, { timeout: 120000 });
                //await page.waitForSelector('tbody[id*="inscricao"] tr div[class*="ui-dt-c"]', { timeout: 120000 });
                listconsulta[id_el].cda = await cda_outros_tipos(page);
                await console.log(`linha ${(id_el+1)} / ${listconsulta.length}: ${array_registro[i].numero}`);
            }
            await writeexcel(nome_arquivo_excel, 'Provisória');
        }    
    }
    await writeexcel(nome_arquivo_excel, 'Final');
    let result = await 'Concluído ';
    browser.close();
    return result
};  

scrape().then((value) => {
   console.log(value)
});