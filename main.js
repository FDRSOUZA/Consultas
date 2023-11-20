//const functions = require('./functions');
//const functions_saj = require('./functions_saj');
//const saj_selectors = require('./saj_selectors');
let Excel = require('exceljs');
const fs = require('fs');
var path = require('path');
const puppeteer = require('puppeteer');
const header_padrao = ['PJE', 'Classe', 'Responsável', 'Tipo de atuação', 'Tipo de Petição', 'Descrição'];
let descricao_tibunal = [];
let prop_obj = ['numero', 'classe', 'pfn', 'atuacao', 'peticao', 'descricao'];
let distribuir = [];
let registrar = [];

let tribunal = [{estado: 'Alagoas', secao: 'Seção Judiciária Alagoas e Subseções', processos: []},
        {estado: 'Ceará', secao: 'Seção Judiciária Ceará e Subseções', processos: []},
        {estado: 'Paraíba', secao: 'Seção Judiciária Paraíba e Subseções', processos: []},
        {estado: 'Pernambuco', secao: 'Seção Judiciária Pernambuco e Subseções', processos: []},
        {estado: 'Rio Grande do Norte', secao: 'Seção Judiciária Rio Grande do Norte e Subseções', processos: []},
        {estado: 'Sergipe', secao: 'Seção Judiciária Sergipe e Subseções', processos: []},
        {estado: 'TRF5', secao: 'Tribunal Regional Federal da 5ª Região', processos: []}]
const unidade = 'PRFN5 (Sede)'


let ler_excel = async (nome_arquivo) => {
    let sheets = [];
    var wb = new Excel.Workbook();
    var filePath = path.resolve(__dirname, nome_arquivo);
    await wb.xlsx.readFile(filePath);
    await wb.eachSheet(function (worksheet) {
        sheets.push(worksheet.name); //Coloca o nome das abas da planilha em um array
    });
    let indice = 0;
    await sheets.map(item => {
        let worksheet = wb.getWorksheet(item);
        worksheet.eachRow({ includeEmpty: false }, async (row) => {//ler cada uma das linhas da planilha
            await tribunal[indice].processos.push(row.values[1]); //Salva cada linha em um array
        })
        indice++; 
    });
    /*await tribunal.forEach(element => {
        console.log(element.secao);
        console.log(element.processos);
        console.log('___________________________')
      });*/
    
    /*let sh = await wb.getWorksheet(aba);
    let planilha_lida = new Array();
    let planilha_final = new Array();

    let noelia_header_line = await functions.header_line('Brainiac 1.5.xlsx', 'NOÉLIA', 3, header_padrao);//Array que determina quais colunas da planilha serão usadas
    //let triadores_header_line = await functions.header_col('Brainiac 1.5.xlsx', 'BDT', 1, 'PFNs');
    let col_pfns_header = await functions.header_col('Brainiac 1.5.xlsx', 'NOÉLIA', 3, 'Responsável'); // Retorna um array do nome que está nas linha da coluna Responsável
    let tabela = await functions.read_file('Brainiac 1.5.xlsx', 'NOÉLIA', 3, noelia_header_line, prop_obj); // Retorna um array de objetos com os dados da planilha
    
    const pfnSemRepeticao = await [...new Set(col_pfns_header)]; // Exclui do Array as repetições
    await pfnSemRepeticao.map(item => {// Percorre o array dos nomes dos pfns
        let dados_proc = [];
        tabela.map(function(obj) {// Percorre os registros da planilha que foram salvos no objeto
            if (obj.pfn === item) { // Se o nome da vez do ARRAY DE OBJETOS for IGUAL ao nome da vez do ARRAY , coloca os dados em um objeto e salva
                let temp = {numero: obj.numero, classe: obj.classe};
                dados_proc.push(temp);
            }
        }); 
        distribuir.push({pfn: item, dados: dados_proc});

          // ## Agrupa os processos que vão ser feito registro da atuação
        let filtrados = tabela.filter(function(obj) {//Filtra por pfn
            return (item === obj.pfn);
        });
        filtrados.map(function(obj) { // Agrupa
            let id = registrar.findIndex( arr => arr.pfn === obj.pfn && arr.atuacao === obj.atuacao && arr.peticao === obj.peticao && arr.descricao === obj.descricao && arr.arquivo === obj.arquivo && arr.conteudo === obj.conteudo);
            let processo = {numero: obj.numero, classe: obj.classe};
            if (id == -1) {
                registrar.push({pfn: obj.pfn, atuacao: obj.atuacao, peticao: obj.peticao, descricao: obj.descricao, dados:[processo], status: 'Pendente'});    
            } else {
                registrar[id].dados.push(processo);
            }
        });   

    }); */ 
    let listfinal = await tribunal.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
        return (obj.processos.length > 0);
    }); 
    return listfinal;
};

let formar_lote_processo = async function (array_processos) {
    let lista = await array_processos.map(el => el.numero).join(',');
    const lista_acabada = await lista.replace(/,+/g, "\n");
    return lista_acabada;
}

let integracao = async function (page, array, secao) {
    await page.goto('https://saj.pgfn.fazenda.gov.br/integracao/pages/solicitacao/cadastrar.jsf', {waitUntil: 'networkidle0'}),//AMBIENTE DE PRODUÇÃO
    await page.click('input[id="frmConsulta:chkConsultaDados_input"]');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "block"');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "none"');
    //await new Promise(r => setTimeout(r, 200)); 
    await page.click('input[id="frmConsulta:chkConsultaInteiroTeor_input"]');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "block"');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "none"');
    await page.click('label[id="frmConsulta:listaRegiao_label"]');
    await page.waitForSelector('div[id="frmConsulta:listaRegiao_panel"] ul li', {visible: true});
    await new Promise(r => setTimeout(r, 400));
    await page.click('div[id="frmConsulta:listaRegiao_panel"] ul li:nth-child(6)');
    await page.waitForXPath(`//*[@class='ui-selectonemenu-label ui-inputfield ui-corner-all' and contains(., "5ª Região")]`);
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "block"');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "none"');
    await page.click('label[id="frmConsulta:listaSistemas_label"]');
    await page.waitForSelector('div[id="frmConsulta:listaSistemas_panel"] ul li', {visible: true});
    await new Promise(r => setTimeout(r, 400));
    await page.click('div[id="frmConsulta:listaSistemas_panel"] ul li:nth-child(2)');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "block"');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "none"');
    await new Promise(r => setTimeout(r, 400));
    await page.click('#frmConsulta\\:listaOrgaoJustica > button');
    //await page.click('input[id="frmConsulta:listaOrgaoJustica_input"]');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "block"');
    await page.waitForFunction('document.getElementById("j_idt109:statusAguarde").style.display === "none"');
    await page.waitForSelector('span[id="frmConsulta:listaOrgaoJustica_panel"] ul li', {visible: true});
    if (await descricao_tibunal.length === 0){
        descricao_tibunal = await page.evaluate(() => Array.from(document.querySelectorAll("#frmConsulta\\:listaOrgaoJustica_panel > ul > li")).map((el)=>{return el.innerText}));
    }
    let id = await descricao_tibunal.findIndex( desc => desc === secao)+1;
    await new Promise(r => setTimeout(r, 400));
    await page.click(`span[id="frmConsulta:listaOrgaoJustica_panel"] ul li:nth-child(${id})`);
    await page.evaluate((array) => {document.querySelector('textarea[id="frmConsulta:txtAreaNumerosCNJ"]').value = array}, array);
    await new Promise(r => setTimeout(r, 3000));
}

let scrape = async () => {
    let listintegracao = await ler_excel('Processos Integração.xlsx');
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        //executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
        executablePath: "C:/Program Files/Google/Chrome/Application/chrome.exe",
        headless: false, args: ['--start-maximized'],//torna visível 
        ignoreHTTPSErrors: true,
    });
    
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf', {waitUntil: 'networkidle0'});
    await page.waitForNavigation(); 
    await page.waitForSelector('input[name$="formMenus"]');
    
    //await page.waitForNavigation();
    //const page = await functions.abrir_browser();
    //await page.goto('https://saj.pgfn.fazenda.gov.br/saj/login.jsf', {waitUntil: 'networkidle0'}); //AMBIENTE DE PRODUÇÃO
    //await page.waitForNavigation();
    /*for (let m = 0; m < await distribuir.length; m++) {
        let lotes = await formar_lote_processo(distribuir[m].dados);
        await console.log(lotes);
        await console.log(distribuir[m].pfn);
        await console.log('____________________________');
        await functions_saj.texto_unidade(page, saj_selectors.sel_label_unid_distrib, saj_selectors.sel_lista_unid_distrib, saj_selectors.sel_click_unid_distrib, unidade);
        await functions_saj.incluir_lote_processo(page, lotes, '2ª Instância', saj_selectors.sel_triang_incluir_distrib, saj_selectors.sel_link_incluir_distrib, saj_selectors.sel_text_area)
        
    }*/
    for (let i = 0; i < listintegracao.length; i++) {
        lotes = await listintegracao[i].processos.map(el => el).join(',');
        const lista_acabada = await lotes.replace(/,+/g, "\n");
        await integracao(page, lista_acabada, listintegracao[i].secao)
    }

    //lotes = await lista.map(el => el).join(',');
    //const lista_acabada = await lotes.replace(/,+/g, "\n");
    //await functions_saj.integracao(page, lista_acabada);
}  

scrape().then(() => {
   //console.log()
});