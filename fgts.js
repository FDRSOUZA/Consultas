"use strict";
const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var path = require('path');
const fs = require('fs');
let listprocessos = [];
let listconsulta = [];
let d1;
let d2;
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `OUTPUT - Execução - ${dia}-${mes}-${ano}.xlsx`;

function formataDinheiro(n) {
    return n.toFixed(2).replace('.', ',');
    //return "R$ " + n.toFixed(2).replace('.', ',').replace(/(\d)(?=(\d{3})+\,)/g, "$1.");
}

let separa_cda = async (conteudo) => {
   let texto = await conteudo.replace(/\s/g, '');
   let arr_debcads = await texto.split(";");
   return arr_debcads
};

   //########## LER A PLANILHA COM OS PROCESSO A SEREM PESQUISADOS ############### 
let ler_dados = async () => {
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, 'Dados.xlsx')
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('FGTS'); // Primeira aba do arquivo excel - Planilha
        d1 = sh.getRow(1).getCell(1).text;
        d2 = sh.getRow(1).getCell(2).text;
    }
};

   //########## LER A PLANILHA COM OS PROCESSO A SEREM PESQUISADOS ############### 
let ler_excel = async (arq) => {
    var wb = await new Excel.Workbook();
    var filePath = await path.resolve(__dirname, arq);
    if (fs.existsSync(filePath)) { 
        await wb.xlsx.readFile(filePath); 
        let sh = await wb.getWorksheet('Auxiliar'); // Primeira aba do arquivo excel - Planilha
        await sh.eachRow(function(cell, rowNumber) {
            let coluna_cda;
            if (sh.getRow(rowNumber).getCell(2).text === 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)' && sh.getRow(rowNumber).getCell(4).text !== '') {
                coluna_cda = sh.getRow(rowNumber).getCell(4).text.toUpperCase();
            } else { 
                coluna_cda = sh.getRow(rowNumber).getCell(4).text;
            }
            listprocessos.push({numero: sh.getRow(rowNumber).getCell(1).text, classe: sh.getRow(rowNumber).getCell(2).text, prevento: sh.getRow(rowNumber).getCell(3).text, cda: coluna_cda, prescricao: sh.getRow(rowNumber).getCell(5).text, valor: sh.getRow(rowNumber).getCell(6).text, parte: sh.getRow(rowNumber).getCell(7).text, juizo: sh.getRow(rowNumber).getCell(8).text, registro: sh.getRow(rowNumber).getCell(9).text, grupo: sh.getRow(rowNumber).getCell(10).text, demanda: sh.getRow(rowNumber).getCell(11).text, vinculados: sh.getRow(rowNumber).getCell(12).text, linha: rowNumber});
        });
    }
    await listprocessos.shift();
};

let writeexcel = async (arq) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Auxiliar");
    worksheet.columns = [
        {header: 'Processo', key: 'processo', width: 25},
        {header: 'Classe Judicial', key: 'classe', width: 25},
        {header: 'Procurador prevento', key: 'procurador', width: 30},
        {header: 'CDA / DEBCAD / NDFG / FNDE', key: 'cda_debcad', width: 60},
        {header: 'Controle_presc', key: 'controle_presc', width: 11},
        {header: 'Valor Atualizado', key: 'valor', width: 17},
        {header: 'CPF/CNPJ polo passivo', key: 'polo', width: 25},
        {header: 'Juízo', key: 'juizo', width: 20},
        {header: 'última Autuação', key: 'autuacao', width: 20},
        {header: 'Rating Grupo', key: 'rating_grupo', width: 20},
        {header: 'Demanda Analytics', key: 'demanda_analytics', width: 20},
        {header: 'Processos Vinculados', key: 'vinculados', width: 30}
    ];
    for (let lin = 1; lin < 10; lin++) {worksheet.getColumn(lin).font = {name: 'Times New Roman', size: 10};}
    worksheet.getColumn(6).alignment = {horizontal: 'left'}
    worksheet.getRow(1).font = {name: 'Calibri', size: 11, bold: true} // Coloque o cabeçalho em negrito
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
   //worksheet.getColumn(6).numFmt = ' "R$" #. ## 0,00; [Red] \ - "R$" #. ## 0,00 ' ;
    for(let i = 0; i < listprocessos.length; i++){
       worksheet.addRow({processo: listprocessos[i].numero, classe: listprocessos[i].classe, procurador: listprocessos[i].prevento, cda_debcad: listprocessos[i].cda, controle_presc: listprocessos[i].prescricao, valor: listprocessos[i].valor, polo: listprocessos[i].parte, juizo: listprocessos[i].juizo, autuacao: listprocessos[i].registro, rating_grupo: listprocessos[i].grupo, demanda_analytics: listprocessos[i].demanda, vinculados: listprocessos[i].vinculados}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    }  
    await wbook.xlsx.writeFile(arq);
}

let scrape = async () => {
    ler_dados();
    await console.log('Lendo arquivos excel ' + "\n");
    await ler_excel(nome_arquivo_excel);
    let array_pesquisa = await listprocessos.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
        return (obj.classe === 'Execução Fiscal (FGTS e Contr. Sociais da LC 110)' && obj.cda !== '' && (obj.cda.substring(0,2).toUpperCase() === 'FG' || obj.cda.substring(0,2).toUpperCase() === 'CS') && (obj.parte !== '' && obj.parte.length === 18) && obj.valor === '');
    });
    await console.log('Total de processo da Planilha Input ' + listprocessos.length);
    await console.log('Lidos ' + (array_pesquisa.length) + ' Processos para pesquisa' + "\n");
    const browser = await puppeteer.launch({ //cria uma instância do navegador
        executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', 
        args:['--start-maximized'],//maximiza
        headless: false, //torna visível 
        ignoreHTTPSErrors: true
    });
    const page = await browser.newPage();
    page.setDefaultTimeout(600*1000);
    await page.setViewport({width:0, height:0});
    var pages = await browser.pages();
    await pages[0].close();
    await new Promise(r => setTimeout(r, 1000));
    await page.goto('https://portalfge.caixa.gov.br/PortalPgfn/Portal/Login/Login.asp', {waitUntil: 'networkidle0'});
    const frame = await page.frames().find(frame => frame.name() === 'PortalPFN');
    await page.focus('td input[name="txtPis"]');
    await page.$eval('input[name="txtPis"]', (el, value) => el.value = value, d1);
    await page.$eval('input[name="txtSenha"]', (el, value) => el.value = value, d2);
    await page.click('button img[src="https://portalfge.caixa.gov.br/PortalPgfn/Portal/_images/botao_confirmar.gif');
    await page.waitForSelector('button img[src="https://portalfge.caixa.gov.br/PortalPgfn/Portal/_images/botao_sair.gif', {visible: true});
    await page.goto('https://portalfge.caixa.gov.br/PortalPgfn/FGE/060/061/648/fgeps648.asp', {waitUntil: 'networkidle0'});
    await page.waitForSelector('button img[src="https://portalfge.caixa.gov.br/PortalPgfn/FGE/Images/botsaldo.gif"]', {visible: true});
    let contador = await 1;
    for (let i = 0; i < await array_pesquisa.length; i++) {
        let id_alt = await listprocessos.findIndex(element => element.numero === array_pesquisa[i].numero);    
        let array_fgts = await separa_cda(array_pesquisa[i].cda);
        //await console.log('--------- Linha '+ array_pesquisa[i].linha + '  -  Processo: '+array_pesquisa[i].numero);
        let cnpj = await array_pesquisa[i].parte.replace(/[^0-9]/g,"");
        let acumulador_cda_fase_valor; 
        let valor_temp = 0;
        await page.$eval('input[name="Num_Inscricao_Emp_input"]', (el, value) => el.value = value, cnpj);
        await new Promise(r => setTimeout(r, 50));
        await page.click('button[name="saldo"] img[src="https://portalfge.caixa.gov.br/PortalPgfn/FGE/Images/botsaldo.gif"]');
        await new Promise(r => setTimeout(r, 50));
        let end = page.url().substring(page.url().length-9, page.url().length);
        if (await end === 'aux=Saldo') {
            await page.waitForSelector('table[class="fieldset"] tbody tr td[class="txtcentral"]', {visible: true});
            await new Promise(r => setTimeout(r, 50));
            let dados = await page.evaluate(() =>Array.from(document.querySelectorAll('table[class="fieldset"] tbody tr td')).map((el)=>{return el.innerText}));
            for (let r = 0; r < await array_fgts.length; r++) {
                let id = await dados.findIndex(el => el == array_fgts[r]);
                if (await id !== -1) {
                    let valor;
                    let situacao = await dados[id+2];
                    await dados[id+1] !== '' ? valor = await dados[id+1].replace(/\./g,"").replace(/\,/g,".") : valor = await 0;
                    if (await r === 0) { 
                        acumulador_cda_fase_valor = await `${array_fgts[r]}; ${situacao}; R$ ${valor}`;
                        valor_temp = await parseFloat(valor);  
                    } else {
                        acumulador_cda_fase_valor = await `${acumulador_cda_fase_valor};\n${array_fgts[r]}; ${situacao}; R$ ${valor}`;
                        valor_temp = await (parseFloat(valor_temp) + parseFloat(valor));
                    } 
                }
            }
            if (valor_temp !== '' && valor_temp !== '.') {
                let total_formatado = await formataDinheiro(valor_temp);
                //let total_formatado = await valor_temp.toFixed(2);s
                valor_temp = await total_formatado;
            }     
        } else {
            await console.log('PÁGINA COM BLOQUEIO');
            acumulador_cda_fase_valor = await array_pesquisa[i].cda;
            valor_temp = await '.' ;
        } 
        listprocessos[id_alt].cda = await acumulador_cda_fase_valor;
        listprocessos[id_alt].valor = await valor_temp;
        await new Promise(r => setTimeout(r, 50));
        let falta = await array_pesquisa.length-contador;
        await contador++
        if (falta !== 0) {
            await console.log(`Linha ${array_pesquisa[i].linha}: ${array_pesquisa[i].numero} - Restam: ${falta} consulta(s)`);
            //await  console.log('QUANTIDADE DE CONSULTAS RESTANTES: '+falta+'\n');
        } else { await  console.log('CONSULTAS TERMINADAS'+'\n'); }
        await page.goto('https://portalfge.caixa.gov.br/PortalPgfn/FGE/060/061/648/fgeps648.asp', {waitUntil: 'networkidle0'});
        await page.waitForSelector('button img[src="https://portalfge.caixa.gov.br/PortalPgfn/FGE/Images/botsaldo.gif"]', {visible: true});
        await writeexcel(nome_arquivo_excel);
    } 
    await new Promise(r => setTimeout(r, 500));
    browser.close()
    let result = `Processo pesquisados:${listprocessos.length}`;
    return result
}  

scrape().then((value) => {
   console.log(value)
});