"use strict";
const puppeteer = require ('puppeteer')
let Excel = require('exceljs');
var path = require('path');
const fs = require('fs');
let listprocessos = [];
let listconsulta = [];
const datetime = new Date();
const ano = datetime.getFullYear();
const mes = datetime.getMonth()+1;
const dia = datetime.getDate();
const nome_arquivo_excel = `OUTPUT - Execução - ${dia}-${mes}-${ano}.xlsx`;

function formataDinheiro(n) {
   return n.toFixed(2).replace('.', ',');
   //return "R$ " + n.toFixed(2).replace('.', ',').replace(/(\d)(?=(\d{3})+\,)/g, "$1.");
}

      //################# FUNÇÃO PARA DAR CLICK EM LINK - PROCESSO OU MATÉRIA
let click_link = async (page, caminho, tempo) => {
   try {
      await page.waitForSelector(caminho, { timeout: tempo }); //Aguarda o seletor
      await page.click(caminho); // clicar - abrir processo encontrado ou ver matéria ou classe
   } 
   catch(e) { console.error('Não clicou no link - Por favor, click manualmente no link de matéria')}
};

let prepara_para_float = async (conversor) => {
   conversor = await conversor.trim();
   let moeda = await conversor.replace(/\./g, '');
   let moeda_convertida = await moeda.replace(',', '.');
   return moeda_convertida;   
}; 

let ler_resultado_consulta = async (page, seletor) => {
   let consulta = await page.evaluate((seletor) => { //Ler a matéria
      let quant = document.querySelector(seletor).value; // Ler se tem matéria
      return quant
   }, seletor);
   return consulta;
};

let separa_cda = async (conteudo) => {
   let texto = await conteudo.replace(/\s/g, '');
   let arr_debcads = await texto.split(";");
   return arr_debcads
};

let consulta_debcad = async (page, arr_insc) => {
   let acumulador_cda_fase_valor = await '';
   let dados_gerais = await [];
   let pesquisa;
   let acumula_valor = await 0.00;
   let valor_final;
   for (let x = 0; x < await arr_insc.length; x++) {
      let descricao = await arr_insc[x].replace(/\-/g, '');
      if (descricao.length === 9) {
         await page.waitForSelector('input[id="SDF_NU_0009_ID_PROCESSO"]');
         await page.focus('input[id="SDF_NU_0009_ID_PROCESSO"]');
         await page.type('input[id="SDF_NU_0009_ID_PROCESSO"]', descricao);
         await page.keyboard.press('Enter');
         const navigationPromise = await page.waitForNavigation();
         await navigationPromise; 
         await new Promise(r => setTimeout(r, 400));
         await page.waitForSelector('input[id="SDF_AN_20_VL_HONOR_ATUAL"]');
         let desc_fase = await ler_resultado_consulta(page, 'input[id="SDF_AN_0040_TE_FASE"]');
         pesquisa = await ler_resultado_consulta(page, 'input[id="SDF_AN_0020_VL_TOTAL"]');//Valor Total = Principal + Multa Isolada + Multa Ofício + Multa Mora + Juros + Encargo
         let valor = await pesquisa.trim(); // Tira o espaços em branco da variável pesquisa
         let valor_debito = await prepara_para_float(valor); //Transforma em valor flutuante     
         pesquisa = await ler_resultado_consulta(page, 'input[id="SDF_AN_20_VL_HONOR_ATUAL"]'); // Ler o valor dos honorários
         let honorarios = await pesquisa.trim(); //Tira espaços da variável
         let valor_honorarios = await prepara_para_float(honorarios); // Trnasforma em flutuante
         if (await valor_debito == '') {valor_debito = '0'}
         if (await valor_honorarios == '') {valor_honorarios = '0'}
         let total = await (parseFloat(valor_debito) + await parseFloat(valor_honorarios)); //Soma o principal + honorários
         acumula_valor = await (parseFloat(acumula_valor) + await parseFloat(total)); //Acumula
         let total_formatado = await formataDinheiro(total);
         await x === await 0 ? acumulador_cda_fase_valor = `${arr_insc[x]}; ${desc_fase}; ${total_formatado}` : acumulador_cda_fase_valor =  `${acumulador_cda_fase_valor};\n${arr_insc[x]}; ${desc_fase}; ${total_formatado}`
      } else if (descricao.length !== 9) {
         await x === await 0 ? acumulador_cda_fase_valor = `${arr_insc[x]}; ${' -'}; ${' -'}` : acumulador_cda_fase_valor =  `${acumulador_cda_fase_valor};\n${arr_insc[x]}; ${' -'}; ${' -'}`
      }
   }
   if (valor_final !== '' & valor_final !== ' - ') {
      valor_final = await formataDinheiro(acumula_valor);
   }
   await dados_gerais.push(acumulador_cda_fase_valor); 
   await dados_gerais.push(valor_final);
   //await console.log('Total: '+formataDinheiro(acumula_valor));
   return dados_gerais
};

   //########## LER A PLANILHA COM OS PROCESSO A SEREM PESQUISADOS ############### 
let ler_excel = async (lista, arq) => {
   var wb = await new Excel.Workbook();
   var filePath = await path.resolve(__dirname, arq);
   if (fs.existsSync(filePath)) { 
      await wb.xlsx.readFile(filePath); 
      let sh = await wb.getWorksheet('Auxiliar'); // Primeira aba do arquivo excel - Planilha
      await sh.eachRow(function(cell, rowNumber) {
         lista.push({numero: sh.getRow(rowNumber).getCell(1).text, classe: sh.getRow(rowNumber).getCell(2).text, prevento: sh.getRow(rowNumber).getCell(3).text, cda: sh.getRow(rowNumber).getCell(4).text, prescricao: sh.getRow(rowNumber).getCell(5).text, valor: sh.getRow(rowNumber).getCell(6).text, parte: sh.getRow(rowNumber).getCell(7).text, juizo: sh.getRow(rowNumber).getCell(8).text, registro: sh.getRow(rowNumber).getCell(9).text, grupo: sh.getRow(rowNumber).getCell(10).text, demanda: sh.getRow(rowNumber).getCell(11).text, vinculados: sh.getRow(rowNumber).getCell(12).text, linha: rowNumber});
      });
      lista.shift();
   }
};

let writeexcel = async (arq) => { //funcao para criar o excel de exportacao
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("Auxiliar");
    worksheet.columns = [
        {header: 'Processo', key: 'processo', name: 'Times New Roman', size: 9, width: 25},
        {header: 'Classe Judicial', key: 'classe', name: 'Times New Roman', size: 9, width: 25},
        {header: 'Procurador prevento', key: 'procurador', name: 'Times New Roman', size: 9, width: 30},
        {header: 'CDA / DEBCAD / NDFG / FNDE', key: 'cda_debcad', name: 'Times New Roman', size: 9, width: 60},
        {header: 'Controle_presc', key: 'controle_presc', name: 'Times New Roman', size: 9, width: 11},
        {header: 'Valor Atualizado', key: 'valor', name: 'Times New Roman', size: 9, width: 17},
        {header: 'CPF/CNPJ polo passivo', key: 'polo', name: 'Times New Roman', size: 9,  width: 25},
        {header: 'Juízo', key: 'juizo', name: 'Times New Roman', size: 9, width: 20},
        {header: 'última Autuação', key: 'autuacao', name: 'Times New Roman', size: 9, width: 20},
        {header: 'Rating Grupo', key: 'rating_grupo', width: 20},
        {header: 'Demanda Analytics', key: 'demanda_analytics', width: 20},
        {header: 'Processos Vinculados', key: 'vinculados', width: 30}
    ];
    for (let lin = 1; lin < 10; lin++) {worksheet.getColumn(lin).font = {name: 'Times New Roman', size: 10};}
    worksheet.getColumn(6).alignment = {horizontal: 'left'}
    worksheet.getRow(1).font = {name: 'Calibri', size: 11, bold: true} // Coloque o cabeçalho em negrito
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsulta.length; i++){
        worksheet.addRow({processo: listconsulta[i].numero, classe: listconsulta[i].classe, procurador: listconsulta[i].prevento, cda_debcad: listconsulta[i].cda, controle_presc: listconsulta[i].prescricao, valor: listconsulta[i].valor, polo: listconsulta[i].parte, juizo: listconsulta[i].juizo, autuacao: listconsulta[i].registro, rating_grupo: listconsulta[i].grupo, demanda_analytics: listconsulta[i].demanda, vinculados: listconsulta[i].vinculados}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    }  
    await wbook.xlsx.writeFile(arq);
}

let scrape = async () => {
   await console.log('Lendo arquivos excel ' + "\n");
   await ler_excel(listprocessos, nome_arquivo_excel);
   let array_pesquisa = await listprocessos.filter(function(obj) {//Seleciona em um array os dados do processos de 2ª instância
      return ((obj.classe === 'Execução Fiscal Previdenciária' || obj.classe === 'Execução Fiscal (FNDE)') && obj.cda !== '' && obj.valor === '');
   });
   await console.log('Total de processo da Planilha Input ' + listprocessos.length);
   await console.log('Lidos ' + (array_pesquisa.length) + ' Processos para pesquisa' + "\n");
   const browser = await puppeteer.launch({ //cria uma instância do navegador
      //executablePath:'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe', 
      args:['--start-maximized'],
      headless: false, //torna visível 
      ignoreHTTPSErrors: true
   });
   const page = await browser.newPage();
   page.setDefaultTimeout(600*1000);
   await page.setViewport({width:0, height:0});
   var pages = await browser.pages();
   await pages[0].close();
   await new Promise(r => setTimeout(r, 1000));
   await page.goto('https://imsaarfbsp1.sec.prevnet/cassso/login?service=http%3A%2F%2Fw3b9.sec.prevnet%2FMV2%2F', { timeout: 100000 });
   await page.waitForSelector('body > div.container > div.span-12.last > form > fieldset > div > p:nth-child(2) > input[type="image"]');
   click_link(page, 'body > div.container > div.span-12.last > form > fieldset > div > p:nth-child(2) > input[type="image"]'); 
   await page.waitForSelector('table > tbody > tr > td > table > tbody > tr:nth-child(7) > td > table > tbody > tr:nth-child(5) > td:nth-child(2) > div > a');
   await page.goto('http://w3b9.sec.prevnet/divida/Gerenciador');
   await page.waitForSelector('input[name="SDF_NAVG"]');     
   await page.focus('input[name="SDF_NAVG"]');
   await page.keyboard.press('Delete');  
   await page.type('input[name="SDF_NAVG"]', 'CCRED');
   await page.keyboard.press('Enter'); 
   await page.waitForSelector('input[name="SDF_NU_0009_ID_PROCESSO"]');
   await page.focus('input[name="SDF_NU_0009_ID_PROCESSO"]'); 
   await new Promise(r => setTimeout(r, 2000));
   let contador = await 0;
   for (let i = 0; i < await array_pesquisa.length; i++) {
      let id_alt = await listprocessos.findIndex(element => element.numero === array_pesquisa[i].numero);
      let cda = '';
      let array_debcad = await separa_cda(array_pesquisa[i].cda);
      cda = await consulta_debcad(page, array_debcad);
      listprocessos[id_alt].cda = await cda[0];
      await listprocessos[id_alt].classe === 'Execução Fiscal (FNDE)' && cda[0].substring(cda[0].length,cda[0].length-1) == '-' ? listprocessos[id_alt].valor = '-' : listprocessos[id_alt].valor = await cda[1];
      //await listprocessos[id_alt].classe === 'Execução Fiscal (FNDE)' && cda[0].substring(cda[0].length,cda[0].length-1) == '-' ? listprocessos[id_alt].valor = '-' : listprocessos[id_alt].valor = await cda[1].toFixed(2);
      await new Promise(r => setTimeout(r, 200));   
      listconsulta = await listprocessos;
      await contador++;
      let falta = await array_pesquisa.length-contador;
      if (falta !== 0) {
         await console.log(`Linha ${array_pesquisa[i].linha}: ${array_pesquisa[i].numero} - Restam: ${falta} consulta(s)`);
         //await  console.log('QUANTIDADE DE CONSULTAS RESTANTES: '+falta+'\n');
      } else { await  console.log('CONSULTAS TERMINADAS'+'\n'); }
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