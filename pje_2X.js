// MACRO QUE COLETA AS INTIMAÇÕES NO PJE

const puppeteer = require('puppeteer');
let Excel = require('exceljs');
var listdistribuicao = new Array ();
var listtriagem = new Array ();
var listconsolidado = new Array ();
var datetime = new Date();
var ano = datetime.getFullYear();
var mes = datetime.getMonth()+1; 
var dia = datetime.getDate();
const nome_arquivo_excel = `./Extração PJE - ${dia}-${mes}-${ano}.xlsx`;
let acesso = [
    {login: "https://pje1g.trf5.jus.br/pje/login.seam", quadro_aviso: "https://pje1g.trf5.jus.br/pje/QuadroAviso/", painel: "https://pje1g.trf5.jus.br/pje/Painel/painel_usuario/advogado.seam", origem: "PJE-", trib: "JEF", quant: 0},
    //{login: "https://pje2g.trf5.jus.br/pje/login.seam", quadro_aviso: "https://pje2g.trf5.jus.br/pje/QuadroAviso/", painel: "https://pje2g.trf5.jus.br/pje/Painel/painel_usuario/advogado.seam", origem: "-TR-PJE", trib: "TR", quant: 0},
    {login: "https://pjett.trf5.jus.br/pje/login.seam", quadro_aviso: "https://pjett.trf5.jus.br/pje/QuadroAviso/", painel: "https://pjett.trf5.jus.br/pje/Painel/painel_usuario/advogado.seam", origem: "TRU-TRF5", trib: "TT", quant: 0},
    {login: "https://pje1g.trf1.jus.br/pje/login.seam", quadro_aviso: "https://pje1g.trf1.jus.br/pje/QuadroAviso/", painel: "https://pje1g.trf1.jus.br/pje/Painel/painel_usuario/advogado.seam", origem: "JFBA-Juazeiro-PJE", trib: "TRF1", quant: 0}  
]

const converg = [{turma: 'Turma Recursal da Paraíba', secao: 'PJE-TR-PB'},
                {turma: 'Turma Recursal de Pernambuco', secao: 'PJE-TR-PE'},
                {turma: 'Turma Recursal do Ceará', secao: 'PJE-TR-CE'},
                {turma: 'Turma Recursal do Rio Grande do Norte', secao: 'PJE-TR-RN'},
                {turma: 'Turma Recursal de Sergipe', secao: 'PJE-TR-SE'},
                {turma: 'Turma Recursal de Alagoas', secao: 'PJE-TR-AL'},
            ]

let quantidades = [{secao: "PJE-AL", quant: 0}, {secao: "PJE-BA-Juazeiro", quant: 0}, {secao: "PJE-CE", quant: 0}, {secao: "PJE-PB", quant: 0},  {secao: "PJE-PE", quant: 0}, {secao: "PJE-RN", quant: 0}, {secao: "PJE-SE", quant: 0},
                   {secao: 'PJE-TR-AL', quant: 0}, {secao: 'PJE-TR-CE', quant: 0}, {secao: "PJE-TR-PB", quant: 0}, {secao: 'PJE-TR-PE', quant: 0}, {secao: 'PJE-TR-RN', quant: 0}, {secao: 'PJE-TR-SE', quant: 0}
                  
            ]

let writeexcel = async (arq) => { //funcao para criar o excel de exportacao 
    const wbook = new Excel.Workbook();
    const worksheet = wbook.addWorksheet("NUJEF");
    worksheet.columns = [
        {header: 'Processo', key: 'processo', width: 25},
        {header: 'Classe', key: 'classe', width: 13},
        {header: 'Origem', key: 'orgao', width: 15},
        {header: 'Órgão Julgador', key: 'juris', width: 20},
        {header: 'Data', key: 'data', width: 11},
        {header: 'Polo ativo', key: 'parte1', width: 33},
        {header: 'Polo passivo', key: 'parte2', width: 33},
        {header: 'Expediente', key: 'documento', width: 25},
        {header: 'Prazo', key: 'prazo', width: 8},
        {header: 'Assunto', key: 'materia', width: 40}
    ];
    worksheet.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listdistribuicao.length; i++){
        worksheet.addRow({processo: listdistribuicao[i].processo, classe: listdistribuicao[i].classe, orgao: listdistribuicao[i].orgao, juris: listdistribuicao[i].juris, data: listdistribuicao[i].data, parte1: listdistribuicao[i].parte1, parte2: listdistribuicao[i].parte2, documento: listdistribuicao[i].documento, prazo: listdistribuicao[i].prazo, materia: listdistribuicao[i].materia, origem: listdistribuicao[i].origem}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    } 
    
    const worksheet_1 = wbook.addWorksheet("JF");
    worksheet_1.columns = [
        {header: 'Processo', key: 'processo', width: 25},
        {header: 'Classe', key: 'classe', width: 13},
        {header: 'Origem', key: 'orgao', width: 15},
        {header: 'Órgão Julgador', key: 'juris', width: 20},
        {header: 'Data', key: 'data', width: 11},
        {header: 'Polo ativo', key: 'parte1', width: 33},
        {header: 'Polo passivo', key: 'parte2', width: 33},
        {header: 'Expediente', key: 'documento', width: 25},
        {header: 'Prazo', key: 'prazo', width: 8},
        {header: 'Assunto', key: 'materia', width: 40}
    ];
    worksheet_1.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listtriagem.length; i++){
        worksheet_1.addRow({processo: listtriagem[i].processo, classe: listtriagem[i].classe, orgao: listtriagem[i].orgao, juris: listtriagem[i].juris, data: listtriagem[i].data, parte1: listtriagem[i].parte1, parte2: listtriagem[i].parte2, documento: listtriagem[i].documento, prazo: listtriagem[i].prazo, materia: listtriagem[i].materia, origem: listtriagem[i].origem}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    } 

    const worksheet_2 = wbook.addWorksheet("Quant Sintético");
    worksheet_2.columns = [
        {header: 'Tribunal', key: 'trib', width: 25},
        {header: 'Quantidade', key: 'quant', width: 15}
    ];
    worksheet_2.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < acesso.length; i++){
        worksheet_2.addRow({trib: acesso[i].trib, quant: parseInt(acesso[i].quant)}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    } 

    const worksheet_3 = wbook.addWorksheet("Consolidado");
    worksheet_3.columns = [
        {header: 'Tribunal', key: 'trib', width: 25},
        {header: 'Quantidade', key: 'quant', width: 15}
    ];
    worksheet_3.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < quantidades.length; i++){
        worksheet_3.addRow({trib: quantidades[i].secao, quant: quantidades[i].quant}); //loop para escrever o nome dos procuradores e processos na planilha exportada
    }

    const worksheet_4 = wbook.addWorksheet("Detalhado");
    worksheet_4.columns = [
        {header: 'Seção', key: 'trib', width: 25},
        {header: 'Quantidade', key: 'quant', width: 15}
    ];
    worksheet_4.getRow(1).font = {bold: true} // Coloque o cabeçalho em negrito
    for(let i = 0; i < listconsolidado.length; i++){
        worksheet_4.addRow({trib: listconsolidado[i].secao, quant: parseInt(listconsolidado[i].quant)}); 
    }

    await wbook.xlsx.writeFile(arq);
}

let trf1 = async (page, quant_pagina, orig) => { //funcao para coletar dados da Subs
    let pagina = 1;
    let marca;
    let contador_real = 0;
    do {
        let coletor = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="formExpedientes:tbExpedientes:tb"] tr')).map(i=>{return i.innerText}));
        for (let limpar=0; limpar < coletor.length; limpar++){
            let col = await coletor[limpar].split("\n").filter(function (i) {//o novo array recebe os dados sem linhas em branco e o filtro é para tirar os elementos com valor vazio
                return i;
            });
            let juizo = await col[8].substring(1,col[8].length);
            if (await juizo === 'Juizado Especial Cível e Criminal Adjunto à Vara Federal da SSJ de Juazeiro-BA') {
                await contador_real++;
                let n_proc = await col[6].substring(col[6].indexOf("-")-7, col[6].indexOf("-")+18);
                let ato = await col[2];
                let data_expediente = await col[3].substring(col[3].indexOf("(")+1, col[3].indexOf("(")+11);
                let prazo_expediente;
                (col[4] == ' sem prazo' || col[4] == 'sem prazo' || col[4] == ' ') ? prazo_expediente = '0' : prazo_expediente = await col[4].replace(/[^0-9]/g,'');
                //let prazo_expediente = await col[4].substring(col[4].indexOf(":")+1, col[4].length);
                let divide_proc_classe_assunto = await col[6].split(' '+n_proc+' ');
                let divide_partes = await col[7].split(' X ');
                await listdistribuicao.push({"processo": n_proc, "classe": divide_proc_classe_assunto[0], "orgao": 'PJE-BA-Juazeiro', "juris": juizo, "data": data_expediente, "parte1": divide_partes[0], "parte2": divide_partes[1], "prazo": prazo_expediente, "documento": ato, "materia": divide_proc_classe_assunto[1], "origem": orig})
            }    
        }
        await pagina++;
        if (await page.$('.rich-datascr-button') != null) {
            await pagina == 2 ? marca = 5 : marca = marca + 1;
            let seletor_proxima_pagina = await page.$(`table[id="formExpedientes:tbExpedientes:scPendentes_table"] tbody tr td:nth-child(${marca})`);
            //let seletor_proxima_pagina = await page.$('table[id="formExpedientes:tbExpedientes:scPendentes_table"] tbody tr td[class="rich-datascr-inact"]');//Seletor do botão que avança para a próxima página
            await seletor_proxima_pagina.click();
            await page.waitForFunction('document.getElementById("_viewRoot:status.start").style.display === "none"');
        }
    } while (pagina <= quant_pagina) ;
    acesso[3].quant = await contador_real;
    quantidades[1].quant = await contador_real;
    await writeexcel(nome_arquivo_excel);
}

(async function abrir_painel_procurado(){
    const browser = await puppeteer.launch({headless: false, args: ['--start-maximized']});
    const page = await browser.newPage();
    page.setDefaultTimeout(300*1000);
    var pages = await browser.pages();
    await pages[0].close();
    console.log('Aguarde, processando a requisição...')
    // acessa o PJE, faz login e vai até o painel do procurador
    await page.setViewport({ width: 0, height: 0});
    for (a= 0; a < acesso.length; a++) {
        await page.goto(acesso[a].login, {waitUntil: 'networkidle0'});
        //await new Promise(r => setTimeout(r, 3000));
        //await page.waitForSelector('div[class="conteudo-login"] iframe', {visible: true});    
        //var frames = (await page.frames());
        //await console.log(frames);
               
        //await page.click('#loginAplicacaoButton');
        //await page.click('#kc-pje-office');

        await page.waitForSelector('#barraSuperiorPrincipal > div > div.navbar-collapse > ul > li > a > span.hidden-xs.nome-sobrenome.tip-bottom');
        await page.click('#barraSuperiorPrincipal > div > div.navbar-collapse > ul > li > a > span.hidden-xs.nome-sobrenome.tip-bottom'); //clica na barra para ver qtos perfis o usuário possui
        if (await page.$('select[id="papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao"] option') !== null) { //Se existir mais de um perfil do usuário
            let opcoes_usuario = await page.evaluate(() => Array.from(document.querySelectorAll('select[id="papeisUsuarioForm:usuarioLocalizacaoDecoration:usuarioLocalizacao"] option')).map((el)=>{return el.innerText}));//Lista todos os perfis do usuário em um array
            let indice_usuario;
            if (await opcoes_usuario.length > 1) { //Existe mais de um perfil
                let u = await 0;
                do {
                    let inicio_elemento = await opcoes_usuario[u].substring(0,12); //Procura o perfil da Procuradoria
                    if (await inicio_elemento == 'Procuradoria') {
                        indice_usuario = await u;
                        u = await opcoes_usuario.length
                    } 
                    await u++;
                } while (await u < opcoes_usuario.length);
                //let indice_usuario = await opcoes_usuario.findIndex(element => element === 'Procuradoria - Procuradoria da Fazenda Nacional / Representante processual');
                if (await indice_usuario !== -1) {
                    await page.select('#papeisUsuarioForm\\:usuarioLocalizacaoDecoration\\:usuarioLocalizacao', indice_usuario.toString());
                }
            }
        } else {await page.click('#barraSuperiorPrincipal > div > div.navbar-collapse > ul > li > a > span.hidden-xs.nome-sobrenome.tip-bottom');}
    
        await new Promise(resolve => setTimeout(resolve, 3*1000));
        await browser.waitForTarget(target => target.url().substring(0, 42) == acesso[a].quadro_aviso);
        console.log("Acessou tela quadro de avisos "+ acesso[a].trib)
        await page.goto(acesso[a].painel);
        //await new Promise(resolve => setTimeout(resolve, 2*1000));
        let tarefa = await page.evaluate(()=> Array.from(document.querySelectorAll('.nomeTarefa')).map((el)=>{return el.innerText})); //lista as tarefas do painel
        let link = await page.evaluate(()=> Array.from(document.querySelectorAll('div[id*="formAbaExpediente:listaAgrSitExp"] a')).map((el)=>{return el.id}));//lista as ids das tarefas do painel
        let quant_tarefa = await page.evaluate(()=> Array.from(document.querySelectorAll('div span[class="pull-right mr-10"]')).map((el)=>{return el.innerText}));//lista a quantidade de processos em cada tarefa do painel
        let indice_tarefa = await tarefa.findIndex(element => element === 'Apenas pendentes de ciência'); //Pesquisa o índice da tarefa desejada - "Apenas pendentes de ciência"
        //let indice_tarefa = await tarefa.findIndex(element => element === acesso[a].servico); //Pesquisa o índice da tarefa desejada - "Apenas pendentes de ciência"
        if (await quant_tarefa[indice_tarefa] !== '0') { //Se a quantidade de processo for diferente de zero? entra no bloco
            let quant_consolidado_inst = await page.$eval(`a[id="${link[indice_tarefa]}"] span[class= "pull-right mr-10"]`, el => el.innerText);
            acesso[a].quant = await quant_consolidado_inst;
            await page.click(`a[id="${link[indice_tarefa]}"]`);
            await page.waitForSelector('td[class="rich-tree-node-text treeNodeItem"]>a>span[class="nomeTarefa"]');
            await console.log('Acessou painel e abriu os: Apenas pendentes de ciência');
            let id_childs;
            await console.log("_________________________________"+"\n");
            if (acesso[a].trib !== "TRF1") {
                let linhas_contexto_tarefa = await page.evaluate((indice_tarefa)=> Array.from(document.querySelectorAll(`div[id="formAbaExpediente:listaAgrSitExp:${indice_tarefa}:trPend:childs"] div`)).map((el)=>{return el.id}), indice_tarefa);//lista as ids das linhas no contexto das tarefas do painel
                id_childs = await linhas_contexto_tarefa.filter(function(el) {//Seleciona em um array os dados do processos de 2ª instância
                    return (el !== '');
                });
            }

            // coleta o NOME e ID de todas as caixas disponíveis no painel do PJE e armazena cada qual numa matriz
            var painel_caixas_disponiveis = await page.evaluate(()=>{
                var allcaixas_span = Array.from(document.querySelectorAll('td[class="rich-tree-node-text treeNodeItem"]>a>span[class="nomeTarefa"]'))
                var allcaixas_a = Array.from(document.querySelectorAll('td[class="rich-tree-node-text treeNodeItem"]>a'))
                var allcaixas_span_array = allcaixas_span.map(i=>{return i.innerText}) // retorna o nome de todas as jurisdições que constam no painel
                var allcaixas_a_array = allcaixas_a.map(i=>{return i.id}) // retorna o id de todas as jurisdições que constam no painel
                return [allcaixas_span_array, allcaixas_a_array]
            }) // encerra localizar NOME e ID das caixas disponíveis 
            
            let painel_caixas_disponiveis_nome = painel_caixas_disponiveis[0];
            let painel_caixas_disponiveis_id = painel_caixas_disponiveis[1];
            let entrada = 0;
            if (await acesso[a].trib == 'TRF1') {
                let indice_Juazeiro = await painel_caixas_disponiveis_nome.findIndex(element => element === 'Subseção Judiciária de Juazeiro-BA');
                let seletor_id_juazeiro = await painel_caixas_disponiveis_id[indice_Juazeiro];
                painel_caixas_disponiveis_nome = await []; 
                painel_caixas_disponiveis_id = await [];
                await painel_caixas_disponiveis_nome.push('Subseção Judiciária de Juazeiro-BA');
                await painel_caixas_disponiveis_id.push(seletor_id_juazeiro);
                let linhas_contexto_tarefa = await page.evaluate((indice_Juazeiro)=> Array.from(document.querySelectorAll(`div[id="formAbaExpediente:listaAgrSitExp:${indice_Juazeiro}:trPend:childs"] div`)).map((el)=>{return el.id}), indice_Juazeiro);//lista as ids das linhas no contexto das tarefas do painel
                id_childs = await linhas_contexto_tarefa.filter(function(el) {//Seleciona em um array os dados do processos de 2ª instância
                    return (el !== '');
                });
                entrada = await indice_Juazeiro + 1;
            }
    
            let acumulador;
            for(var i_caixa_painel=0; i_caixa_painel<painel_caixas_disponiveis_id.length; i_caixa_painel++){
                let marca; //Variável criada para clicar no número da página, caso a caixa do painel tenha + de 1 página
                await page.click(`a[id="${painel_caixas_disponiveis_id[i_caixa_painel]}"]`); //Clica na caixa que será coletado os dados dos processos
                await page.waitForFunction('document.getElementById("_viewRoot:status.start").style.display === "none"');
                let instancia = await `${painel_caixas_disponiveis_nome[i_caixa_painel]}  Caixa de entrada`;// Nome da caixa na tag h6
                await page.waitForXPath(`//*[@class='col-xs-8 col-md-8' and contains(., '${instancia}')]`); // aguarda alterar o nome da caixa na tag h6
                let temp = await id_childs[entrada];
                await page.click(`div[id="${id_childs[entrada]}"] table tbody tr`); 
                await page.waitForFunction('document.getElementById("_viewRoot:status.start").style.display === "none"');
                let res = await page.evaluate((temp) => (document.querySelector(`div[id="${temp}"] table tbody tr td div[class="col-md-1 itemContador"]`).innerText), temp);
                await page.waitForSelector(`a[id="${painel_caixas_disponiveis_id[i_caixa_painel]}"] span[class="pull-right mr-10"]`)// aguarda carregar a página inicial da caixa desejada - com o total de processos
                let quant_proc_caixa = await res;
                await console.log(`Processos disponíveis para ${painel_caixas_disponiveis_nome[i_caixa_painel]}: ${quant_proc_caixa}`);
                await listconsolidado.push({secao : painel_caixas_disponiveis_nome[i_caixa_painel], quant: quant_proc_caixa});
                if (await acesso[a].trib === "JEF") {//Se estiver acessando o 1º grau - acumula e salva os valores das quantiades de processos no array quantidades
                    let indice = await quantidades.findIndex(element => element.secao === `PJE-${painel_caixas_disponiveis_nome[i_caixa_painel].substring(0,2)}`);
                    quantidades[indice].quant = await (parseInt(quantidades[indice].quant) + parseInt(quant_proc_caixa));
                } else if (await acesso[a].trib === "TR") {
                    let indice = await converg.findIndex(element => element.turma === painel_caixas_disponiveis_nome[i_caixa_painel]);
                    let ind = await quantidades.findIndex(element => element.secao === converg[indice].secao);
                    quantidades[ind].quant = await parseInt(quant_proc_caixa);
                } 
                acumulador == null ? acumulador = await quant_proc_caixa : acumulador = parseInt(acumulador) +  parseInt(await quant_proc_caixa); 
                var total_paginas = await Math.ceil(quant_proc_caixa/40);
                await console.log('Total de páginas: '+total_paginas);
                await console.log('Acumulado: '+acumulador);
                await console.log('%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% '+"\n");
                entrada = await entrada + 2;
                let pagina = 1;
                let orgao_julgador;
                let jurisdicao;
                q = 0
                if (await acesso[a].trib == 'TRF1') { 
                    await trf1(page, total_paginas, acesso[a].origem);
                } else {
                
                    do {
                        let numero_processos = await page.evaluate(() => Array.from(document.querySelectorAll('tbody[id="formExpedientes:tbExpedientes:tb"] tr')).length);//Verifica qtos processos existem na página
                        for (let m = 1; m <= numero_processos; m++) {
                            if (await acesso[a].trib == 'JEF') {
                                orgao_julgador = await acesso[a].origem + await painel_caixas_disponiveis_nome[i_caixa_painel].slice(0,2);
                                let juris = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div[class="col-md-8 informacoes-linha-expedientes"] div[class="col-md-12"] div:nth-child(3)`, el => el.textContent);
                                jurisdicao = await juris.slice(1, juris.length); 
                            } else if (await acesso[a].trib == 'TR') {
                                let ind = await converg.findIndex( c => c.turma === painel_caixas_disponiveis_nome[i_caixa_painel] );
                                //orgao_julgador =  converg[ind].secao + acesso[a].origem;
                                orgao_julgador =  converg[ind].secao;
                                jurisdicao = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div[class="col-md-8 informacoes-linha-expedientes"] div[class="col-md-12"] div:nth-child(3)`, el => el.textContent);   
                            } else if (await acesso[a].trib == 'TT') {
                                orgao_julgador = await 'PJE-TRU-PRFN5'; 
                                jurisdicao = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div[class="col-md-8 informacoes-linha-expedientes"] div[class="col-md-12"] div:nth-child(3)`, el => el.textContent);   
                            }
                            let proc_link = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div a[class*="numero-processo-expediente"]`, el => el.textContent);
                            let proc_final = await proc_link.slice(proc_link.length-25, proc_link.length)
                            let classe_acao = await proc_link.slice(0, proc_link.length-26);
                            let materia_texto = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div[class="col-md-8 informacoes-linha-expedientes"] div[class="col-md-12"] div`, el => el.textContent);
                            let materia = await materia_texto.slice(proc_link.length+1, materia_texto.length); 
                            let tipo_documento = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div span[title="Tipo de documento"]`, el => el.textContent);
                            let partes = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div[class="col-md-8 informacoes-linha-expedientes"] div[class="col-md-12"] div:nth-child(2)`, el => el.textContent);
                            let parte1 = await partes.slice(0,partes.indexOf(" X "));
                            let parte2 = await partes.slice(partes.indexOf(" X ")+3, partes.length);
                            let meio_comunicacao = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div span[title="Meio de comunicação"] span[title="Data de criação do expediente"]`, el => el.textContent);
                            let data = await meio_comunicacao.slice(0, 10);
                            let prazo;
                            //let prazo_texto = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div[title="Prazo para manifestação"]`, el => el.textContent.replace(/[^0-9]/g,''));
                            let prazo_texto = await page.$eval(`tbody[id="formExpedientes:tbExpedientes:tb"] tr:nth-child(${m}) td:nth-child(2) div div div:nth-child(4)`, el => el.textContent.replace(/[^0-9]/g,'')); 
                            (await prazo_texto == '' || await prazo_texto == ' ') ? prazo = '0' : prazo = await prazo_texto; 
                            //let prazo = await prazo_texto.replace(/[^0-9]/g,'');
                            //let prazo = await prazo_texto.slice(prazo_texto.indexOf(":")+1, prazo_texto.length);
                            if (await orgao_julgador == 'PJE-PE' && classe_acao !== 'PJEC' && classe_acao !== 'CumSen' && classe_acao !== 'CumSenFaz') {await listtriagem.push({"processo": proc_final, "classe": classe_acao, "orgao": orgao_julgador, "juris": jurisdicao, "data": data, "parte1": parte1, "parte2": parte2, "prazo": prazo, "documento": tipo_documento, "materia": materia, "origem": acesso[a].origem})}
                            else {await listdistribuicao.push({"processo": proc_final, "classe": classe_acao, "orgao": orgao_julgador, "juris": jurisdicao, "data": data, "parte1": parte1, "parte2": parte2, "prazo": prazo, "documento": tipo_documento, "materia": materia, "origem": acesso[a].origem})};
                        }
                        await pagina++;
                        if (await page.$('.rich-datascr-button') != null) {
                            await pagina == 2 ? marca = 5 : marca = marca + 1;
                            let seletor_proxima_pagina = await page.$(`table[id="formExpedientes:tbExpedientes:scPendentes_table"] tbody tr td:nth-child(${marca})`);//Seletor para clicar no número da página
                            //let seletor_proxima_pagina = await page.$('table[id="formExpedientes:tbExpedientes:scPendentes_table"] tbody tr td[class="rich-datascr-inact"]');//Seletor do botão que avança para a próxima página
                            await seletor_proxima_pagina.click();
                            await page.waitForFunction('document.getElementById("_viewRoot:status.start").style.display === "none"');
                        }
                        //await pagina++;
                    } while (pagina <= total_paginas)
                    await writeexcel(nome_arquivo_excel);
                    //console.log(`Total de processos extraídos: ${acumulador}`) 
                }  
            }
        }
    }
    await browser.close();

    await console.log('Procedimento finalizado com sucesso!')

})(); // finaliza a função principal


