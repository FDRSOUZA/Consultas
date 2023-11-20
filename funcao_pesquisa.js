let pesquisa_processo = async (page, numero_pj, classe) => {
    let autuacao;
    await page.goto('https://saj.pgfn.fazenda.gov.br/saj/pages/pesquisarProcessoJudicial.jsf?'); 
    await page.waitForSelector('input[name$="numProcesso"]'); //Aguarda carregar a página da "Consulta"
    await page.focus('input[name$="numProcesso"]'); //caixa de texto do número do pj recebe o foco
    let numero_pj_semformatacao = numero_pj.replace(/[.,-]/g, ''); //retira o "." e "-" do nº do processo
    await page.$eval('input[name$="numProcesso"]', (el, value) => el.value = value, numero_pj_semformatacao); //inclui o numero do processo na caixa de texto
    await page.keyboard.press('Enter');  //pressiona enter
    const navigationPromise = await page.waitForNavigation(); // aguarda a navegação
    let cadastrado = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(2)', anchors => { return anchors.map(anchor => anchor.textContent)}); //lista, em um array, as classes do processo inserido na caixa de texto
    await navigationPromise; // final da navegação
    let dados_processo = [];
    if (await cadastrado.length > 0) { 
        let painel_caixas_disponiveis = await page.evaluate(()=> Array.from(document.querySelectorAll('tbody[id*="dtTable_data"] tr td div[class="ui-dt-c"] a')).map(i=>{return i.id})); // lê os ids dos processos cadastrados - link para abrir o processo escolhido
        let numeros = await page.$$eval('tbody[id*="dtTable_data"] tr td:nth-child(1)', anchors => { return anchors.map(anchor => anchor.textContent)});
        if (cadastrado.some(el => instancia_1.includes(el)) || cadastrado.some(el => tipo_execucao.includes(el))) { //Se tiver uma classe de execução ou de 1º instância
            let id_classe = await cadastrado.findIndex(element => element === classe); // perquisa a classe que vc quer 
            if (await id_classe !== -1) { // se a classe for igual a que vc quer
                dados_processo.push({"num": numeros[id_classe], "classe": cadastrado[id_classe], "link":`tbody[id*="dtTable_data"] tr td div a[id="${painel_caixas_disponiveis[id_classe]}"]`, "indice": id_classe, "quant": cadastrado.length}); // guarde para clicar
            } else if (await id_classe == -1 && cadastrado.length == 1) { // se não guarde a primeira ocorrência
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