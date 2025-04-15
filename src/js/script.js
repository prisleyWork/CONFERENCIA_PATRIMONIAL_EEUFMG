// Declaração de variáveis para armazenar os workbooks
var workbooks = {
    ativos: null,
    baixados: null,
    internos: null
};
let resultadosAvancado = [];
let fragmentoAtual = '';
let indiceAvancadoAtual = 0;

// Função para carregar as planilhas
function carregarPlanilhas() {
    Promise.all([
        fetch('ATIVOS.xlsx').then(response => {
            if (!response.ok) throw new Error("Erro ao carregar ATIVOS.xlsx");
            return response.arrayBuffer();
        }).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.ativos = XLSX.read(arrayData, { type: 'array' });
        }).catch(error => console.error("Erro no carregamento da planilha de ativos:", error)),

        fetch('BAIXADOS.xlsx').then(response => {
            if (!response.ok) throw new Error("Erro ao carregar BAIXADOS.xlsx");
            return response.arrayBuffer();
        }).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.baixados = XLSX.read(arrayData, { type: 'array' });
        }).catch(error => console.error("Erro no carregamento da planilha de baixados:", error)),

        fetch('INTERNOS.xlsx').then(response => {
            if (!response.ok) throw new Error("Erro ao carregar INTERNOS.xlsx");
            return response.arrayBuffer();
        }).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.internos = XLSX.read(arrayData, { type: 'array' });
        }).catch(error => console.error("Erro no carregamento da planilha de internos:", error))
    ]).then(() => {
        console.log("Todas as planilhas foram carregadas.");
    }).catch(error => console.error("Erro ao carregar as planilhas:", error));
}

// Executa o carregamento das planilhas ao carregar a página
carregarPlanilhas();

// Atualiza o ano automaticamente no rodapé
const anoAtualSpan = document.getElementById("anoAtual");
if (anoAtualSpan) {
    const dataAtual = new Date();
    anoAtualSpan.textContent = dataAtual.getFullYear();
} else {
    console.warn("Elemento #anoAtual não encontrado no documento.");
}

// Função para formatar a entrada do usuário
function formatInput(input) {
    if (!input) return '';
    var noLeadingZeros = input.replace(/^0+/, ''); // Remove zeros à esquerda
    var formattedInput = noLeadingZeros.replace(/-\d{1,2}$/, ''); // Remove sufixos específicos
    return formattedInput.trim();
}

// Função para normalizar código ATM
function normalizarCodigoATM(codigo) {
    if (!codigo) return ''; // Retorna string vazia se o código for null ou indefinido
    return codigo.replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase(); // Remove espaços e caracteres especiais
}

// Função para traduzir a condição do bem
function translateCondition(condition) {
    var translations = {
        'BM': 'Bom',
        'AE': 'Anti-Econômico',
        'IR': 'Irrecuperável',
        'OC': 'Ocioso',
        'BX': 'Baixado',
        'RE': 'Recuperável'
    };
    return translations[condition] || condition; // Retorna a tradução ou o valor original
}

// Função para traduzir a situação do bem
function translateSituation(situation) {
    var translations = {
        'NI': 'Não encontrado no local da guarda',
        'NO': 'Normal'
    };
    return translations[situation] || situation; // Retorna a tradução ou o valor original
}

// Função para calcular a similaridade entre dois códigos
function calcularSimilaridade(codigo1, codigo2) {
    if (!codigo1 || !codigo2) return 0; // Retorna 0 se algum código for inválido
    let iguais = 0;
    let comprimento = Math.min(codigo1.length, codigo2.length);

    for (let i = 0; i < comprimento; i++) {
        if (codigo1[i] === codigo2[i]) iguais++;
    }

    return (iguais / comprimento) * 100; // Retorna a porcentagem de similaridade
}

// Função para buscar informações básicas
function buscar() {
    document.getElementById('resultado').innerHTML = ''; // limpa resultados anteriores
    var inputField = document.getElementById('numero');
    var formattedInput = formatInput(inputField.value);
    var resultado = document.getElementById('resultado');
    var listaHistorico = document.getElementById('lista-historico');

    if (!formattedInput) {
        resultado.innerHTML = '<p>Por favor, insira um número válido para buscar.</p>';
        return;
    }

    // Função interna para buscar em uma planilha específica
    function buscarNaPlanilha(workbook, estrutura) {
        if (!workbook || !workbook.Sheets) return null;

        var worksheet = workbook.Sheets[workbook.SheetNames[0]];
        var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        for (var i = 0; i < jsonData.length; i++) {
            if (estrutura === "ATIVOS_INTERNOS") {
                if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
                    var translatedCondition = translateCondition(jsonData[i][3]);
                    var translatedSituation = translateSituation(jsonData[i][5]);
                    var atmInfo = jsonData[i][2];
                    var atmDisplay = atmInfo ? `<br><b>ATM:</b> ${atmInfo}` : '';

                    return `
                        <b>Número de patrimônio:</b> ${jsonData[i][0]}-${jsonData[i][1]}<br>
                        <b>Tipo:</b> ${jsonData[i][25]}<br>
                        <b>Descrição:</b> ${jsonData[i][8]}<br>
                        <b>Situação:</b> ${translatedSituation}<br>
                        <b>Condição do Bem:</b> ${translatedCondition}<br>
                        <b>Local da Guarda:</b> ${jsonData[i][17]}<br>
                        <b>Responsável:</b> ${jsonData[i][27]}${atmDisplay}
                    `;
                }
            } else if (estrutura === "BAIXADOS") {
                if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
                    return `
                        <b>Número de patrimônio:</b> ${jsonData[i][0]}-${jsonData[i][1]}<br>
                        <b>Número ATM:</b> ${jsonData[i][3]}<br>
                        <b>Setor:</b> ${jsonData[i][10]}<br>
                        <b>Descrição:</b> ${jsonData[i][2]}<br>
                        <b>Último Local da Guarda:</b> ${jsonData[i][13]}<br>
                        <b>Observação:</b> <b>Bens baixados devem ser mantidos no local de guarda atual. Caso deseje desfazer do bem, cadastre no sistema de desfazimento.</b>
                    `;
                }
            }
        }
        return null; // Retorna null se não encontrou nada
    }

    // Consultar nas três planilhas
    var resultadoTexto = buscarNaPlanilha(workbooks.ativos, "ATIVOS_INTERNOS") ||
        buscarNaPlanilha(workbooks.baixados, "BAIXADOS") ||
        buscarNaPlanilha(workbooks.internos, "ATIVOS_INTERNOS") ||
        "Número não encontrado no banco de dados.";

    resultado.innerHTML = resultadoTexto;

    // Adicionar resultado ao histórico
    if (resultadoTexto !== "Número não encontrado no banco de dados.") {
        var novoItem = document.createElement('li');
        novoItem.innerHTML = resultadoTexto;
        listaHistorico.appendChild(novoItem);
    }

    inputField.value = ''; // Limpa o campo de entrada
    inputField.focus();
}

// Parte 2

// Função para busca avançada com base no fragmento
function buscarATMAvancado() {
    const resultadoDiv = document.getElementById('resultado');
    resultadoDiv.innerHTML = '';
    const inputField = document.getElementById('numero');
    const modoBusca = document.getElementById('modo-busca').value;
    let fragmento = inputField.value.trim();
    const resultados = [];


    if (!fragmento) {
        resultadoDiv.innerHTML = '<p>Por favor, insira pelo menos um fragmento para busca avançada.</p>';
        return;
    }

    if (modoBusca === 'descricao') {
        setTimeout(() => buscarPorDescricao(), 500);
        return;
    }

    if (modoBusca === 'patrimonio') {
        fragmento = formatInput(fragmento);
        if (!/^[0-9]+$/.test(fragmento)) {
            resultadoDiv.innerHTML = '<p>Por favor, insira apenas números para busca por patrimônio.</p>';
            return;
        }
    }

    if (modoBusca === 'atm') {
        fragmento = fragmento.replace(/[^0-9]/g, '');
        if (!/^[0-9]{1,12}$/.test(fragmento)) {
            resultadoDiv.innerHTML = '<p>Fragmento inválido para ATM. Use apenas números.</p>';
            return;
        }
    }

        setTimeout(() => {
        resultadosAvancado = [];
        fragmentoAtual = fragmento;
        indiceAvancadoAtual = 0;

        ['ativos', 'baixados', 'internos'].forEach(estrutura => {
            const workbook = workbooks[estrutura];
            if (!workbook) return;

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            jsonData.forEach(item => {
                const codigoBanco = modoBusca === 'atm' ? String(item[2]).replace(/[^0-9]/g, '') : String(item[0]);
                if (!codigoBanco) return;

                if (codigoBanco.includes(fragmento)) {
                    let resultadoTratado;
                    if (estrutura === 'baixados') {
                        resultadoTratado = `
                            <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                            <b>Número ATM:</b> ${item[3]}<br>
                            <b>Setor:</b> ${item[10]}<br>
                            <b>Descrição:</b> ${item[2]}<br>
                            <b>Último Local da Guarda:</b> ${item[13]}<br>
                            <b>Observação:</b> <b>Bens baixados devem ser mantidos no local de guarda atual. Caso deseje desfazer do bem, cadastre no sistema de desfazimento.</b>
                        `;
                    } else {
                        let translatedCondition = translateCondition(item[3]);
                        let translatedSituation = translateSituation(item[5]);
                        let atmInfo = item[2];
                        let atmDisplay = atmInfo ? `<br><b>ATM:</b> ${atmInfo}` : '';

                        resultadoTratado = `
                            <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                            <b>Tipo:</b> ${item[25]}<br>
                            <b>Descrição:</b> ${item[8]}<br>
                            <b>Situação:</b> ${translatedSituation}<br>
                            <b>Condição do Bem:</b> ${translatedCondition}<br>
                            <b>Local da Guarda:</b> ${item[17]}<br>
                            <b>Responsável:</b> ${item[27]}${atmDisplay}
                        `;
                    }

                    resultadosAvancado.push({ resultadoTratado });
                }
            });
        });

        mostrarLoteAvancado(); // nova função!
    }, 500);
}

// Atualizar exibirResultadosTratados()
function mostrarLoteAvancado() {
    const resultadoDiv = document.getElementById('resultado');

    if (indiceAvancadoAtual === 0) {
        resultadoDiv.innerHTML = `<h3>🔎 Resultados para: <strong>${fragmentoAtual}</strong></h3>`;
    }

    const lote = resultadosAvancado.slice(indiceAvancadoAtual, indiceAvancadoAtual + 10);
    lote.forEach(resultado => {
        resultadoDiv.innerHTML += `
            <p>
                ${resultado.resultadoTratado}<br>
                <button onclick="adicionarAoHistorico(\`${resultado.resultadoTratado}\`)">Adicionar ao Histórico</button>
            </p>
            <hr>
        `;
    });

    indiceAvancadoAtual += 10;

    if (indiceAvancadoAtual < resultadosAvancado.length) {
        resultadoDiv.innerHTML += `
            <div style="text-align:center; margin: 1rem;">
                <button onclick="mostrarLoteAvancado()">Carregar mais</button>
            </div>
        `;
    } else if (indiceAvancadoAtual === 0) {
        resultadoDiv.innerHTML = `<p>🔎 Nenhuma correspondência encontrada para "${fragmentoAtual}".</p>`;
    }
}


// Atualizar exibirResultadosDescricao()
function exibirResultadosDescricao(resultados, palavraChave) {
    const resultadoDiv = document.getElementById('resultado');
    if (resultados.length > 0) {
        resultadoDiv.innerHTML = `<h3>🔎 Resultados para: <strong>${palavraChave}</strong></h3>`;
        resultados.forEach(resultado => {
            resultadoDiv.innerHTML += `
                <p>
                    ${resultado.resultadoTratado}<br>
                    <button onclick="adicionarAoHistorico(\`${resultado.resultadoTratado}\`)">Adicionar ao Histórico</button>
                </p>
                <hr>
            `;
        });
    } else {
        resultadoDiv.innerHTML = `<p>🔎 Nenhuma correspondência encontrada para "${palavraChave}".</p>`;
    }
}

// Atualizar buscarPorDescricao para busca exata
function buscarPorDescricao() {
    const resultadoDiv = document.getElementById('resultado');
    resultadoDiv.innerHTML = '';
    const inputField = document.getElementById('numero');
    const palavraChave = removerAcentos(inputField.value.trim().toLowerCase());
    const resultados = [];

    if (!palavraChave) {
        resultadoDiv.innerHTML = '<p>Por favor, insira ao menos uma palavra-chave para busca por descrição.</p>';
        return;
    }

    setTimeout(() => {
        ['ativos', 'baixados', 'internos'].forEach(estrutura => {
            const workbook = workbooks[estrutura];
            if (!workbook) return;

            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            jsonData.forEach(item => {
                // Campos para busca: descrição e códigos que contenham texto
                const descricaoCampo = estrutura === 'ativos' ? item[8] : item[2];
                const codigoCampo = String(item[0]) + String(item[1]); // exemplo: 2024nE000359

                const descricao = descricaoCampo ? removerAcentos(String(descricaoCampo).toLowerCase()) : '';
                const codigo = removerAcentos(codigoCampo.toLowerCase());

                if (descricao.includes(palavraChave) || codigo.includes(palavraChave)) {
                    let resultadoTratado;

                    if (estrutura === 'baixados') {
                        resultadoTratado = `
                            <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                            <b>Número ATM:</b> ${item[3]}<br>
                            <b>Setor:</b> ${item[10]}<br>
                            <b>Descrição:</b> ${item[2]}<br>
                            <b>Último Local da Guarda:</b> ${item[13]}<br>
                            <b>Observação:</b> <b>Bens baixados devem ser mantidos no local de guarda atual. Caso deseje desfazer do bem, cadastre no sistema de desfazimento.</b>
                        `;
                    } else {
                        const translatedCondition = translateCondition(item[3]);
                        const translatedSituation = translateSituation(item[5]);
                        const atmInfo = item[2];
                        const atmDisplay = atmInfo ? `<br><b>ATM:</b> ${atmInfo}` : '';

                        resultadoTratado = `
                            <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                            <b>Tipo:</b> ${item[25]}<br>
                            <b>Descrição:</b> ${item[8]}<br>
                            <b>Situação:</b> ${translatedSituation}<br>
                            <b>Condição do Bem:</b> ${translatedCondition}<br>
                            <b>Local da Guarda:</b> ${item[17]}<br>
                            <b>Responsável:</b> ${item[27]}${atmDisplay}
                        `;
                    }

                    resultados.push({ resultadoTratado });
                    if (resultados.length >= 10) return;
                }
            });
        });

        exibirResultadosDescricao(resultados, palavraChave);
    }, 500);
}

function removerAcentos(texto) {
    return texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

// Função para exibir resultados tratados (usado para busca avançada e por descrição)
function exibirResultadosTratados(resultados, fragmento) {
    var resultadoDiv = document.getElementById('resultado');
    if (resultados.length > 0) {
        resultadoDiv.innerHTML = '<h3>Resultados Encontrados:</h3>';
        resultados.forEach(resultado => {
            resultadoDiv.innerHTML += `
                <p>
                    ${resultado.resultadoTratado}<br>
                    <button onclick="adicionarAoHistorico(\`${resultado.resultadoTratado}\`)">Adicionar ao Histórico</button>
                </p>
                <hr>
            `;
        });
    } else {
        resultadoDiv.innerHTML = `<p>Nenhuma correspondência encontrada para o fragmento "${fragmento}".</p>`;
    }
}

let resultadosDescricao = [];
let palavraChaveAtual = '';
let indiceAtual = 0;

function exibirResultadosDescricao(resultados, palavraChave) {
    resultadosDescricao = resultados;
    palavraChaveAtual = palavraChave;
    indiceAtual = 0;
    mostrarLoteDescricao();
}

function mostrarLoteDescricao() {
    const resultadoDiv = document.getElementById('resultado');

    if (indiceAtual === 0) {
        resultadoDiv.innerHTML = `<h3>🔎 Resultados para: <strong>${palavraChaveAtual}</strong></h3>`;
    }

    const lote = resultadosDescricao.slice(indiceAtual, indiceAtual + 10);
    lote.forEach(resultado => {
        resultadoDiv.innerHTML += `
            <p>
                ${resultado.resultadoTratado}<br>
                <button onclick="adicionarAoHistorico(\`${resultado.resultadoTratado}\`)">Adicionar ao Histórico</button>
            </p>
            <hr>
        `;
    });

    indiceAtual += 10;

    if (indiceAtual < resultadosDescricao.length) {
        resultadoDiv.innerHTML += `
            <div style="text-align:center; margin: 1rem;">
                <button onclick="mostrarLoteDescricao()">Carregar mais</button>
            </div>
        `;
    }
}

// Função para adicionar um item ao histórico
function adicionarAoHistorico(resultadoTratado) {
    var listaHistorico = document.getElementById('lista-historico');
    var novoItem = document.createElement('li');
    novoItem.innerHTML = resultadoTratado; // Adiciona o resultado já formatado
    listaHistorico.appendChild(novoItem);
}

// Função para normalizar o código ATM
function normalizarCodigoATM(codigo) {
    if (!codigo) { // Verifica se o código está vazio ou indefinido
        return ''; // Retorna uma string vazia
    }
    return codigo.replace(/\s+/g, '').replace(/[^a-zA-Z0-9]/g, '').toUpperCase();
}

// Ativa busca quando pressionar ENTER
document.getElementById('numero').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        const checkbox = document.getElementById('ativar-busca-avancada');
        if (checkbox.checked) {
            verificarModoBusca(); // Corrigido aqui
        } else {
            buscar();
        }
    }
});

// Alterna entre busca simples e avançada (botão e menu)
function toggleBuscaAvancada() {
    const checkbox = document.getElementById('ativar-busca-avancada');
    const menuBuscaAvancada = document.getElementById('menu-busca-avancada');
    const botaoBusca = document.querySelector('.Busca button');

    if (checkbox.checked) {
        menuBuscaAvancada.style.display = 'block';
        botaoBusca.textContent = 'Buscar Avançada';
        botaoBusca.setAttribute('onclick', 'verificarModoBusca()');
    } else {
        menuBuscaAvancada.style.display = 'none';
        botaoBusca.textContent = 'Buscar';
        botaoBusca.setAttribute('onclick', 'buscar()');
    }
}

// Decide qual função de busca chamar com base no modo selecionado
function verificarModoBusca() {
    const modo = document.getElementById('modo-busca').value;
    if (modo === 'descricao') {
        buscarPorDescricao();
    } else {
        buscarATMAvancado();
    }
}

function limparHistorico() {
    const listaHistorico = document.getElementById('lista-historico');
    listaHistorico.innerHTML = ''; // Remove todos os itens do histórico
}
function copiarhistorico() {
    const lista = document.getElementById('lista-historico');
    const itens = Array.from(lista.getElementsByTagName('li')).map(item => item.innerText).join('\n');

    if (!itens) {
        alert('O histórico está vazio.');
        return;
    }

    navigator.clipboard.writeText(itens).then(() => {
        alert('Histórico copiado para a área de transferência.');
    }).catch(err => {
        console.error('Erro ao copiar histórico:', err);
        alert('Não foi possível copiar o histórico.');
    });
}

