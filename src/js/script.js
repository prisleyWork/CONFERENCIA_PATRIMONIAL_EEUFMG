// Declaração de variáveis para armazenar os workbooks
var workbooks = {
    ativos: null,
    baixados: null,
    internos: null
};

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
    var inputField = document.getElementById('numero');
    var modoBusca = document.getElementById('modo-busca').value; // Verifica o modo selecionado (ATM ou patrimônio)
    var fragmento = inputField.value.trim(); // Remove espaços extras do input
    var resultados = []; // Armazena os resultados encontrados

    // Validação: o fragmento de busca não pode estar vazio
    if (!fragmento) {
        document.getElementById('resultado').innerHTML = '<p>Por favor, insira pelo menos um fragmento para busca avançada.</p>';
        return;
    }

    // Realiza a busca em cada planilha
    ['ativos', 'baixados', 'internos'].forEach(estrutura => {
        var workbook = workbooks[estrutura];
        if (workbook) {
            var worksheet = workbook.Sheets[workbook.SheetNames[0]];
            var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            jsonData.forEach(item => {
                // Define qual campo verificar com base no modo de busca (ATM ou patrimônio)
                var codigoBanco = modoBusca === 'atm' ? String(item[2]) : String(item[0]);

                // Verifica se o fragmento está contido no campo correspondente
                if (codigoBanco.includes(fragmento)) {
                    let translatedCondition = translateCondition(item[3]);
                    let translatedSituation = translateSituation(item[5]);
                    let atmInfo = item[2];
                    let atmDisplay = atmInfo ? `<br><b>ATM:</b> ${atmInfo}` : '';

                    // Formatação dos resultados
                    let resultadoTratado = `
                        <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                        <b>Tipo:</b> ${item[25]}<br>
                        <b>Descrição:</b> ${item[8]}<br>
                        <b>Situação:</b> ${translatedSituation}<br>
                        <b>Condição do Bem:</b> ${translatedCondition}<br>
                        <b>Local da Guarda:</b> ${item[17]}<br>
                        <b>Responsável:</b> ${item[27]}${atmDisplay}
                    `;

                    resultados.push({
                        resultadoTratado: resultadoTratado,
                        codigoBanco: codigoBanco,
                        modo: modoBusca === 'atm' ? 'ATM' : 'Patrimônio'
                    });

                    // Limita os resultados a 5 itens
                    if (resultados.length >= 5) return;
                }
            });
        } else {
            console.error(`Planilha ${estrutura} não carregada corretamente.`);
        }
    });

    // Exibe os resultados ou uma mensagem de erro
    exibirResultadosTratados(resultados, fragmento);
}

// Função para busca por descrição com base em palavras-chave
function buscarPorDescricao() {
    var inputField = document.getElementById('numero');
    var palavraChave = inputField.value.trim(); // Remove espaços extras do input
    var resultados = []; // Armazena os resultados encontrados

    // Validação: a palavra-chave não pode estar vazia
    if (!palavraChave) {
        document.getElementById('resultado').innerHTML = '<p>Por favor, insira ao menos uma palavra-chave para busca por descrição.</p>';
        return;
    }

    // Realiza a busca em cada planilha
    ['ativos', 'baixados'].forEach(estrutura => {
        var workbook = workbooks[estrutura];
        if (workbook) {
            var worksheet = workbook.Sheets[workbook.SheetNames[0]];
            var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            jsonData.forEach(item => {
                // Definir a descrição com base na estrutura
                var descricaoCampo = estrutura === 'ativos' ? item[8] : item[2]; // Seleciona a coluna correta
                var descricao = String(descricaoCampo); // Converte para string

                // Verifica se a palavra-chave está contida na descrição
                if (descricao.toLowerCase().includes(palavraChave.toLowerCase())) {
                    let translatedCondition = translateCondition(item[3]);
                    let translatedSituation = translateSituation(item[5]);

                    // Formatação dos resultados
                    let resultadoTratado = `
                        <b>Número de patrimônio:</b> ${item[0]}-${item[1]}<br>
                        <b>Tipo:</b> ${item[25]}<br>
                        <b>Descrição:</b> ${descricao}<br>
                        <b>Situação:</b> ${translatedSituation}<br>
                        <b>Condição do Bem:</b> ${translatedCondition}<br>
                        <b>Local da Guarda:</b> ${item[17]}<br>
                        <b>Responsável:</b> ${item[27]}
                    `;

                    resultados.push({
                        resultadoTratado: resultadoTratado,
                        descricao: descricao
                    });

                    // Limita os resultados a 5 itens
                    if (resultados.length >= 5) return;
                }
            });
        } else {
            console.error(`Planilha ${estrutura} não carregada corretamente.`);
        }
    });

    // Exibe os resultados ou uma mensagem de erro
    exibirResultadosDescricao(resultados, palavraChave);
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

function exibirResultadosDescricao(resultados, palavraChave) {
    var resultadoDiv = document.getElementById('resultado');
    if (resultados.length > 0) {
        resultadoDiv.innerHTML = `<h3>Resultados para "${palavraChave}":</h3>`;
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
        resultadoDiv.innerHTML = `<p>Nenhuma correspondência encontrada para "${palavraChave}".</p>`;
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

// Função para ativar busca avançada ao pressionar Enter
document.getElementById('numero').addEventListener('keypress', function (e) {
    if (e.key === 'Enter') {
        var checkbox = document.getElementById('ativar-busca-avancada');
        if (checkbox.checked) {
            buscarATMAvancado();
        } else {
            buscar(); // Realiza a busca normal
        }
    }
});

// Função para exibir/ocultar o menu avançado baseado na checkbox
function toggleBuscaAvancada() {
    var checkbox = document.getElementById('ativar-busca-avancada');
    var menuBuscaAvancada = document.getElementById('menu-busca-avancada');
    var botaoBusca = document.querySelector('.Busca button'); // Seleciona o botão de busca padrão

    if (checkbox.checked) {
        // Exibe o menu de busca avançada e altera o texto do botão
        menuBuscaAvancada.style.display = 'block';
        botaoBusca.textContent = 'Buscar Avançada'; // Altera o texto do botão
        botaoBusca.setAttribute('onclick', 'buscarATMAvancado()'); // Atualiza a ação do botão
    } else {
        // Oculta o menu de busca avançada e restaura o botão padrão
        menuBuscaAvancada.style.display = 'none';
        botaoBusca.textContent = 'Buscar'; // Restaura o texto original do botão
        botaoBusca.setAttribute('onclick', 'buscar()'); // Restaura a ação do botão padrão
    }
}