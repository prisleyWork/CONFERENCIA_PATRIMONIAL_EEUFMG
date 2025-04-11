var workbooks = {
    ativos: null,
    baixados: null,
    internos: null
};

// Função para carregar as planilhas
function carregarPlanilhas() {
    Promise.all([
        fetch('ATIVOS.xlsx').then(response => response.arrayBuffer()).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.ativos = XLSX.read(arrayData, {type: 'array'});
        }),
        fetch('BAIXADOS.xlsx').then(response => response.arrayBuffer()).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.baixados = XLSX.read(arrayData, {type: 'array'});
        }),
        fetch('INTERNOS.xlsx').then(response => response.arrayBuffer()).then(data => {
            var arrayData = new Uint8Array(data);
            workbooks.internos = XLSX.read(arrayData, {type: 'array'});
        })
    ]).then(() => {
        console.log("Todas as planilhas foram carregadas.");
    });
}

carregarPlanilhas();

// Atualizar o ano no rodapé
const anoAtualSpan = document.getElementById("anoAtual");
const dataAtual = new Date();
anoAtualSpan.textContent = dataAtual.getFullYear();

// Função para formatar entrada
function formatInput(input) {
    var noLeadingZeros = input.replace(/^0+/, '');
    var formattedInput = noLeadingZeros.replace(/-\d{1,2}$/, '');
    return formattedInput;
}

// Funções para traduzir condição e situação
function translateCondition(condition) {
    var translations = {
        'BM': 'Bom',
        'AE': 'Anti-Econômico',
        'IR': 'Irrecuperável',
        'OC': 'Ocioso',
        'BX': 'Baixado',
        'RE': 'Recuperável'
    };
    return translations[condition] || condition;
}

function translateSituation(situation) {
    var translations = {
        'NI': 'Não encontrado no local da guarda',
        'NO': 'Normal'
    };
    return translations[situation] || situation;
}

// Função para buscar informações
function buscar() {
    var inputField = document.getElementById('numero');
    var formattedInput = formatInput(inputField.value);
    var resultado = document.getElementById('resultado');
    var listaHistorico = document.getElementById('lista-historico');

    // Função interna para buscar em uma planilha específica
    function buscarNaPlanilha(workbook, estrutura) {
        var worksheet = workbook.Sheets[workbook.SheetNames[0]];
        var jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});

        for (var i = 0; i < jsonData.length; i++) {
            if (estrutura === "ATIVOS_INTERNOS") {
                // Estrutura para as planilhas ATIVOS e INTERNOS
                if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
                    var translatedCondition = translateCondition(jsonData[i][3]);
                    var translatedSituation = translateSituation(jsonData[i][5]);

                    // Verificar campo ATM
                    var atmInfo = jsonData[i][2];
                    var atmDisplay = atmInfo ? '<br><b>ATM:</b> ' + atmInfo : '';

                    return '<b>Número de patrimônio:</b> ' + jsonData[i][0] + '-' + jsonData[i][1] + '<br>' +
                           '<b>Tipo:</b> ' + jsonData[i][25] + '<br>' +
                           '<b>Descrição:</b> ' + jsonData[i][8] + '<br>' +
                           '<b>Situação:</b> ' + translatedSituation + '<br>' +
                           '<b>Condição do Bem:</b> ' + translatedCondition + '<br>' +
                           '<b>Local da Guarda:</b> ' + jsonData[i][17] + '<br>' +
                           '<b>Responsável:</b> ' + jsonData[i][27] +
                           atmDisplay;
                }
            } else if (estrutura === "BAIXADOS") {
                // Estrutura específica da planilha BAIXADOS
                if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
                    return '<b>Número de patrimônio:</b> ' + jsonData[i][0] + '-' + jsonData[i][1] + '<br>' +
                           '<b>Número ATM:</b> ' + jsonData[i][3] + '<br>' +
                           '<b>Setor:</b> ' + jsonData[i][10] + '<br>' +
                           '<b>Descrição:</b> ' + jsonData[i][2] + '<br>' +
                           '<b>Último Local da Guarda:</b> ' + jsonData[i][13] + '<br>' +
                           '<b>Observação:</b> <b>Bens baixados devem ser mantidos no local de guarda atual. Caso deseje desfazer do bem, cadastre no sistema de desfazimento.</b>';
                }
            }
        }
        return null;
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

    inputField.value = '';
    inputField.focus();
}

// Função para copiar o histórico mantendo a formatação
function copiarHistorico() {
    var listaHistorico = document.getElementById('lista-historico');
    var spanMensagem = document.createElement('span');
    spanMensagem.id = 'mensagem-copiado';
    spanMensagem.textContent = 'Histórico copiado!';
    spanMensagem.style.color = 'green';

    // Criação de um elemento temporário para copiar o conteúdo
    var range = document.createRange();
    range.selectNode(listaHistorico);
    var selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);

    try {
        document.execCommand('copy');
        // Exibir mensagem de sucesso
        listaHistorico.appendChild(spanMensagem);
        setTimeout(() => spanMensagem.remove(), 3000); // Remove o aviso após 3 segundos
    } catch (err) {
        console.error('Falha ao copiar', err);
    }

    // Remover a seleção
    selection.removeAllRanges();
}

// Configurar evento para buscar ao pressionar Enter
document.getElementById('numero').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        buscar();
    }
});

// Focar automaticamente no campo de entrada ao carregar a página
window.onload = function() {
    document.getElementById('numero').focus();
};