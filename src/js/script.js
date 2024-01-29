var workbook;

fetch('BENS.xlsx')
.then(response => response.arrayBuffer())
.then(data => {
   var data = new Uint8Array(data);
   workbook = XLSX.read(data, {type: 'array'});
});

const anoAtualSpan = document.getElementById("anoAtual");
        const dataAtual = new Date();
        const ano = dataAtual.getFullYear();
        anoAtualSpan.textContent = ano;

function formatInput(input) {
    var noLeadingZeros = input.replace(/^0+/, '');
    var formattedInput = noLeadingZeros.replace(/-\d$/, '');
    return formattedInput;
}

var codigo = new URLSearchParams(window.location.search).get("=");
document.getElementById("numero").value = codigo;

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

document.getElementById('numero').addEventListener('keypress', function(e) {
    if (e.key === 'Enter') {
        buscar();
    }
});

function buscar() {
    var inputField = document.getElementById('numero');
    var formattedInput = formatInput(inputField.value);
    
    var worksheet = workbook.Sheets[workbook.SheetNames[0]];
    var jsonData = XLSX.utils.sheet_to_json(worksheet, {header: 1});
   
   for (var i = 0; i < jsonData.length; i++) {
       if (jsonData[i][0] == formattedInput || jsonData[i][2] == formattedInput) {
           var translatedCondition = translateCondition(jsonData[i][3]);
           var translatedSituation = translateSituation(jsonData[i][4]);
           
           document.getElementById('resultado').innerHTML = 'Número de patrimônio: ' + jsonData[i][0] + '-' + jsonData[i][1] +'<br>' +
                                                            'Tipo: ' + jsonData[i][8] + '<br>' +
                                                            'Descrição: ' + jsonData[i][5] + '<br>' +
                                                            'Situação: ' + translatedSituation +'<br>' +
                                                            'Condição do Bem: ' + translatedCondition + '<br>' +
                                                            'Local da Guarda: ' + jsonData[i][7] + '<br>' +
                                                            'Responsável: ' + jsonData[i][9];
           inputField.value = '';
           inputField.focus();
           return;
       }
   }
   
   document.getElementById('resultado').innerHTML = "Número não encontrado no banco de dados.";
   
   inputField.value = '';
   inputField.focus();
}

window.onload = function() {
  document.getElementById('numero').focus();
};
