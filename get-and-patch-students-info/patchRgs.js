// storing students data into an array
var data = [];

// creating function to fill DOM elements
function fill(elementProps, value, index, msgElement) {
    var fillWithInteval = setInterval(function(elementProps, value, index = 0, msgElement = document.querySelector('div.msg-mensagem')) {
        var element = document.querySelectorAll(elementProps)[index]
        if (element && !msgElement) {
            element.value = value;
            return clearInterval(fillWithInteval);
        }
    }, 500, elementProps, value, index, msgElement);
}

// creating function to clickOn DOM elements
function clickOn(elementProps, index, msgElement) {
    var clickOnWithInterval = setInterval(function(elementProps, index = 0, msgElement = document.querySelector('div.msg-mensagem')) {
        var element = document.querySelectorAll(elementProps)[index]
        if (element && !msgElement) {
            element.click();      
            return clearInterval(clickOnWithInterval);
        }
    }, 500, elementProps, index, msgElement);
}

// running data array
for (i = 0; i < data.length; i++) {
    // selecting RA on select tag    
    fill('#TipoConsultaFichaAluno option', '1');

    // fill RA field
    fill('#txtRa', data[i].ra);

    // click on search button
    clickOn('#btnPesquisar')

    // click on button to open student details
    clickOn('td a i.icone-tabela-editar')

    // fill rg data
    fill('#RgAluno', data[i].rg);
    fill('#DigRgAluno', data[i].dig);
    fill('#sgUfRg option', data[i].uf);
    fill('#dtEmisRg', data[i].emis);

    // fill missing document justification
    fill('#JustiificativaDocumento option', '1');

    // click on "Atualizar" button
    clickOn('div.modal-lg div.modal-content div.modal-footer button.btn-info');

    // click on "Voltar" button
    clickOn('div.modal-lg div.modal-content div.modal-footer button.btn-info', 1);
}
