let hot; // Inicialização da variável hot
const fileInput = document.getElementById('fileInput');
const controls = document.getElementById('controls');
const historyContainer = document.getElementById('historyContainer');
const historyList = document.getElementById('historyList');
const history = []; // Array para armazenar as tabelas salvas

// Inicializa a Handsontable com dados vazios ou dados passados
function initHandsontable(data = [[]]) {
    const container = document.getElementById('hot');

    // Se já existir uma instância de hot, destrua-a
    if (hot) {
        hot.destroy();
    }

    // Cria uma nova instância da Handsontable
    hot = new Handsontable(container, {
        data: data,
        rowHeaders: true,
        colHeaders: true,
        filters: true,
        dropdownMenu: true,
        contextMenu: true,
        licenseKey: 'non-commercial-and-evaluation',
        afterChange: function (changes, source) {
            if (source === 'edit') {
                changes.forEach(([row, col, oldValue, newValue]) => {
                    if (newValue && newValue.startsWith('=')) {
                        const result = evaluateFormula(newValue.substring(1)); // Remove o "=" e avalia
                        hot.setDataAtCell(row, col, result); // Define o resultado na célula
                    }
                });
            }
        }
    });
}

// Função para avaliar a fórmula (apenas somas neste exemplo)
function evaluateFormula(formula) {
    const cellReferences = formula.split('+');
    let sum = 0;

    cellReferences.forEach(ref => {
        const col = ref.charCodeAt(0) - 65; // Converte 'A' para índice 0
        const row = parseInt(ref.substring(1)) - 1;

        const cellValue = hot.getDataAtCell(row, col);
        if (!isNaN(cellValue)) {
            sum += parseFloat(cellValue);
        }
    });

    return sum;
}

// Função para carregar arquivo Excel
function loadExcelFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        if (jsonData.length === 0 || jsonData[0].length === 0) {
            alert('O arquivo Excel não contém dados válidos.');
            return;
        }

        // Inicializa a Handsontable com os dados carregados
        initHandsontable(jsonData);
        controls.style.display = 'flex'; // Exibe os botões de controle
    };
    reader.readAsArrayBuffer(file);
}

// Event Listener para o input de arquivo
fileInput.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        loadExcelFile(file);
    }
});

// Event Listener para o botão Criar Tabela
document.getElementById('createTable').addEventListener('click', () => {
    initHandsontable(); // Inicia com uma tabela vazia
    controls.style.display = 'flex'; // Exibe os botões de controle
});

// Adiciona uma linha
document.getElementById('addRow').addEventListener('click', () => {
    hot.alter('insert_row', hot.countRows()); // Adiciona na última posição
});

// Adiciona uma coluna
document.getElementById('addCol').addEventListener('click', () => {
    hot.alter('insert_col', hot.countCols()); // Adiciona na última posição
});

// Exclui a última linha
document.getElementById('deleteRow').addEventListener('click', () => {
    const lastRowIndex = hot.countRows() - 1;
    if (lastRowIndex >= 0) {
        hot.alter('remove_row', lastRowIndex); // Remove a última linha
    }
});

// Exclui a última coluna
document.getElementById('deleteCol').addEventListener('click', () => {
    const lastColIndex = hot.countCols() - 1;
    if (lastColIndex >= 0) {
        hot.alter('remove_col', lastColIndex); // Remove a última coluna
    }
});

// Salva a tabela
document.getElementById('saveTable').addEventListener('click', () => {
    const tableData = hot.getData();
    if (tableData.length === 0) {
        alert('Não há dados para salvar.');
        return;
    }

    // Verifica se a tabela já existe no histórico
    const existingEntryIndex = history.findIndex(entry => 
        JSON.stringify(entry.data) === JSON.stringify(tableData)
    );

    if (existingEntryIndex !== -1) {
        const overwrite = confirm('Esta tabela já existe. Deseja sobrescrever as alterações?');
        if (overwrite) {
            const tableEntry = {
                name: history[existingEntryIndex].name,
                data: tableData,
                date: new Date()
            };
            history[existingEntryIndex] = tableEntry;
            alert('Tabela atualizada com sucesso!');
        }
    } else {
        const tableName = prompt('Nomeie sua tabela:') || 'Tabela Sem Nome';
        const tableEntry = {
            name: tableName,
            data: tableData,
            date: new Date()
        };
        history.push(tableEntry);
        alert('Tabela salva com sucesso!');
    }

    displayHistory(); // Atualiza a exibição do histórico
});

// Fechar a tabela atual
document.getElementById('closeTable').addEventListener('click', closeTable);

// Baixar CSV
document.getElementById('downloadCsv').addEventListener('click', () => {
    const csvData = hot.getData().map(row => row.join(',')).join('\n'); // Converte para CSV
    const blob = new Blob([csvData], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'tabela.csv'; // Nome do arquivo a ser baixado
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a); // Remove o elemento temporário
});

// Mostrar/ocultar histórico
document.getElementById('toggleHistory').addEventListener('click', () => {
    if (historyContainer.style.display === 'none') {
        historyContainer.style.display = 'block'; // Mostra o histórico
        displayHistory(); // Exibe o histórico
    } else {
        historyContainer.style.display = 'none'; // Oculta o histórico
    }
});

// Exibir histórico
function displayHistory() {
    historyList.innerHTML = ''; // Limpa a lista
    history.forEach((entry, index) => {
        const div = document.createElement('div');
        div.textContent = `${entry.name} - ${entry.date.toLocaleString()}`;
        div.style.cursor = 'pointer';
        
        const buttonContainer = document.createElement('span');
        const openButton = document.createElement('button');
        openButton.textContent = 'Abrir';
        openButton.onclick = () => openTable(entry);
        buttonContainer.appendChild(openButton);

        const deleteButton = document.createElement('button');
        deleteButton.textContent = 'Excluir';
        deleteButton.onclick = () => {
            const confirmDelete = confirm(`Deseja realmente excluir a tabela "${entry.name}"?`);
            if (confirmDelete) {
                history.splice(index, 1);
                displayHistory();
                closeTable();
                alert('Tabela excluída com sucesso!');
            }
        };
        buttonContainer.appendChild(deleteButton);
        div.appendChild(buttonContainer);
        
        historyList.appendChild(div);
    });
}

// Abre tabela do histórico
function openTable(entry) {
    initHandsontable(entry.data);
    controls.style.display = 'flex';
}

// Fecha a tabela atual
function closeTable() {
    initHandsontable([]);
    controls.style.display = 'none';
}
