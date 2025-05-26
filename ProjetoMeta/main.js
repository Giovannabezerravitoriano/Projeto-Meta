let dadosExcel = [];
const FILTROS_KEY = 'filtrosUtilizados'; // Chave para o localStorage

// --- Funções para Gerenciar Histórico de Filtros ---

/**
 * Carrega os filtros salvos do localStorage.
 * @returns {Array} Um array de objetos de filtro.
 */
function carregarHistoricoFiltros() {
    const filtrosSalvosJSON = localStorage.getItem(FILTROS_KEY);
    return filtrosSalvosJSON ? JSON.parse(filtrosSalvosJSON) : [];
}

/**
 * Salva o histórico de filtros no localStorage.
 * @param {Array} filtros - O array de objetos de filtro a ser salvo.
 */
function salvarHistoricoFiltros(filtros) {
    localStorage.setItem(FILTROS_KEY, JSON.stringify(filtros));
}

/**
 * Adiciona um novo filtro ao histórico, evitando duplicatas, e atualiza a exibição.
 * @param {Object} filtro - O objeto de filtro a ser adicionado.
 */
function adicionarFiltroAoHistorico(filtro) {
    let filtros = carregarHistoricoFiltros();

    // Lógica para verificar se o filtro já existe, incluindo entradas manuais
    const filtroExistente = filtros.find(f =>
        f.campo === filtro.campo &&
        f.percentual === filtro.percentual &&
        // Verifica se ambos são entradas manuais e compara os valores manuais
        ((f.totalManual !== null && f.totalManual !== undefined && filtro.totalManual !== null && filtro.totalManual !== undefined && f.campoValorManual === filtro.campoValorManual && f.totalManual === filtro.totalManual) ||
         // Ou se nenhum é manual e compara os valores normais
         (f.totalManual === null && filtro.totalManual === null && f.valor === filtro.valor))
    );

    if (!filtroExistente) {
        filtros.push(filtro);
        salvarHistoricoFiltros(filtros);
        exibirHistoricoFiltros(); // Atualiza a exibição após adicionar
    }
}

/**
 * Exibe os filtros do histórico na interface do usuário.
 */
function exibirHistoricoFiltros() {
    const listaFiltros = document.getElementById('listaFiltros');
    listaFiltros.innerHTML = ''; // Limpa a lista antes de preencher

    const filtros = carregarHistoricoFiltros();

    if (filtros.length === 0) {
        listaFiltros.innerHTML = '<li>Nenhum filtro utilizado ainda.</li>';
        return;
    }

    filtros.forEach((filtro, index) => {
        const li = document.createElement('li');
        let filtroTexto = '';

        if (filtro.totalManual !== null && filtro.totalManual !== undefined) {
            // Se for um filtro com valor manual
            const campoNome = filtro.campo ? `Campo: <strong>${filtro.campo}</strong><br>` : '';
            const valorManualNome = filtro.campoValorManual ? `Valor: <strong>${filtro.campoValorManual}</strong><br>` : '';
            filtroTexto = `${campoNome}${valorManualNome}Total Manual: <strong>R$ ${parseFloat(filtro.totalManual).toFixed(2).replace('.', ',')}</strong>`;
        } else if (filtro.campo && filtro.valor) {
            // Se for um filtro da planilha com campo e valor específico
            filtroTexto = `<strong>${filtro.campo}</strong> = <strong>${filtro.valor}</strong>`;
        } else if (filtro.campo && !filtro.valor) {
            // Se for um filtro da planilha com campo selecionado mas valor "(Todos)"
            filtroTexto = `<strong>${filtro.campo}</strong>: (Todos os valores)`;
        } else {
            // Se nenhum filtro foi aplicado (Total geral da planilha)
            filtroTexto = 'nenhum (Total Geral da Planilha)';
        }

        li.innerHTML = `
            <div>
                Filtro: ${filtroTexto}<br>
                Aumento: <strong>${filtro.percentual}%</strong>
            </div>
            <button class="remove-filtro-btn" data-index="${index}">X</button>
        `;
        listaFiltros.appendChild(li);
    });

    // Adiciona event listeners para os botões de remoção individual
    document.querySelectorAll('.remove-filtro-btn').forEach(button => {
        button.addEventListener('click', function() {
            const indexToRemove = parseInt(this.dataset.index);
            removerFiltroIndividual(indexToRemove);
        });
    });
}

/**
 * Remove um filtro específico do histórico e atualiza a exibição.
 * @param {number} index - O índice do filtro a ser removido no array do histórico.
 */
function removerFiltroIndividual(index) {
    let filtros = carregarHistoricoFiltros();
    if (index >= 0 && index < filtros.length) {
        filtros.splice(index, 1); // Remove o item no índice especificado
        salvarHistoricoFiltros(filtros);
        exibirHistoricoFiltros(); // Atualiza a exibição após remover
    }
}

/**
 * Limpa todo o histórico de filtros salvos no localStorage.
 */
function limparHistoricoFiltros() {
    if (confirm('Tem certeza que deseja limpar todo o histórico de filtros?')) {
        localStorage.removeItem(FILTROS_KEY);
        exibirHistoricoFiltros(); // Atualiza a exibição após limpar
    }
}

// --- Funções Principais e Event Listeners ---

// Executado quando o DOM está completamente carregado
document.addEventListener('DOMContentLoaded', function() {
    // Carrega o percentual salvo ao iniciar a página
    const savedPercentual = localStorage.getItem('percentual');
    if (savedPercentual) {
        document.getElementById('percentual').value = savedPercentual;
    }
    exibirHistoricoFiltros(); // Exibe o histórico de filtros ao carregar a página
});

// Event listener para a seleção do arquivo Excel
document.getElementById('arquivo').addEventListener('change', function () {
    const file = this.files[0];
    if (!file) return;

    const loading = document.getElementById('loading');
    loading.style.display = 'block'; // Mostra mensagem de carregamento

    const reader = new FileReader();

    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            const sheetName = workbook.SheetNames[0]; // Assume a primeira aba da planilha
            const sheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(sheet); // Converte a aba para JSON

            if (json.length === 0) {
                alert("A planilha está vazia ou mal formatada.");
            } else {
                dadosExcel = json; // Armazena os dados do Excel globalmente
                document.getElementById('campo').disabled = false; // Habilita o select de campo
                document.getElementById('valorFiltro').disabled = false; // Habilita o select de valor do filtro
                preencherOpcoes(); // Preenche as opções do filtro e a opção manual

                // Tenta carregar os valores salvos de campo, valor e campos manuais
                const savedCampo = localStorage.getItem('campo');
                const savedValor = localStorage.getItem('valorFiltro');
                const savedValorManual = localStorage.getItem('valorManual');
                const savedCampoValorManual = localStorage.getItem('campoValorManual');

                if (savedCampo) {
                    document.getElementById('campo').value = savedCampo;
                    preencherOpcoes(); // Chama novamente para preencher opções com base no campo salvo
                    if (savedValor) {
                        document.getElementById('valorFiltro').value = savedValor;
                        // Se o valor salvo for 'manual-input', mostra os campos manuais e preenche-os
                        if (savedValor === 'manual-input') {
                           document.getElementById('manualInputContainer').style.display = 'block';
                           if (savedValorManual) document.getElementById('valorManual').value = savedValorManual;
                           if (savedCampoValorManual) document.getElementById('campoValorManual').value = savedCampoValorManual;
                        } else {
                           document.getElementById('manualInputContainer').style.display = 'none'; // Esconde se não for manual
                           document.getElementById('campoValorManual').value = '';
                           document.getElementById('valorManual').value = '';
                        }
                    }
                }
            }
        } catch (erro) {
            alert("Erro ao ler o arquivo. Verifique se o formato está correto e se é um arquivo .xlsx válido.");
            console.error(erro);
        } finally {
            loading.style.display = 'none'; // Esconde mensagem de carregamento
        }
    };

    reader.onerror = function () {
        alert("Erro ao carregar o arquivo.");
        loading.style.display = 'none';
    };

    reader.readAsArrayBuffer(file); // Lê o arquivo como um ArrayBuffer
});

// Event listener para o input de percentual (salva no localStorage)
document.getElementById('percentual').addEventListener('input', function() {
    localStorage.setItem('percentual', this.value);
});

// Event listener para o select de campo (salva no localStorage e reseta outros campos)
document.getElementById('campo').addEventListener('change', function() {
    localStorage.setItem('campo', this.value);
    preencherOpcoes(); // Repreencha as opções de valor com base no novo campo
    localStorage.removeItem('valorFiltro'); // Limpa o valor do filtro salvo
    document.getElementById('valorFiltro').value = ''; // Reseta o select de valor
    localStorage.removeItem('valorManual'); // Limpa o valor manual salvo
    localStorage.removeItem('campoValorManual'); // Limpa o nome/valor manual do campo salvo
    document.getElementById('valorManual').value = ''; // Limpa o input de valor manual
    document.getElementById('campoValorManual').value = ''; // Limpa o input de nome/valor manual
    document.getElementById('manualInputContainer').style.display = 'none'; // Esconde o container manual
});

// Event listener para o select de valor do filtro (salva no localStorage e controla a exibição do input manual)
document.getElementById('valorFiltro').addEventListener('change', function() {
    localStorage.setItem('valorFiltro', this.value);
    const manualInputContainer = document.getElementById('manualInputContainer');
    const valorManualInput = document.getElementById('valorManual');
    const campoValorManualInput = document.getElementById('campoValorManual');

    if (this.value === 'manual-input') {
        manualInputContainer.style.display = 'block'; // Mostra o container manual
        campoValorManualInput.focus(); // Foca no campo de nome manual para digitação
    } else {
        manualInputContainer.style.display = 'none'; // Esconde o container manual
        valorManualInput.value = ''; // Limpa o valor manual
        campoValorManualInput.value = ''; // Limpa o nome/valor manual
        localStorage.removeItem('valorManual'); // Remove do localStorage
        localStorage.removeItem('campoValorManual'); // Remove do localStorage
    }
});

// Event listener para o input de valor manual (salva no localStorage)
document.getElementById('valorManual').addEventListener('input', function() {
    localStorage.setItem('valorManual', this.value);
});

// Event listener para o input de nome/valor manual do campo (salva no localStorage)
document.getElementById('campoValorManual').addEventListener('input', function() {
    localStorage.setItem('campoValorManual', this.value);
});


/**
 * Preenche as opções do select de valor do filtro com base no campo selecionado
 * e adiciona a opção de entrada manual.
 */
function preencherOpcoes() {
    const campo = document.getElementById('campo').value;
    const valorFiltroSelect = document.getElementById('valorFiltro');
    const manualInputContainer = document.getElementById('manualInputContainer');

    valorFiltroSelect.innerHTML = '<option value="">(Todos)</option>'; // Sempre começa com a opção "Todos"
    manualInputContainer.style.display = 'none'; // Esconde o campo manual por padrão
    document.getElementById('valorManual').value = ''; // Garante que o valor manual esteja vazio
    document.getElementById('campoValorManual').value = ''; // Garante que o nome/valor manual esteja vazio

    if (!campo || dadosExcel.length === 0) {
        valorFiltroSelect.disabled = true; // Desabilita o select se não houver campo ou dados
        return;
    }

    // Coleta valores únicos para o campo selecionado na planilha
    const valoresUnicos = [...new Set(dadosExcel.map(d => d[campo]).filter(v => v !== undefined && v !== null))];

    // Ordena os valores únicos (trata números e strings)
    valoresUnicos.sort((a, b) => {
        const numA = parseFloat(a);
        const numB = parseFloat(b);
        if (!isNaN(numA) && !isNaN(numB)) {
            return numA - numB; // Ordenação numérica
        }
        return String(a).localeCompare(String(b)); // Ordenação alfabética
    }).forEach(v => {
        const opt = document.createElement('option');
        opt.value = v;
        opt.textContent = v;
        valorFiltroSelect.appendChild(opt);
    });

    // Adiciona a opção "Inserir valor manualmente" após os valores do Excel
    const optManual = document.createElement('option');
    optManual.value = 'manual-input';
    optManual.textContent = `-- Inserir valor manual para ${campo} --`; // Adapta o texto
    valorFiltroSelect.appendChild(optManual);

    valorFiltroSelect.disabled = false; // Habilita o select
}

/**
 * Calcula a meta de faturamento com base nos filtros aplicados (ou entrada manual).
 */
function calcularMeta() {
    const percentualInput = document.getElementById('percentual');
    const campoSelect = document.getElementById('campo');
    const valorFiltroSelect = document.getElementById('valorFiltro');
    const valorManualInput = document.getElementById('valorManual');
    const campoValorManualInput = document.getElementById('campoValorManual');
    const manualInputContainer = document.getElementById('manualInputContainer');

    const percentual = parseFloat(percentualInput.value);
    const campo = campoSelect.value;
    const valor = valorFiltroSelect.value; // Pode ser um valor do Excel ou 'manual-input'
    const valorManual = parseFloat(valorManualInput.value);
    const campoValorManual = campoValorManualInput.value.trim(); // Pega o texto do campo manual e remove espaços

    if (isNaN(percentual)) {
        document.getElementById('resultado').innerHTML = "<p>Informe o percentual de aumento corretamente.</p>";
        return;
    }

    let total = 0;
    let filtroDisplay = '';
    let isManualInput = false;
    let valorParaHistorico = valor; // Valor a ser salvo no histórico (para filtros do Excel)
    let campoValorManualParaHistorico = null; // Para salvar o nome/valor manual do campo

    // Lógica para determinar o total a ser usado: prioriza entrada manual
    if (valor === 'manual-input') {
        if (!campo) {
            document.getElementById('resultado').innerHTML = "<p>Por favor, selecione um 'Filtrar por' (Consultor, Filial, Produto) antes de inserir um valor manual.</p>";
            return;
        }
        if (campoValorManual === '') {
             document.getElementById('resultado').innerHTML = `<p>Por favor, insira o nome/valor do ${campo} para a entrada manual (Ex: João Vitor).</p>`;
             return;
        }
        if (isNaN(valorManual) || valorManual < 0) {
            document.getElementById('resultado').innerHTML = "<p>Por favor, insira um valor numérico válido e positivo para o faturamento manual.</p>";
            return;
        }
        total = valorManual;
        filtroDisplay = `Total manual para <strong>${campo}: ${campoValorManual}</strong>`;
        isManualInput = true;
        valorParaHistorico = null; // Não salva o valor do filtro se for manual
        campoValorManualParaHistorico = campoValorManual; // Salva o nome/valor manual
    } else {
        // Lógica de filtro da planilha (Excel)
        let dadosFiltrados = dadosExcel;

        if (campo && valor) { // Filtro por campo e valor específico do Excel
            dadosFiltrados = dadosExcel.filter(linha => String(linha[campo]) === String(valor));
            filtroDisplay = `<strong>${campo}</strong> = <strong>${valor}</strong>`;
        } else if (campo && !valor) { // Filtro por campo, mas com "(Todos)" selecionado
            // Se o campo está selecionado, mas o valor é "(Todos)", calculamos o total geral da planilha
            // mas ainda exibimos o campo para contexto.
            filtroDisplay = `<strong>${campo}</strong>: (Todos os valores)`;
        } else { // Sem filtro (Total Geral da Planilha)
            filtroDisplay = 'nenhum (Total Geral da Planilha)';
        }

        // Soma os valores da coluna 'Valor' dos dados filtrados
        dadosFiltrados.forEach(linha => {
            if (linha['Valor'] !== undefined && !isNaN(parseFloat(linha['Valor']))) {
                total += parseFloat(linha['Valor']);
            }
        });

        // Se o total filtrado da planilha for 0 e não for uma entrada manual
        if (total <= 0 && !isManualInput && (campo && valor)) {
            document.getElementById('resultado').innerHTML = `<p>Nenhum dado encontrado para o filtro <strong>${campo} = ${valor}</strong>. Tente um filtro diferente ou use a opção "Inserir valor manual para ${campo}".</p>`;
            return;
        }
    }

    const meta = total * (1 + percentual / 100);

    const resultadoDiv = document.getElementById('resultado');
    resultadoDiv.innerHTML = `
        <p>Filtro aplicado: ${filtroDisplay}</p>
        <p>Total ${isManualInput ? 'informado' : 'filtrado'}: <strong>R$ ${total.toFixed(2).replace('.', ',')}</strong></p>
        <p>Meta com ${percentual}% de aumento: <strong>R$ ${meta.toFixed(2).replace('.', ',')}</strong></p>
    `;

    // Salva o filtro atual no histórico
    adicionarFiltroAoHistorico({
        campo: campo,
        valor: valorParaHistorico, // null se for entrada manual, senão o valor do filtro da planilha
        percentual: percentual,
        totalManual: isManualInput ? total : null, // Salva o total manual se for o caso
        campoValorManual: campoValorManualParaHistorico // Salva o nome/valor manual do campo
    });
}