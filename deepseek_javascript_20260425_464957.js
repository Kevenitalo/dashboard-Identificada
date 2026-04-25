let charts = {};
let dadosCompletos = [];

// Elementos DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const dashboard = document.getElementById('dashboard');
const dataAtualizacaoSpan = document.getElementById('dataAtualizacao');
const btnLimparDados = document.getElementById('btnLimparDados');
const filtroProduto = document.getElementById('filtroProduto');
const filtroTipo = document.getElementById('filtroTipo');

// Configurar event listeners
uploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileUpload);
if (btnLimparDados) btnLimparDados.addEventListener('click', limparDados);
if (filtroProduto) filtroProduto.addEventListener('input', filtrarTabela);
if (filtroTipo) filtroTipo.addEventListener('change', filtrarTabela);

// Drag and drop
if (uploadArea) {
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
            processExcel(file);
        } else {
            alert('Por favor, envie um arquivo .xlsx ou .xls válido');
        }
    });
}

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
        processExcel(file);
    }
}

function limparDados() {
    if (confirm('Tem certeza que deseja limpar todos os dados salvos?')) {
        localStorage.removeItem('dashboardData');
        localStorage.removeItem('lastUpdate');
        dadosCompletos = [];
        dashboard.style.display = 'none';
        uploadArea.style.display = 'block';
        alert('Dados limpos com sucesso!');
    }
}

// Função para encontrar coluna por sinônimos
function encontrarColuna(row, possiveisNomes) {
    for (let nome of possiveisNomes) {
        if (row.hasOwnProperty(nome)) {
            return nome;
        }
        // Também procura case insensitive
        for (let key in row) {
            if (key.toLowerCase() === nome.toLowerCase()) {
                return key;
            }
        }
    }
    return null;
}

function processExcel(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            let todosDados = [];
            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                jsonData.forEach(row => {
                    row._aba = sheetName;
                });
                todosDados = todosDados.concat(jsonData);
            });
            
            console.log('Primeira linha do arquivo:', todosDados[0]);
            console.log('Colunas encontradas:', Object.keys(todosDados[0] || {}));
            
            processData(todosDados);
            
            localStorage.setItem('dashboardData', JSON.stringify(todosDados));
            localStorage.setItem('lastUpdate', new Date().toISOString());
            
            dashboard.style.display = 'block';
            uploadArea.style.display = 'none';
            
            const hoje = new Date().toLocaleDateString('pt-BR');
            dataAtualizacaoSpan.textContent = hoje;
            
            alert(`✅ Processado! ${todosDados.length} registros encontrados.`);
        } catch (error) {
            alert('Erro ao processar o arquivo: ' + error.message);
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function processData(data) {
    dadosCompletos = [];
    
    let perdasMaturacao = 0;
    let perdasAvaria = 0;
    let perdasVencimento = 0;
    let valorMaturacao = 0;
    let valorAvaria = 0;
    let valorVencimento = 0;
    
    let perdasPorLoja = {};
    let perdasPorProduto = {};
    let perdasPorMotivo = {};
    
    // Primeira linha para identificar colunas
    const primeiraLinha = data[0] || {};
    
    // Mapear colunas (aceita vários nomes)
    const colLoja = encontrarColuna(primeiraLinha, ['Loja', 'LOJA', 'loja', 'FILIAL', 'Filial']);
    const colProduto = encontrarColuna(primeiraLinha, ['Produto', 'PRODUTO', 'produto', 'ITEM', 'Item', 'DESCRICAO', 'Descricao']);
    const colMotivo = encontrarColuna(primeiraLinha, ['Desc. Motivo', 'Desc Motivo', 'MOTIVO', 'Motivo', 'TIPO', 'Tipo']);
    const colPeso = encontrarColuna(primeiraLinha, ['Peso (Kg)', 'Peso', 'peso', 'KG', 'Kg']);
    const colQuantidade = encontrarColuna(primeiraLinha, ['Quantidade', 'quantidade', 'QTD', 'Qtd']);
    const colValor = encontrarColuna(primeiraLinha, ['Valor', 'valor', 'VALOR', 'R$', 'Vlr']);
    
    console.log('Colunas mapeadas:', { colLoja, colProduto, colMotivo, colPeso, colQuantidade, colValor });
    
    data.forEach(row => {
        // Extrair valores usando as colunas encontradas
        const loja = colLoja ? (row[colLoja] || 'Não especificada').toString() : 'Não especificada';
        const produto = colProduto ? (row[colProduto] || 'Não especificado').toString() : 'Não especificado';
        let descMotivo = colMotivo ? (row[colMotivo] || '').toString() : '';
        
        // Se não achou Desc. Motivo, tenta outras colunas
        if (!descMotivo) {
            for (let key in row) {
                if (key.toLowerCase().includes('motivo') || key.toLowerCase().includes('tipo')) {
                    descMotivo = row[key] || '';
                    break;
                }
            }
        }
        
        // Peso ou Quantidade
        let peso = 0;
        if (colPeso) peso = Math.abs(parseFloat(row[colPeso]) || 0);
        if (peso === 0 && colQuantidade) peso = Math.abs(parseFloat(row[colQuantidade]) || 0);
        
        const valor = colValor ? Math.abs(parseFloat(row[colValor]) || 0) : 0;
        
        const qtdPerda = peso;
        
        // Classificar tipo
        let tipo = 'Outros';
        const textoMotivo = descMotivo.toLowerCase();
        
        if (textoMotivo.includes('maturação') || textoMotivo.includes('maturacao')) {
            tipo = 'Maturação';
            perdasMaturacao += qtdPerda;
            valorMaturacao += valor;
        } else if (textoMotivo.includes('avaria')) {
            tipo = 'Avaria';
            perdasAvaria += qtdPerda;
            valorAvaria += valor;
        } else if (textoMotivo.includes('vencimento')) {
            tipo = 'Vencimento';
            perdasVencimento += qtdPerda;
            valorVencimento += valor;
        }
        
        // Acumular estatísticas
        if (!perdasPorLoja[loja]) perdasPorLoja[loja] = 0;
        perdasPorLoja[loja] += qtdPerda;
        
        const nomeProduto = produto.length > 40 ? produto.substring(0, 40) + '...' : produto;
        if (!perdasPorProduto[nomeProduto]) perdasPorProduto[nomeProduto] = 0;
        perdasPorProduto[nomeProduto] += qtdPerda;
        
        const motivoChave = descMotivo.split(' - ')[0] || descMotivo || 'Sem motivo';
        if (!perdasPorMotivo[motivoChave]) perdasPorMotivo[motivoChave] = 0;
        perdasPorMotivo[motivoChave] += qtdPerda;
        
        // Adicionar aos detalhes
        if (qtdPerda > 0 && tipo !== 'Outros') {
            dadosCompletos.push({
                loja,
                produto: nomeProduto,
                descMotivo: descMotivo || 'Não informado',
                tipo,
                quantidade: qtdPerda,
                valor
            });
        }
    });
    
    // Atualizar cards
    const totalMaturacaoEl = document.getElementById('totalMaturacao');
    const totalAvariaEl = document.getElementById('totalAvaria');
    const totalVencimentoEl = document.getElementById('totalVencimento');
    const valorMaturacaoEl = document.getElementById('valorMaturacao');
    const valorAvariaEl = document.getElementById('valorAvaria');
    const valorVencimentoEl = document.getElementById('valorVencimento');
    
    if (totalMaturacaoEl) totalMaturacaoEl.textContent = formatarNumero(perdasMaturacao);
    if (totalAvariaEl) totalAvariaEl.textContent = formatarNumero(perdasAvaria);
    if (totalVencimentoEl) totalVencimentoEl.textContent = formatarNumero(perdasVencimento);
    if (valorMaturacaoEl) valorMaturacaoEl.textContent = `R$ ${valorMaturacao.toFixed(2)}`;
    if (valorAvariaEl) valorAvariaEl.textContent = `R$ ${valorAvaria.toFixed(2)}`;
    if (valorVencimentoEl) valorVencimentoEl.textContent = `R$ ${valorVencimento.toFixed(2)}`;
    
    criarGraficos(perdasMaturacao, perdasAvaria, perdasVencimento, perdasPorLoja, perdasPorProduto, perdasPorMotivo);
    atualizarTabela();
}

function formatarNumero(valor) {
    if (valor >= 1000) {
        return (valor / 1000).toFixed(1) + 'k kg';
    }
    return valor.toFixed(1) + ' kg';
}

function criarGraficos(maturacao, avaria, vencimento, perdasPorLoja, perdasPorProduto, perdasPorMotivo) {
    // Gráfico comparativo
    const ctxComp = document.getElementById('chartComparativo');
    if (ctxComp && (maturacao > 0 || avaria > 0 || vencimento > 0)) {
        if (charts.comparativo) charts.comparativo.destroy();
        charts.comparativo = new Chart(ctxComp, {
            type: 'pie',
            data: {
                labels: ['Maturação', 'Avaria', 'Vencimento'],
                datasets: [{
                    data: [maturacao, avaria, vencimento],
                    backgroundColor: ['#FFB347', '#FF6B6B', '#4ECDC4'],
                    borderWidth: 0
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { position: 'bottom' } }
            }
        });
    }
    
    // Gráfico por loja
    const lojasOrdenadas = Object.entries(perdasPorLoja).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctxLojas = document.getElementById('chartLojas');
    if (ctxLojas && lojasOrdenadas.length > 0) {
        if (charts.lojas) charts.lojas.destroy();
        charts.lojas = new Chart(ctxLojas, {
            type: 'bar',
            data: {
                labels: lojasOrdenadas.map(l => l[0].length > 15 ? l[0].substring(0, 15) + '...' : l[0]),
                datasets: [{
                    label: 'Perda (kg)',
                    data: lojasOrdenadas.map(l => l[1]),
                    backgroundColor: '#667eea',
                    borderRadius: 8
                }]
            },
            options: { responsive: true, maintainAspectRatio: true }
        });
    }
    
    // Gráfico por produto
    const produtosOrdenados = Object.entries(perdasPorProduto).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const ctxProdutos = document.getElementById('chartProdutos');
    if (ctxProdutos && produtosOrdenados.length > 0) {
        if (charts.produtos) charts.produtos.destroy();
        charts.produtos = new Chart(ctxProdutos, {
            type: 'bar',
            data: {
                labels: produtosOrdenados.map(p => p[0].length > 20 ? p[0].substring(0, 20) + '...' : p[0]),
                datasets: [{
                    label: 'Perda (kg)',
                    data: produtosOrdenados.map(p => p[1]),
                    backgroundColor: '#764ba2',
                    borderRadius: 8
                }]
            },
            options: { responsive: true, maintainAspectRatio: true, indexAxis: 'y' }
        });
    }
}

function atualizarTabela() {
    const tbody = document.getElementById('tabelaBody');
    if (!tbody) return;
    
    const tipoSelecionado = filtroTipo ? filtroTipo.value : 'todos';
    const textoFiltro = filtroProduto ? filtroProduto.value.toLowerCase() : '';
    
    let dadosFiltrados = dadosCompletos;
    
    if (tipoSelecionado !== 'todos') {
        dadosFiltrados = dadosFiltrados.filter(d => d.tipo === tipoSelecionado);
    }
    
    if (textoFiltro) {
        dadosFiltrados = dadosFiltrados.filter(d => 
            d.produto.toLowerCase().includes(textoFiltro) || 
            d.descMotivo.toLowerCase().includes(textoFiltro)
        );
    }
    
    dadosFiltrados.sort((a, b) => b.quantidade - a.quantidade);
    
    tbody.innerHTML = '';
    
    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align: center;">📁 Nenhum registro encontrado</td></tr>';
        const contador = document.getElementById('contadorRegistros');
        if (contador) contador.textContent = '';
        return;
    }
    
    dadosFiltrados.slice(0, 200).forEach(item => {
        const row = tbody.insertRow();
        row.insertCell(0).textContent = item.loja;
        row.insertCell(1).textContent = item.produto;
        row.insertCell(2).textContent = item.descMotivo;
        row.insertCell(3).textContent = item.tipo;
        row.insertCell(4).textContent = item.quantidade.toFixed(2);
        row.insertCell(5).textContent = `R$ ${item.valor.toFixed(2)}`;
    });
    
    const contador = document.getElementById('contadorRegistros');
    if (contador) {
        contador.textContent = `Mostrando ${Math.min(dadosFiltrados.length, 200)} de ${dadosFiltrados.length} registros`;
    }
}

function filtrarTabela() {
    atualizarTabela();
}

function loadSavedData() {
    const savedData = localStorage.getItem('dashboardData');
    if (savedData) {
        const data = JSON.parse(savedData);
        processData(data);
        dashboard.style.display = 'block';
        uploadArea.style.display = 'none';
        
        const lastUpdate = localStorage.getItem('lastUpdate');
        if (lastUpdate && dataAtualizacaoSpan) {
            const dataUp = new Date(lastUpdate).toLocaleDateString('pt-BR');
            dataAtualizacaoSpan.textContent = dataUp;
        }
    }
}

loadSavedData();
