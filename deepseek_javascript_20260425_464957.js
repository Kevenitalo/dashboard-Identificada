let charts = {};

// Elementos DOM
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const dashboard = document.getElementById('dashboard');
const dataAtualizacaoSpan = document.getElementById('dataAtualizacao');

// Configurar event listeners
uploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileUpload);

// Drag and drop
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
    if (file && file.name.endsWith('.xlsx')) {
        processExcel(file);
    } else {
        alert('Por favor, envie um arquivo .xlsx válido');
    }
});

function handleFileUpload(e) {
    const file = e.target.files[0];
    if (file) {
        processExcel(file);
    }
}

// Processar arquivo Excel
function processExcel(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        
        // Pegar todas as abas (sheets)
        let todosDados = [];
        workbook.SheetNames.forEach(sheetName => {
            const sheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(sheet);
            // Adicionar informação da aba
            jsonData.forEach(row => {
                row._aba = sheetName;
            });
            todosDados = todosDados.concat(jsonData);
        });
        
        processData(todosDados);
        
        // Salvar no localStorage
        localStorage.setItem('dashboardData', JSON.stringify(todosDados));
        localStorage.setItem('lastUpdate', new Date().toISOString());
        
        // Mostrar dashboard
        dashboard.style.display = 'block';
        uploadArea.style.display = 'none';
        
        const hoje = new Date().toLocaleDateString('pt-BR');
        dataAtualizacaoSpan.textContent = hoje;
    };
    reader.readAsArrayBuffer(file);
}

// Função principal de processamento
function processData(data) {
    let perdasMaturacao = 0;
    let perdasAvaria = 0;
    let perdasVencimento = 0;
    let perdasOutros = 0;
    let valorTotalPerdas = 0;
    let detalhes = [];
    
    // Estatísticas por loja
    let perdasPorLoja = {};
    let perdasPorProduto = {};
    let perdasPorMotivo = {};
    
    data.forEach(row => {
        // Mapear colunas do seu Excel
        const loja = row['Loja'] || 'Não especificada';
        const produto = row['Produto'] || row['Produto'] || 'Não especificado';
        const descMotivo = row['Desc. Motivo'] || '';
        const quantidade = Math.abs(parseFloat(row['Quantidade'] || row['Quantidade'] || 0));
        const peso = Math.abs(parseFloat(row['Peso (Kg)'] || row['Peso (Kg)'] || 0));
        const valor = Math.abs(parseFloat(row['Valor'] || row['Valor'] || 0));
        const tipoMov = row['Tipo Mov'] || '';
        const descricao = row['Descrição'] || '';
        
        // Usar peso como principal (já que tem coluna específica)
        const qtdPerda = peso > 0 ? peso : quantidade;
        
        // Classificar tipo de perda baseado no "Desc. Motivo"
        let tipo = 'Outros';
        if (descMotivo.toLowerCase().includes('maturação')) {
            tipo = 'Maturação';
            perdasMaturacao += qtdPerda;
        } else if (descMotivo.toLowerCase().includes('avaria')) {
            tipo = 'Avaria';
            perdasAvaria += qtdPerda;
        } else if (descMotivo.toLowerCase().includes('vencimento')) {
            tipo = 'Vencimento';
            perdasVencimento += qtdPerda;
        } else {
            perdasOutros += qtdPerda;
        }
        
        valorTotalPerdas += valor;
        
        // Acumular por loja
        if (!perdasPorLoja[loja]) perdasPorLoja[loja] = 0;
        perdasPorLoja[loja] += qtdPerda;
        
        // Acumular por produto
        const nomeProduto = String(produto).substring(0, 40);
        if (!perdasPorProduto[nomeProduto]) perdasPorProduto[nomeProduto] = 0;
        perdasPorProduto[nomeProduto] += qtdPerda;
        
        // Acumular por motivo detalhado
        const motivoChave = descMotivo.split(' - ')[0] || descMotivo;
        if (!perdasPorMotivo[motivoChave]) perdasPorMotivo[motivoChave] = 0;
        perdasPorMotivo[motivoChave] += qtdPerda;
        
        // Adicionar aos detalhes para tabela
        if (qtdPerda > 0) {
            detalhes.push({
                loja,
                produto: nomeProduto,
                descMotivo,
                tipo,
                quantidade: qtdPerda,
                peso,
                valor,
                tipoMov,
                descricao
            });
        }
    });
    
    // Atualizar cards
    document.getElementById('totalMaturacao').textContent = formatarNumero(perdasMaturacao);
    document.getElementById('totalAvaria').textContent = formatarNumero(perdasAvaria);
    document.getElementById('totalVencimento').textContent = formatarNumero(perdasVencimento);
    
    // Adicionar card de resumo
    adicionarCardsResumo(perdasMaturacao, perdasAvaria, perdasVencimento, perdasOutros, valorTotalPerdas);
    
    // Atualizar tabela
    updateTabela(detalhes);
    
    // Criar gráficos
    createCharts(perdasMaturacao, perdasAvaria, perdasVencimento, perdasPorLoja, perdasPorProduto, perdasPorMotivo);
}

function formatarNumero(valor) {
    if (valor >= 1000) {
        return (valor / 1000).toFixed(1) + 'k kg';
    }
    return valor.toFixed(1) + ' kg';
}

function adicionarCardsResumo(maturacao, avaria, vencimento, outros, valorTotal) {
    const cardsContainer = document.querySelector('.cards');
    
    // Card de total geral
    let totalCard = document.getElementById('cardTotal');
    if (!totalCard) {
        totalCard = document.createElement('div');
        totalCard.className = 'card';
        totalCard.id = 'cardTotal';
        totalCard.innerHTML = `
            <h3>📦 Total de Perdas</h3>
            <div class="valor" id="totalPerdas">${(maturacao + avaria + vencimento + outros).toFixed(1)} kg</div>
            <div class="valor-pequeno">💰 R$ ${valorTotal.toFixed(2)}</div>
        `;
        cardsContainer.appendChild(totalCard);
    } else {
        document.getElementById('totalPerdas').textContent = (maturacao + avaria + vencimento + outros).toFixed(1) + ' kg';
    }
}

// Tabela de detalhes
function updateTabela(detalhes) {
    const tbody = document.getElementById('tabelaBody');
    tbody.innerHTML = '';
    
    // Ordenar por quantidade (maiores perdas primeiro)
    detalhes.sort((a, b) => b.quantidade - a.quantidade);
    
    // Mostrar top 100 registros
    detalhes.slice(0, 100).forEach(item => {
        const row = tbody.insertRow();
        row.insertCell(0).textContent = item.loja;
        row.insertCell(1).textContent = item.produto;
        row.insertCell(2).textContent = item.descMotivo;
        row.insertCell(3).textContent = item.tipo;
        row.insertCell(4).textContent = item.quantidade.toFixed(2);
        row.insertCell(5).textContent = item.peso > 0 ? item.peso.toFixed(2) : '-';
        row.insertCell(6).textContent = item.valor > 0 ? 'R$ ' + item.valor.toFixed(2) : '-';
        
        // Adicionar classe de cor conforme tipo
        if (item.tipo === 'Maturação') row.style.borderLeft = '4px solid #FFB347';
        else if (item.tipo === 'Avaria') row.style.borderLeft = '4px solid #FF6B6B';
        else if (item.tipo === 'Vencimento') row.style.borderLeft = '4px solid #4ECDC4';
    });
    
    if (detalhes.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" style="text-align: center;">Nenhum dado encontrado. Faça upload do arquivo Excel.</td></tr>';
    }
}

// Gráficos
function createCharts(maturacao, avaria, vencimento, perdasPorLoja, perdasPorProduto, perdasPorMotivo) {
    // Gráfico 1: Comparativo por tipo (Pizza)
    const ctxComparativo = document.getElementById('chartComparativo');
    if (ctxComparativo) {
        if (charts.comparativo) charts.comparativo.destroy();
        charts.comparativo = new Chart(ctxComparativo, {
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
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${ctx.raw.toFixed(1)} kg` } }
                }
            }
        });
    }
    
    // Gráfico 2: Perdas por Loja (Barras)
    const lojasOrdenadas = Object.entries(perdasPorLoja)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);
    
    const ctxLojas = document.getElementById('chartLojas');
    if (ctxLojas) {
        if (charts.lojas) charts.lojas.destroy();
        charts.lojas = new Chart(ctxLojas, {
            type: 'bar',
            data: {
                labels: lojasOrdenadas.map(l => l[0].substring(0, 15)),
                datasets: [{
                    label: 'Perda (kg)',
                    data: lojasOrdenadas.map(l => l[1]),
                    backgroundColor: '#667eea'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { position: 'top' } }
            }
        });
    }
    
    // Gráfico 3: Top Produtos com maior perda
    const produtosOrdenados = Object.entries(perdasPorProduto)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    const ctxProdutos = document.getElementById('chartProdutos');
    if (ctxProdutos) {
        if (charts.produtos) charts.produtos.destroy();
        charts.produtos = new Chart(ctxProdutos, {
            type: 'bar',
            data: {
                labels: produtosOrdenados.map(p => p[0].substring(0, 20)),
                datasets: [{
                    label: 'Perda (kg)',
                    data: produtosOrdenados.map(p => p[1]),
                    backgroundColor: '#764ba2'
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                indexAxis: 'y'
            }
        });
    }
    
    // Gráfico 4: Top Motivos de perda
    const motivosOrdenados = Object.entries(perdasPorMotivo)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);
    
    const ctxMotivos = document.getElementById('chartMotivos');
    if (ctxMotivos) {
        if (charts.motivos) charts.motivos.destroy();
        charts.motivos = new Chart(ctxMotivos, {
            type: 'pie',
            data: {
                labels: motivosOrdenados.map(m => m[0]),
                datasets: [{
                    data: motivosOrdenados.map(m => m[1]),
                    backgroundColor: ['#FFB347', '#FF6B6B', '#4ECDC4', '#95A5A6', '#3498DB', '#E74C3C', '#2ECC71', '#F39C12']
                }]
            },
            options: {
                responsive: true,
                plugins: { legend: { position: 'right' } }
            }
        });
    }
}

// Carregar dados salvos
function loadSavedData() {
    const savedData = localStorage.getItem('dashboardData');
    if (savedData) {
        const data = JSON.parse(savedData);
        processData(data);
        dashboard.style.display = 'block';
        uploadArea.style.display = 'none';
        
        const lastUpdate = localStorage.getItem('lastUpdate');
        if (lastUpdate) {
            const dataUp = new Date(lastUpdate).toLocaleDateString('pt-BR');
            dataAtualizacaoSpan.textContent = dataUp;
        }
    }
}

loadSavedData();