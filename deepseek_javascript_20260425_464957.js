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
btnLimparDados.addEventListener('click', limparDados);
filtroProduto.addEventListener('input', filtrarTabela);
filtroTipo.addEventListener('change', filtrarTabela);

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
    if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
        processExcel(file);
    } else {
        alert('Por favor, envie um arquivo .xlsx ou .xls válido');
    }
});

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

function processExcel(file) {
    mostrarLoading('Processando arquivo...');
    
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
            
            processData(todosDados);
            
            localStorage.setItem('dashboardData', JSON.stringify(todosDados));
            localStorage.setItem('lastUpdate', new Date().toISOString());
            
            dashboard.style.display = 'block';
            uploadArea.style.display = 'none';
            
            const hoje = new Date().toLocaleDateString('pt-BR');
            dataAtualizacaoSpan.textContent = hoje;
            
            esconderLoading();
            alert(`✅ Processado com sucesso! ${todosDados.length} registros encontrados.`);
        } catch (error) {
            esconderLoading();
            alert('Erro ao processar o arquivo. Verifique se o formato está correto.');
            console.error(error);
        }
    };
    reader.readAsArrayBuffer(file);
}

function mostrarLoading(mensagem) {
    let loading = document.getElementById('loading');
    if (!loading) {
        loading = document.createElement('div');
        loading.id = 'loading';
        loading.className = 'loading';
        loading.innerHTML = mensagem;
        document.body.appendChild(loading);
    } else {
        loading.innerHTML = mensagem;
        loading.style.display = 'block';
    }
}

function esconderLoading() {
    const loading = document.getElementById('loading');
    if (loading) loading.style.display = 'none';
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
    
    data.forEach(row => {
        const loja = row['Loja'] || 'Não especificada';
        const produto = (row['Produto'] || row['produto'] || 'Não especificado').toString();
        const descMotivo = row['Desc. Motivo'] || '';
        const peso = Math.abs(parseFloat(row['Peso (Kg)'] || row['peso'] || 0));
        const quantidade = Math.abs(parseFloat(row['Quantidade'] || 0));
        const valor = Math.abs(parseFloat(row['Valor'] || 0));
        
        const qtdPerda = peso > 0 ? peso : quantidade;
        
        let tipo = 'Outros';
        if (descMotivo.toLowerCase().includes('maturação')) {
            tipo = 'Maturação';
            perdasMaturacao += qtdPerda;
            valorMaturacao += valor;
        } else if (descMotivo.toLowerCase().includes('avaria')) {
            tipo = 'Avaria';
            perdasAvaria += qtdPerda;
            valorAvaria += valor;
        } else if (descMotivo.toLowerCase().includes('vencimento')) {
            tipo = 'Vencimento';
            perdasVencimento += qtdPerda;
            valorVencimento += valor;
        }
        
        if (!perdasPorLoja[loja]) perdasPorLoja[loja] = 0;
        perdasPorLoja[loja] += qtdPerda;
        
        const nomeProduto = produto.length > 40 ? produto.substring(0, 40) + '...' : produto;
        if (!perdasPorProduto[nomeProduto]) perdasPorProduto[nomeProduto] = 0;
        perdasPorProduto[nomeProduto] += qtdPerda;
        
        const motivoChave = descMotivo.split(' - ')[0] || descMotivo || 'Sem motivo';
        if (!perdasPorMotivo[motivoChave]) perdasPorMotivo[motivoChave] = 0;
        perdasPorMotivo[motivoChave] += qtdPerda;
        
        if (qtdPerda > 0) {
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
    
    document.getElementById('totalMaturacao').textContent = formatarNumero(perdasMaturacao);
    document.getElementById('totalAvaria').textContent = formatarNumero(perdasAvaria);
    document.getElementById('totalVencimento').textContent = formatarNumero(perdasVencimento);
    document.getElementById('valorMaturacao').textContent = `R$ ${valorMaturacao.toFixed(2)}`;
    document.getElementById('valorAvaria').textContent = `R$ ${valorAvaria.toFixed(2)}`;
    document.getElementById('valorVencimento').textContent = `R$ ${valorVencimento.toFixed(2)}`;
    
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
    if (ctxComp) {
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
                plugins: {
                    legend: { position: 'bottom' },
                    tooltip: { callbacks: { label: (ctx) => `${ctx.label}: ${ctx.raw.toFixed(1)} kg` } }
                }
            }
        });
    }
    
    // Gráfico por loja
    const lojasOrdenadas = Object.entries(perdasPorLoja).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctxLojas = document.getElementById('chartLojas');
    if (ctxLojas) {
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
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { position: 'top' } }
            }
        });
    }
    
    // Gráfico por produto
    const produtosOrdenados = Object.entries(perdasPorProduto).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const ctxProdutos = document.getElementById('chartProdutos');
    if (ctxProdutos) {
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
            options: {
                responsive: true,
                maintainAspectRatio: true,
                indexAxis: 'y'
            }
        });
    }
    
    // Gráfico por motivo
    const motivosOrdenados = Object.entries(perdasPorMotivo).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctxMotivos = document.getElementById('chartMotivos');
    if (ctxMotivos) {
        if (charts.motivos) charts.motivos.destroy();
        charts.motivos = new Chart(ctxMotivos, {
            type: 'pie',
            data: {
                labels: motivosOrdenados.map(m => m[0].length > 20 ? m[0].substring(0, 20) + '...' : m[0]),
                datasets: [{
                    data: motivosOrdenados.map(m => m[1]),
                    backgroundColor: ['#FFB347', '#FF6B6B', '#4ECDC4', '#95A5A6', '#3498DB', '#E74C3C', '#2ECC71', '#F39C12']
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: { legend: { position: 'right' } }
            }
        });
    }
}

function atualizarTabela() {
    const tipoSelecionado = filtroTipo.value;
    const textoFiltro = filtroProduto.value.toLowerCase();
    
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
    
    const tbody = document.getElementById('tabelaBody');
    tbody.innerHTML = '';
    
    if (dadosFiltrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align: center;">📁 Nenhum registro encontrado</td></tr>';
        document.getElementById('contadorRegistros').textContent = '';
        return;
    }
    
    dadosFiltrados.slice(0, 200).forEach(item => {
        const row = tbody.insertRow();
        row.className = `tipo-${item.tipo.toLowerCase()}`;
        row.insertCell(0).textContent = item.loja;
        row.insertCell(1).textContent = item.produto;
        row.insertCell(2).textContent = item.descMotivo;
        row.insertCell(3).textContent = item.tipo;
        row.insertCell(4).textContent = item.quantidade.toFixed(2);
        row.insertCell(5).textContent = `R$ ${item.valor.toFixed(2)}`;
    });
    
    document.getElementById('contadorRegistros').textContent = 
        `Mostrando ${Math.min(dadosFiltrados.length, 200)} de ${dadosFiltrados.length} registros`;
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
        if (lastUpdate) {
            const dataUp = new Date(lastUpdate).toLocaleDateString('pt-BR');
            dataAtualizacaoSpan.textContent = dataUp;
        }
    }
}

loadSavedData();
