let charts = {};
let dadosCompletos = [];

const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const dashboard = document.getElementById('dashboard');
const dataAtualizacaoSpan = document.getElementById('dataAtualizacao');
const btnLimpar = document.getElementById('btnLimpar');
const filtroProduto = document.getElementById('filtroProduto');
const filtroTipo = document.getElementById('filtroTipo');

uploadBtn.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileUpload);
btnLimpar.addEventListener('click', limparDados);
filtroProduto.addEventListener('input', filtrarTabela);
filtroTipo.addEventListener('change', filtrarTabela);

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
        alert('✅ Dados limpos com sucesso!');
    }
}

function encontrarColuna(row, possiveisNomes) {
    for (let nome of possiveisNomes) {
        if (row.hasOwnProperty(nome)) {
            return nome;
        }
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
                todosDados = todosDados.concat(jsonData);
            });
            
            if (todosDados.length === 0) {
                alert('Nenhum dado encontrado no arquivo!');
                return;
            }
            
            processData(todosDados);
            
            localStorage.setItem('dashboardData', JSON.stringify(todosDados));
            localStorage.setItem('lastUpdate', new Date().toISOString());
            
            dashboard.style.display = 'block';
            uploadArea.style.display = 'none';
            dataAtualizacaoSpan.textContent = new Date().toLocaleDateString('pt-BR');
            
            alert(`✅ Sucesso! ${todosDados.length} registros processados.`);
        } catch (error) {
            alert('Erro ao processar o arquivo: ' + error.message);
        }
    };
    reader.readAsArrayBuffer(file);
}

function processData(data) {
    dadosCompletos = [];
    
    let maturacaoKg = 0, avariaKg = 0, vencimentoKg = 0;
    let maturacaoValor = 0, avariaValor = 0, vencimentoValor = 0;
    
    let perdasPorLoja = {};
    let perdasPorProduto = {};
    let perdasPorMotivo = {};
    
    const primeiraLinha = data[0];
    const colLoja = encontrarColuna(primeiraLinha, ['Loja', 'LOJA', 'loja', 'FILIAL']);
    const colProduto = encontrarColuna(primeiraLinha, ['Produto', 'PRODUTO', 'produto', 'ITEM']);
    const colMotivo = encontrarColuna(primeiraLinha, ['Desc. Motivo', 'Desc Motivo', 'MOTIVO', 'Motivo']);
    const colPeso = encontrarColuna(primeiraLinha, ['Peso (Kg)', 'Peso', 'peso', 'KG']);
    const colValor = encontrarColuna(primeiraLinha, ['Valor', 'valor', 'VALOR']);
    
    console.log('Colunas encontradas:', { colLoja, colProduto, colMotivo, colPeso, colValor });
    
    data.forEach(row => {
        const loja = colLoja ? (row[colLoja] || 'Não especificada').toString() : 'Não especificada';
        let produto = colProduto ? (row[colProduto] || 'Não especificado').toString() : 'Não especificado';
        let descMotivo = colMotivo ? (row[colMotivo] || '').toString() : '';
        
        if (produto.length > 45) produto = produto.substring(0, 45) + '...';
        
        let peso = 0;
        if (colPeso) peso = Math.abs(parseFloat(row[colPeso]) || 0);
        if (peso === 0) {
            for (let key in row) {
                if (key.toLowerCase().includes('quantidade') || key.toLowerCase().includes('qtd')) {
                    peso = Math.abs(parseFloat(row[key]) || 0);
                    break;
                }
            }
        }
        
        const valor = colValor ? Math.abs(parseFloat(row[colValor]) || 0) : 0;
        
        let tipo = 'Outros';
        const textoMotivo = descMotivo.toLowerCase();
        
        if (textoMotivo.includes('maturação') || textoMotivo.includes('maturacao')) {
            tipo = 'Maturação';
            maturacaoKg += peso;
            maturacaoValor += valor;
        } else if (textoMotivo.includes('avaria')) {
            tipo = 'Avaria';
            avariaKg += peso;
            avariaValor += valor;
        } else if (textoMotivo.includes('vencimento')) {
            tipo = 'Vencimento';
            vencimentoKg += peso;
            vencimentoValor += valor;
        }
        
        if (peso > 0 && tipo !== 'Outros') {
            if (!perdasPorLoja[loja]) perdasPorLoja[loja] = 0;
            perdasPorLoja[loja] += peso;
            
            if (!perdasPorProduto[produto]) perdasPorProduto[produto] = 0;
            perdasPorProduto[produto] += peso;
            
            const motivoChave = descMotivo.split(' - ')[0] || descMotivo || 'Sem motivo';
            if (!perdasPorMotivo[motivoChave]) perdasPorMotivo[motivoChave] = 0;
            perdasPorMotivo[motivoChave] += peso;
            
            dadosCompletos.push({
                loja,
                produto,
                descMotivo: descMotivo || 'Não informado',
                tipo,
                quantidade: peso,
                valor
            });
        }
    });
    
    document.getElementById('totalMaturacao').textContent = formatarNumero(maturacaoKg);
    document.getElementById('totalAvaria').textContent = formatarNumero(avariaKg);
    document.getElementById('totalVencimento').textContent = formatarNumero(vencimentoKg);
    document.getElementById('valorMaturacao').textContent = `R$ ${maturacaoValor.toFixed(2)}`;
    document.getElementById('valorAvaria').textContent = `R$ ${avariaValor.toFixed(2)}`;
    document.getElementById('valorVencimento').textContent = `R$ ${vencimentoValor.toFixed(2)}`;
    
    criarGraficos(maturacaoKg, avariaKg, vencimentoKg, perdasPorLoja, perdasPorProduto, perdasPorMotivo);
    atualizarTabela();
}

function formatarNumero(valor) {
    if (valor >= 1000) {
        return (valor / 1000).toFixed(1) + 'k kg';
    }
    return valor.toFixed(1) + ' kg';
}

function criarGraficos(m, a, v, lojas, produtos, motivos) {
    const ctxComp = document.getElementById('chartComparativo');
    if (ctxComp) {
        if (charts.comparativo) charts.comparativo.destroy();
        charts.comparativo = new Chart(ctxComp, {
            type: 'pie',
            data: {
                labels: ['Maturação', 'Avaria', 'Vencimento'],
                datasets: [{ data: [m, a, v], backgroundColor: ['#FFB347', '#FF6B6B', '#4ECDC4'] }]
            },
            options: { responsive: true, plugins: { legend: { position: 'bottom' } } }
        });
    }
    
    const lojasOrdenadas = Object.entries(lojas).sort((a, b) => b[1] - a[1]).slice(0, 8);
    const ctxLojas = document.getElementById('chartLojas');
    if (ctxLojas && lojasOrdenadas.length > 0) {
        if (charts.lojas) charts.lojas.destroy();
        charts.lojas = new Chart(ctxLojas, {
            type: 'bar',
            data: {
                labels: lojasOrdenadas.map(l => l[0].length > 20 ? l[0].substring(0, 20) + '...' : l[0]),
                datasets: [{ label: 'Perda (kg)', data: lojasOrdenadas.map(l => l[1]), backgroundColor: '#667eea' }]
            },
            options: { responsive: true }
        });
    }
    
    const produtosOrdenados = Object.entries(produtos).sort((a, b) => b[1] - a[1]).slice(0, 10);
    const ctxProdutos = document.getElementById('chartProdutos');
    if (ctxProdutos && produtosOrdenados.length > 0) {
        if (charts.produtos) charts.produtos.destroy();
        charts.produtos = new Chart(ctxProdutos, {
            type: 'bar',
            data: {
                labels: produtosOrdenados.map(p => p[0]),
                datasets: [{ label: 'Perda (kg)', data: produtosOrdenados.map(p => p[1]), backgroundColor: '#764ba2' }]
            },
            options: { responsive: true, indexAxis: 'y' }
        });
    }
    
    const motivosOrdenados = Object.entries(motivos).sort((a, b) => b[1] - a[1]).slice(0, 6);
    const ctxMotivos = document.getElementById('chartMotivos');
    if (ctxMotivos && motivosOrdenados.length > 0) {
        if (charts.motivos) charts.motivos.destroy();
        charts.motivos = new Chart(ctxMotivos, {
            type: 'pie',
            data: {
                labels: motivosOrdenados.map(m => m[0].length > 25 ? m[0].substring(0, 25) + '...' : m[0]),
                datasets: [{ data: motivosOrdenados.map(m => m[1]), backgroundColor: ['#FFB347', '#FF6B6B', '#4ECDC4', '#95A5A6', '#3498DB', '#2ECC71'] }]
            },
            options: { responsive: true, plugins: { legend: { position: 'right' } } }
        });
    }
}

function atualizarTabela() {
    const tipo = filtroTipo.value;
    const texto = filtroProduto.value.toLowerCase();
    
    let dados = dadosCompletos.filter(d => {
        if (tipo !== 'todos' && d.tipo !== tipo) return false;
        if (texto && !d.produto.toLowerCase().includes(texto) && !d.descMotivo.toLowerCase().includes(texto)) return false;
        return true;
    });
    
    dados.sort((a, b) => b.quantidade - a.quantidade);
    
    const tbody = document.getElementById('tabelaBody');
    tbody.innerHTML = '';
    
    if (dados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" style="text-align: center;">📁 Nenhum registro encontrado</td></tr>';
        document.getElementById('contadorRegistros').textContent = '';
        return;
    }
    
    dados.slice(0, 200).forEach(item => {
        const row = tbody.insertRow();
        row.insertCell(0).textContent = item.loja;
        row.insertCell(1).textContent = item.produto;
        row.insertCell(2).textContent = item.descMotivo;
        row.insertCell(3).textContent = item.tipo;
        row.insertCell(4).textContent = item.quantidade.toFixed(2);
        row.insertCell(5).textContent = `R$ ${item.valor.toFixed(2)}`;
    });
    
    document.getElementById('contadorRegistros').textContent = `📊 Mostrando ${Math.min(dados.length, 200)} de ${dados.length} registros`;
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
            dataAtualizacaoSpan.textContent = new Date(lastUpdate).toLocaleDateString('pt-BR');
        }
    }
}

loadSavedData();