import React, { useState, useEffect } from 'react';
import { 
  BarChart, Bar, PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer 
} from 'recharts';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import './index.css';

const imageUrlToBase64 = async (url) => {
    try {
        const response = await fetch(url);
        const blob = await response.blob();
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    } catch (e) {
        return null;
    }
};

const getDynamicDateRange = (periodo) => {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth(); // 0-indexed
  const pad = (n) => String(n).padStart(2, '0');

  let s = `${y}-${pad(m + 1)}-01`;
  let e = `${y}-${pad(m + 1)}-${pad(new Date(y, m + 1, 0).getDate())}`;

  if (periodo === 'Mês Anterior') {
    const prevMonthDate = new Date(y, m - 1, 1);
    const prevY = prevMonthDate.getFullYear();
    const prevM = prevMonthDate.getMonth();
    s = `${prevY}-${pad(prevM + 1)}-01`;
    e = `${prevY}-${pad(prevM + 1)}-${pad(new Date(prevY, prevM + 1, 0).getDate())}`;
  } else if (periodo === 'Últimos 3 Meses') {
    const threeMonthsAgo = new Date(y, m - 2, 1);
    s = `${threeMonthsAgo.getFullYear()}-${pad(threeMonthsAgo.getMonth() + 1)}-01`;
    e = `${y}-${pad(m + 1)}-${pad(new Date(y, m + 1, 0).getDate())}`;
  } else if (periodo === 'Ano Atual') {
    s = `${y}-01-01`;
    e = `${y}-${pad(m + 1)}-${pad(new Date(y, m + 1, 0).getDate())}`;
  }

  return { start: s, end: e };
};
const ScoreBadge = ({ score }) => {
  const getColor = (s) => s >= 8 ? 'var(--success)' : s >= 6.5 ? 'var(--info)' : s >= 4 ? 'var(--warning)' : 'var(--danger)';
  const getBgColor = (s) => {
    if (s >= 8) return 'rgba(16, 185, 129, 0.1)';
    if (s >= 6.5) return 'rgba(59, 130, 246, 0.1)';
    if (s >= 4) return 'rgba(245, 158, 11, 0.1)';
    return 'rgba(239, 68, 68, 0.1)';
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
      <div style={{ 
        width: '38px', 
        height: '38px', 
        borderRadius: '50%', 
        background: getBgColor(score),
        border: `1.5px solid ${getColor(score)}`, 
        display: 'flex', 
        alignItems: 'center', 
        justifyContent: 'center', 
        fontSize: '0.85rem', 
        fontWeight: '800', 
        color: getColor(score) 
      }}>
        {score.toFixed(1)}
      </div>
    </div>
  );
};

const KPIStatsCards = ({ kpis, compare }) => {
    const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val || 0);
    
    const stats = [
        { label: 'Faturamento Total', value: formatCurrency(kpis?.faturamento), variation: kpis?.varFat, icon: '💰' },
        { label: 'Qtd de Pedidos', value: kpis?.pedidos || 0, variation: kpis?.varPed, icon: '📦' },
        { label: 'Ticket Médio', value: formatCurrency(kpis?.ticketMedio), variation: kpis?.varTicket, icon: '🎟️' },
        { label: 'Volume de Itens', value: kpis?.itens?.toLocaleString('pt-BR') || 0, variation: null, icon: '📊' },
    ];

    return (
        <div className="kpi-panel">
            {stats.map((s, i) => (
                <div key={i} className="kpi-card-v2">
                    <div className="kpi-header">
                        <span className="kpi-icon-bg">{s.icon}</span>
                        <label>{s.label}</label>
                    </div>
                    <div className="kpi-body">
                        <strong>{s.value}</strong>
                        {compare && s.variation !== null && (
                            <span className={`kpi-variation ${parseFloat(s.variation) < 0 ? 'neg' : 'pos'}`}>
                                {parseFloat(s.variation) >= 0 ? '▲' : '▼'} {Math.abs(s.variation)}%
                            </span>
                        )}
                    </div>
                </div>
            ))}
        </div>
    );
};

const View360 = ({ client, onBack, globalFilters }) => {
    const [activeTab, setActiveTab] = useState('resumo');
    const [data, setData] = useState(null);
    const [loading, setLoading] = useState(true);
    const [onlyStock, setOnlyStock] = useState(false);
    const [preOrder, setPreOrder] = useState([]);
    const [brandFilter, setBrandFilter] = useState(null);
    const [hideAssortment, setHideAssortment] = useState(false);
    const [localFilters, setLocalFilters] = useState(() => {
        const initialLocalDates = getDynamicDateRange(globalFilters?.periodo || 'Mês Atual');
        return {
            periodo: globalFilters?.periodo || 'Mês Atual',
            start: globalFilters?.start || initialLocalDates.start,
            end: globalFilters?.end || initialLocalDates.end
        };
    });

    const togglePreOrder = (prod) => {
        setPreOrder(prev => {
            const exists = prev.find(i => i.nome === prod.nome);
            if (exists) return prev.filter(i => i.nome !== prod.nome);
            return [...prev, prod];
        });
    };

    useEffect(() => {
        setLoading(true);
        const url = `/api/clientes/${encodeURIComponent(client.cnpj)}?start=${localFilters.start}&end=${localFilters.end}`;
        fetch(url)
            .then(res => res.json())
            .then(d => { setData(d); setLoading(false); })
            .catch(e => { console.error(e); setLoading(false); });
    }, [client, localFilters.start, localFilters.end]);

    const handlePeriodChange = (p) => {
        if (p === 'Personalizado') {
            setLocalFilters(prev => ({ ...prev, periodo: p }));
            return;
        }
        const today = new Date();
        const y = today.getFullYear();
        const m = today.getMonth();
        const pad = (n) => String(n).padStart(2, '0');
        const lastDayOfMonth = (yr, mo) => new Date(yr, mo + 1, 0).getDate();

        let s, e;
        if (p === 'Mês Atual') {
            s = `${y}-${pad(m + 1)}-01`;
            e = `${y}-${pad(m + 1)}-${pad(lastDayOfMonth(y, m))}`;
        } else if (p === 'Mês Anterior') {
            const prev = new Date(y, m - 1, 1);
            s = `${prev.getFullYear()}-${pad(prev.getMonth() + 1)}-01`;
            e = `${prev.getFullYear()}-${pad(prev.getMonth() + 1)}-${pad(lastDayOfMonth(prev.getFullYear(), prev.getMonth()))}`;
        } else if (p === 'Últimos 30 Dias') {
            const d30 = new Date(today); d30.setDate(d30.getDate() - 30);
            s = `${d30.getFullYear()}-${pad(d30.getMonth() + 1)}-${pad(d30.getDate())}`;
            e = `${y}-${pad(m + 1)}-${pad(today.getDate())}`;
        } else if (p === 'Últimos 6 Meses') {
            const d6 = new Date(y, m - 5, 1);
            s = `${d6.getFullYear()}-${pad(d6.getMonth() + 1)}-01`;
            e = `${y}-${pad(m + 1)}-${pad(lastDayOfMonth(y, m))}`;
        } else if (p === 'Últimos 12 Meses') {
            const d12 = new Date(y, m - 11, 1);
            s = `${d12.getFullYear()}-${pad(d12.getMonth() + 1)}-01`;
            e = `${y}-${pad(m + 1)}-${pad(lastDayOfMonth(y, m))}`;
        } else if (p === 'Ano Atual') {
            s = `${y}-01-01`;
            e = `${y}-${pad(m + 1)}-${pad(lastDayOfMonth(y, m))}`;
        } else {
            const range = getDynamicDateRange(p);
            s = range.start; e = range.end;
        }
        setLocalFilters(prev => ({ ...prev, periodo: p, start: s, end: e }));
    };

    const oppValue = (data?.marketGap || [])
        .filter(p => p.saldo > 0)
        .reduce((acc, p) => acc + (p.total_geral || 0), 0);

    const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(val || 0);
    const formatDate = (dateStr) => {
        if (!dateStr || dateStr === '-') return '-';
        try {
            return dateStr.split('-').reverse().join('/');
        } catch (e) { return dateStr; }
    };
    const toTitleCase = (str) => {
        if (!str) return '';
        return str.toLowerCase().split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' ');
    };
    const SUNNY_COLORS = ['#003087', '#FFD700', '#0056D2', '#FFED80', '#001D54', '#85A5FF', '#ADC6FF'];

    const getSugestaoList = () => {
        const baseDate = new Date();
        return (data.topProdutos || [])
            .filter(p => p.saldo > 0)
            .map(p => {
                const diasUltimaCompra = p.ultima_data ? Math.floor((baseDate - new Date(p.ultima_data)) / (1000*60*60*24)) : 0;
                const diasDesdePrimeira = p.primeira_data ? Math.floor((baseDate - new Date(p.primeira_data)) / (1000 * 60 * 60 * 24)) : 30;
                
                let cicloSku = 30;
                if (p.total_pedidos_sku > 1) {
                    cicloSku = Math.max(1, Math.floor(diasDesdePrimeira / p.total_pedidos_sku));
                }
                
                const mediaCompra = p.total_pedidos_sku > 0 ? (p.qtd / p.total_pedidos_sku) : p.qtd;
                const packNum = parseInt(p.pack) || 1;
                let sugestaoUn = Math.ceil(mediaCompra / packNum) * packNum;
                if (sugestaoUn > p.saldo) sugestaoUn = Math.floor(p.saldo / packNum) * packNum;

                let statusGiro = { cor: '#166534', bg: '#DCFCE7', label: 'Estoque Saudável' };
                if (diasUltimaCompra > cicloSku) {
                    statusGiro = { cor: '#991B1B', bg: '#FEE2E2', label: 'Provável Ruptura' };
                } else if (diasUltimaCompra >= (cicloSku * 0.7)) {
                    statusGiro = { cor: '#92400E', bg: '#FEF3C7', label: 'Compra Iminente' };
                }

                return { ...p, sugestaoUn, packNum, mediaCompra, cicloSku, diasUltimaCompra, statusGiro };
            })
            .filter(p => p.sugestaoUn > 0);
    };

    const exportSuggestionExcel = () => {
        const list = getSugestaoList();
        if(!list.length) return alert('Nenhum produto atende aos critérios para reposição.');
        
        const rows = list.map(p => {
            const percDesconto = p.pv > 0 ? (((p.preco_ultima / p.pv) - 1) * 100) : 0;
            return {
                'Status de Giro': p.statusGiro.label,
                'Cód. Produto': p.produto_id,
                'Descrição': p.nome,
                'Marca': p.marca,
                'Quantidade Sugerida': p.sugestaoUn,
                'Packs': p.sugestaoUn / p.packNum,
                'Ciclo Médio (Dias)': p.cicloSku,
                'Última Compra (Dias)': p.diasUltimaCompra,
                'Demanda Média': parseFloat(p.mediaCompra.toFixed(1)),
                'Preço de Venda (PV)': p.pv,
                'Ultima Compra (R$)': p.preco_ultima || 0,
                'Desconto(%)': percDesconto.toFixed(2) + '%'
            };
        });
        
        const ws = XLSX.utils.json_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Reposição Preditiva");
        XLSX.writeFile(wb, `Reposicao_${client.cnpj}.xlsx`);
    };

    const exportSuggestionPDF = async () => {
        // Filtrar apenas itens com giro atrasado ou em alerta (exclui Estoque Saudável)
        const list = getSugestaoList().filter(p => p.statusGiro.label !== 'Estoque Saudável');
        if(!list.length) return alert('Nenhum produto com giro atrasado ou em alerta encontrado.');
        
        const doc = new jsPDF();
        
        doc.setFontSize(16);
        doc.setTextColor(0, 48, 135);
        doc.text("Inteligência de Reposição Preditiva - Sunny", 14, 20);
        
        doc.setFontSize(10);
        doc.setTextColor(50);
        doc.text(`Cliente: ${client.cliente} (CNPJ: ${client.cnpj})`, 14, 28);
        doc.text(`Vendedor(a): ${client.representante || 'Não Informado'}`, 14, 34);
        doc.text(`Data da Análise: ${new Date().toLocaleDateString('pt-BR')}`, 14, 40);
        
        const tableData = [];
        const images = {};

        for (let i = 0; i < list.length; i++) {
            const p = list[i];
            tableData.push([
                '', // Foto col
                p.statusGiro.label,
                `${p.produto_id} - ${p.nome}`,
                `${p.diasUltimaCompra} dias`,
                `${p.cicloSku} dias`,
                `${p.sugestaoUn} un (${p.sugestaoUn / p.packNum} cx)`
            ]);
            
            if (p.image_url) {
                const base64 = await imageUrlToBase64(p.image_url);
                if (base64) images[i] = base64;
            }
        }
        
        autoTable(doc, {
            startY: 48,
            head: [['Foto', 'Status de Giro', 'Produto', 'Última Compra', 'Ciclo Médio', 'Sugestão']],
            body: tableData,
            headStyles: { fillColor: [0, 48, 135], textColor: [255, 255, 255] },
            alternateRowStyles: { fillColor: [240, 248, 255] },
            styles: { fontSize: 8, minCellHeight: 15, verticalAlign: 'middle' },
            columnStyles: {
                0: { cellWidth: 15 },
                1: { fontStyle: 'bold' },
            },
            didDrawCell: (data) => {
                if (data.section === 'body' && data.column.index === 0 && images[data.row.index]) {
                    doc.addImage(images[data.row.index], 'JPEG', data.cell.x + 2, data.cell.y + 2, 11, 11);
                }
            }
        });
        
        doc.save(`Reposicao_${client.cnpj}.pdf`);
    };

    const exportDossiePDF = async () => {
        // Respeita filtro de sortimento ativo na tela
        const gapList = (data.marketGap || [])
            .filter(p => !hideAssortment || p.sortimento === 'Não');
        const doc = new jsPDF();
        
        doc.setFontSize(18);
        doc.setTextColor(0, 48, 135);
        doc.text("Dossiê Estratégico de Vendas - Sunny", 14, 20);
        
        doc.setFontSize(10);
        doc.setTextColor(50);
        doc.text(`Cliente: ${client.cliente} (CNPJ: ${client.cnpj})`, 14, 28);
        doc.text(`Perfil: ${client.perfil || 'Não Classificado'} | UF: ${client.uf}`, 14, 34);
        doc.text(`Data da Análise: ${new Date().toLocaleDateString('pt-BR')}`, 14, 40);
        doc.setFontSize(8);
        doc.setTextColor(100);
        doc.text(`Análise baseada no desempenho nacional de: ${localFilters.periodo}`, 14, 45);

        doc.setFontSize(12);
        doc.setTextColor(0, 48, 135);
        doc.text("Resumo de Performance", 14, 52);

        doc.setFontSize(9);
        doc.setTextColor(50);
        doc.text(`Faturamento no Período: ${formatCurrency(data.faturamentoPeriodo)}`, 14, 60);
        doc.text(`Benchmark (${client.uf} - ${client.perfil}): ${formatCurrency(data.benchmark?.valor)}`, 14, 66);
        doc.text(`Oportunidade de Gap (vs Benchmark): ${formatCurrency(Math.max(0, (data.benchmark?.valor || 0) - data.faturamentoPeriodo))}`, 14, 72);
        doc.text(`Cobertura de Mix Nacional: ${(data.coberturaMix || 0).toFixed(1)}%`, 14, 78);

        const sumAtual = (data.timeline || []).reduce((a, b) => a + (b.atual || 0), 0);
        const sumAnt = (data.timeline || []).reduce((a, b) => a + (b.anterior || 0), 0);
        if (sumAnt > 0) {
            const varYoy = ((sumAtual - sumAnt) / sumAnt) * 100;
            doc.text(`Evolução YoY: ${varYoy > 0 ? '+' : ''}${varYoy.toFixed(1)}%`, 14, 84);
        }

        // Todas as oportunidades (respeitando filtros ativos)
        if (gapList.length > 0) {
            doc.setFontSize(12);
            doc.setTextColor(0, 48, 135);
            doc.text(`Oportunidades de Expansão de Mix (${gapList.length} produtos)`, 14, 100);

            const gapData = [];
            const images = {};

            for (let i = 0; i < gapList.length; i++) {
                const p = gapList[i];
                gapData.push([
                    '', // Foto col
                    p.rank_nacional <= 10 ? '⭐ Top 10' : `#${p.rank_nacional}`,
                    `${p.produto_id} - ${p.nome}`,
                    p.marca || '-',
                    formatCurrency(p.pv),
                    String(p.saldo || 0)
                ]);

                if (p.image_url) {
                    const base64 = await imageUrlToBase64(p.image_url);
                    if (base64) images[i] = base64;
                }
            }

            autoTable(doc, {
                startY: 105,
                head: [['Foto', 'Ranking', 'Produto', 'Marca', 'PV Sunny', 'Saldo']],
                body: gapData,
                headStyles: { fillColor: [0, 48, 135], textColor: [255, 255, 255] },
                alternateRowStyles: { fillColor: [245, 247, 250] },
                styles: { fontSize: 7.5, cellPadding: 3, minCellHeight: 15, verticalAlign: 'middle' },
                columnStyles: {
                    0: { cellWidth: 15 },
                    1: { fontStyle: 'bold', halign: 'center', cellWidth: 20 },
                    4: { halign: 'right' },
                    5: { halign: 'center' }
                },
                didDrawCell: (data) => {
                    if (data.section === 'body' && data.column.index === 0 && images[data.row.index]) {
                        doc.addImage(images[data.row.index], 'JPEG', data.cell.x + 2, data.cell.y + 2, 11, 11);
                    }
                },
                didParseCell: (hookData) => {
                    // Destaca as linhas Top 10 com fundo amarelo claro
                    if (hookData.section === 'body' && gapData[hookData.row.index]?.[1]?.startsWith('⭐')) {
                        hookData.cell.styles.fillColor = [255, 248, 196];
                        hookData.cell.styles.fontStyle = 'bold';
                    }
                }
            });
        } else {
            doc.setFontSize(10);
            doc.setTextColor(100);
            doc.text('Nenhuma oportunidade encontrada com os filtros atuais.', 14, 100);
        }
        
        doc.save(`Dossie_${client.cnpj}.pdf`);
    };



    if (loading) return (
        <div className="v360-overlay fade-in">
            <div style={{display:'flex', alignItems:'center', justifyContent:'center', height:'100%', color:'white', fontSize:'1.2rem', flexDirection:'column', gap:'16px'}}>
                <div className="loader-spinner"></div>
                Carregando Inteligência Comercial 360°...
            </div>
        </div>
    );

    if (!data) return (
        <div className="v360-overlay fade-in">
             <div style={{display:'flex', alignItems:'center', justifyContent:'center', height:'100%', color:'white', fontSize:'1.2rem', flexDirection:'column', gap:'16px'}}>
                Erro ao carregar dados do ERP. 
                <button className="v360-back" onClick={onBack}>Voltar</button>
            </div>
        </div>
    );

    const totalFat = data.brandShare?.reduce((acc, b) => acc + b.total, 0) || 1;

    return (
        <div className="v360-overlay fade-in">
            <div className="v360-header">
                <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'16px'}}>
                    <button className="v360-back" onClick={onBack}>← Voltar para listagem</button>
                    <div className="filter-select-v2" style={{background:'rgba(255,255,255,0.1)', color:'white', border:'1px solid rgba(255,255,255,0.2)', padding:'4px 12px'}}>
                        <span className="select-icon">📅</span>
                        <select 
                            value={localFilters.periodo} 
                            onChange={(e) => handlePeriodChange(e.target.value)}
                            style={{background:'transparent', color:'white', border:0}}
                        >
                            <option value="Mês Atual" style={{color:'black'}}>Mês Atual</option>
                            <option value="Ano Atual" style={{color:'black'}}>Ano Atual (2026)</option>
                            <option value="Últimos 6 Meses" style={{color:'black'}}>Últimos 6 Meses</option>
                            <option value="Últimos 12 Meses" style={{color:'black'}}>Últimos 12 Meses</option>
                        </select>
                    </div>
                </div>
                <div style={{display:'flex', justifyContent:'space-between', alignItems:'flex-end', width:'100%'}}>
                    <div>
                        <h1 style={{margin:'0 0 8px 0', fontSize:'2.2rem'}}>{client.cliente}</h1>
                        <div style={{display:'flex', gap:'20px', fontSize:'0.85rem', opacity:0.8, fontWeight:'500'}}>
                            <span>ID: {client.id}</span>
                            <span>CNPJ: {client.cnpj}</span>
                            <span>UF: {client.uf}</span>
                            <span style={{color:'var(--sunny-yellow)'}}>📊 Analisando: {localFilters.periodo}</span>
                        </div>
                    </div>
                    <div className="score-badge">Score: {client.score?.toFixed(1) || '0.0'} - {client.scoreLabel}</div>
                </div>
            </div>

            <div className="v360-tabs">
                <button className={`v360-tab ${activeTab === 'resumo' ? 'active' : ''}`} onClick={() => setActiveTab('resumo')}>Resumo Estratégico</button>
                <button className={`v360-tab ${activeTab === 'produtos' ? 'active' : ''}`} onClick={() => setActiveTab('produtos')}>Mix de Produtos {preOrder.length > 0 && <span style={{background:'var(--sunny-yellow)', color:'black', padding:'2px 6px', borderRadius:'10px', fontSize:'0.7rem', marginLeft:'6px', fontWeight:'800'}}>{preOrder.length}</span>}</button>
                <button className={`v360-tab ${activeTab === 'oportunidades' ? 'active' : ''}`} onClick={() => setActiveTab('oportunidades')}>Oportunidades</button>
                <button className={`v360-tab ${activeTab === 'pedidos' ? 'active' : ''}`} onClick={() => setActiveTab('pedidos')}>Histórico de Pedidos</button>
            </div>

            <div className="v360-content">
                {activeTab === 'resumo' && (
                    <div className="fade-in">
                        <div style={{display:'grid', gridTemplateColumns:'2fr 1fr', gap:'24px', marginBottom:'24px'}}>
                            <div className="records-grid" style={{margin:0, gridTemplateColumns:'1fr 1fr'}}>
                                <div className="record-card-360">
                                    <label>🕒 Último Pedido</label>
                                    <strong>{formatCurrency(data.records?.ultimoPedido?.total)}</strong>
                                    <span>Realizado em: {formatDate(data.records?.ultimoPedido?.emissao)}</span>
                                </div>
                                <div className="record-card-360 gold">
                                    <label>🏆 Maior Pedido Histórico</label>
                                    <strong>{formatCurrency(data.records?.maiorPedido?.total)}</strong>
                                    <span>Data do Recorde: {formatDate(data.records?.maiorPedido?.emissao)}</span>
                                </div>

                                {/* Gauge de Cobertura de Mix */}
                                <div className="record-card-360" style={{background:'#F0FDF4', border:'1px solid #BBF7D0', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', padding:'16px 8px'}}>
                                    <label style={{color:'#166534', marginBottom:'4px'}}>📊 Cobertura de Mix Nacional</label>
                                    {(() => {
                                        const pct = Math.min(100, data.coberturaMix || 0);
                                        const r = 52, cx = 70, cy = 70;
                                        const startAngle = 210, endAngle = 510;
                                        const toRad = (d) => (d * Math.PI) / 180;
                                        const arcX = (angle) => cx + r * Math.cos(toRad(angle - 90));
                                        const arcY = (angle) => cy + r * Math.sin(toRad(angle - 90));
                                        const sweepTotal = endAngle - startAngle;
                                        const sweepFill = (pct / 100) * sweepTotal;
                                        const endFill = startAngle + sweepFill;
                                        const largeArc = sweepFill > 180 ? 1 : 0;
                                        const col = pct >= 60 ? '#16A34A' : pct >= 30 ? '#CA8A04' : '#DC2626';
                                        return (
                                            <svg width="140" height="90" viewBox="0 0 140 90">
                                                <path d={`M ${arcX(startAngle)} ${arcY(startAngle)} A ${r} ${r} 0 1 1 ${arcX(endAngle)} ${arcY(endAngle)}`} fill="none" stroke="#DCFCE7" strokeWidth="10" strokeLinecap="round" />
                                                {pct > 0 && <path d={`M ${arcX(startAngle)} ${arcY(startAngle)} A ${r} ${r} 0 ${largeArc} 1 ${arcX(endFill)} ${arcY(endFill)}`} fill="none" stroke={col} strokeWidth="10" strokeLinecap="round" />}
                                                <text x={cx} y={cy - 4} textAnchor="middle" fontSize="18" fontWeight="900" fill={col}>{pct.toFixed(1)}%</text>
                                                <text x={cx} y={cy + 12} textAnchor="middle" fontSize="8" fill="#6B7280">do portfólio ativo</text>
                                            </svg>
                                        );
                                    })()}
                                </div>

                                {/* Benchmark Expandido */}
                                {(() => {
                                    const fat = data.faturamentoPeriodo || 0;
                                    const bench = data.benchmark?.valor || 0;
                                    const gap = bench - fat;
                                    const isAbove = fat >= bench;
                                    const pctFat = bench > 0 ? Math.min(100, (fat / bench) * 100) : 100;
                                    return (
                                        <div className="record-card-360" style={{background: isAbove ? '#F0FDF4' : '#FFFBEB', border:`1px solid ${isAbove ? '#BBF7D0' : '#FDE68A'}`, gridColumn: '1', padding:'16px'}}>
                                            <label style={{color: isAbove ? '#166534' : '#92400E'}}>🎯 Benchmark — {data.benchmark?.uf} ({data.benchmark?.perfil})</label>
                                            <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:'10px', marginTop:'10px'}}>
                                                <div>
                                                    <small style={{display:'block', fontSize:'0.7rem', color:'#6B7280'}}>Sua Performance</small>
                                                    <strong style={{fontSize:'1.15rem', color:'#1E3A8A'}}>{formatCurrency(fat)}</strong>
                                                </div>
                                                <div>
                                                    <small style={{display:'block', fontSize:'0.7rem', color:'#6B7280'}}>Média da Região</small>
                                                    <strong style={{fontSize:'1.15rem', color: isAbove ? '#166534' : '#92400E'}}>{formatCurrency(bench)}</strong>
                                                </div>
                                            </div>
                                            <div style={{marginTop:'10px'}}>
                                                <div style={{display:'flex', justifyContent:'space-between', fontSize:'0.7rem', color:'#6B7280', marginBottom:'4px'}}>
                                                    <span>Progresso vs Benchmark</span>
                                                    <span style={{fontWeight:'700', color: isAbove ? '#166534' : '#92400E'}}>{isAbove ? '⬆ Acima' : `Gap: ${formatCurrency(gap)}`}</span>
                                                </div>
                                                <div style={{width:'100%', height:'8px', background:'#E5E7EB', borderRadius:'4px', overflow:'hidden'}}>
                                                    <div style={{width:`${pctFat}%`, height:'100%', background: isAbove ? '#16A34A' : '#F59E0B', transition:'width 0.8s ease'}}></div>
                                                </div>
                                            </div>
                                        </div>
                                    );
                                })()}
                            </div>

                            <div className="v360-card-custom" style={{margin:0, background:'#F0F9FF', border:'1px solid #BEE3F8'}}>
                                <h4 style={{margin:'0 0 12px 0', color:'var(--sunny-blue)', fontSize:'0.9rem'}}>💎 Fatos Históricos (Vitalício)</h4>
                                <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px'}}>
                                    <div>
                                        <small style={{display:'block', color:'#718096', fontSize:'0.7rem'}}>Fat. Total Acumulado</small>
                                        <strong style={{fontSize:'1.1rem', color:'var(--sunny-blue)'}}>{formatCurrency(data.lifetime?.total)}</strong>
                                    </div>
                                    <div>
                                        <small style={{display:'block', color:'#718096', fontSize:'0.7rem'}}>1ª Compra Realizada</small>
                                        <strong style={{fontSize:'1.1rem'}}>{formatDate(data.lifetime?.primeira_compra)}</strong>
                                    </div>
                                    <div>
                                        <small style={{display:'block', color:'#718096', fontSize:'0.7rem'}}>Total de Pedidos</small>
                                        <strong style={{fontSize:'1.1rem'}}>{data.lifetime?.total_pedidos} notas</strong>
                                    </div>
                                    <div>
                                        <small style={{display:'block', color:'#718096', fontSize:'0.7rem'}}>Fat. no Período Selecionado</small>
                                        <strong style={{fontSize:'1.1rem', color:'#166534'}}>{formatCurrency(data.faturamentoPeriodo)}</strong>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div className="charts-grid-dashboard">
                            <div className="v360-card-custom chart-box-360" style={{padding:'24px', overflow:'hidden', display:'flex', flexDirection:'column'}}>
                                <h3 style={{marginBottom:'16px', fontSize:'1rem', color:'var(--sunny-blue)', fontWeight:'800'}}>🏷️ Share por Marca ({localFilters.periodo})</h3>
                                <div style={{flex:1, overflowY:'auto', borderRadius:'8px', border:'1px solid #E2E8F0'}}>
                                    <table className="analy-table-minimal" style={{width:'100%', fontSize:'0.8rem', borderCollapse:'collapse'}}>
                                        <thead style={{background:'#F7FAFC', position:'sticky', top:0}}>
                                            <tr>
                                                <th style={{padding:'10px', textAlign:'left'}}>Marca</th>
                                                <th style={{padding:'10px', textAlign:'right'}}>Faturamento</th>
                                                <th style={{padding:'10px', textAlign:'right', width:'120px'}}>Part. %</th>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {data.brandShare?.map((b, i) => {
                                                const perc = (b.total * 100 / totalFat);
                                                return (
                                                    <tr key={i} style={{cursor:'pointer', transition:'0.2s'}} className="hover-brand-row">
                                                        <td 
                                                            style={{padding:'12px 10px', fontWeight:'700', color:'var(--sunny-blue)'}}
                                                            onClick={() => {
                                                                setBrandFilter(b.marca);
                                                                setActiveTab('produtos');
                                                            }}
                                                        >
                                                            {toTitleCase(b.marca || 'SEM MARCA')}
                                                        </td>
                                                        <td style={{padding:'12px 10px', textAlign:'right', fontWeight:'600'}}>{formatCurrency(b.total)}</td>
                                                        <td style={{padding:'12px 10px', textAlign:'right'}}>
                                                            <div style={{display:'flex', alignItems:'center', justifyContent:'flex-end', gap:'8px'}}>
                                                                <span style={{fontSize:'0.7rem', fontWeight:'800'}}>{perc.toFixed(1)}%</span>
                                                                <div style={{width:'60px', height:'6px', background:'#EDF2F7', borderRadius:'10px', overflow:'hidden'}}>
                                                                    <div style={{width:`${perc}%`, height:'100%', background:'var(--sunny-blue)'}}></div>
                                                                </div>
                                                            </div>
                                                        </td>
                                                    </tr>
                                                );
                                            })}
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                            <div className="v360-card-custom chart-box-360" style={{position:'relative'}}>
                                <div style={{display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:'16px'}}>
                                    <h3 style={{margin:0}}>Timeline YoY (Faturamento)</h3>
                                    {(() => {
                                        const sumAtual = (data.timeline || []).reduce((a, b) => a + (b.atual || 0), 0);
                                        const sumAnt = (data.timeline || []).reduce((a, b) => a + (b.anterior || 0), 0);
                                        if (sumAnt === 0) return null;
                                        const varYoy = ((sumAtual - sumAnt) / sumAnt) * 100;
                                        const isPos = varYoy > 0;
                                        return (
                                            <div style={{background: isPos ? '#F0FDF4' : '#FEF2F2', color: isPos ? '#166534' : '#991B1B', padding:'4px 12px', borderRadius:'20px', fontWeight:'800', fontSize:'0.85rem', border:`1px solid ${isPos ? '#BBF7D0' : '#FECACA'}`}}>
                                                Evolução YoY: {isPos ? '+' : ''}{varYoy.toFixed(1)}%
                                            </div>
                                        );
                                    })()}
                                </div>
                                <div style={{height: 300}}>
                                    <ResponsiveContainer width="100%" height="100%">
                                        <BarChart data={data.timeline || []}>
                                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E2E8F0" />
                                            <XAxis dataKey="mes" fontSize={12} tickFormatter={(m) => ['Jan','Fev','Mar','Abr','Mai','Jun','Jul','Ago','Set','Out','Nov','Dez'][parseInt(m)-1] || m} />
                                            <YAxis hide />
                                            <Tooltip formatter={(v) => formatCurrency(v)} />
                                            <Legend />
                                            <Bar name="2026" dataKey="atual" fill="#003087" radius={[4, 4, 0, 0]} />
                                            <Bar name="2025" dataKey="anterior" fill="#CBD5E1" radius={[4, 4, 0, 0]} />
                                        </BarChart>
                                    </ResponsiveContainer>
                                </div>
                            </div>
                        </div>
                    </div>
                )}

                {activeTab === 'produtos' && (
                    <div className="v360-card-custom fade-in" style={{overflow:'hidden'}}>
                        <div style={{display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:'24px', flexWrap:'wrap', gap:'20px'}}>
                            <div>
                                <h3 style={{margin:0, fontSize:'1.2rem'}}>🛍️ Inteligência de Mix & Oportunidades</h3>
                                <p style={{fontSize:'0.85rem', color:'#4A5568', margin:'4px 0 0 0', fontWeight:'500'}}>Visão detalhada de faturamento, estoque e precificação estratégica.</p>
                            </div>
                            
                            <div style={{display:'flex', alignItems:'center', gap:'12px', flexWrap:'wrap'}}>
                                <div className="export-dropdown" style={{position:'relative'}}>
                                    <button 
                                        style={{display:'flex', alignItems:'center', gap:'10px', cursor:'pointer', fontSize:'0.9rem', fontWeight:'800', color:'var(--sunny-blue)', background:'#EBF8FF', padding:'10px 16px', borderRadius:'10px', border:'1px solid #BEE3F8', transition:'0.2s', boxShadow:'0 2px 4px rgba(0,0,0,0.05)'}} 
                                        onClick={(e) => {
                                            const menu = document.getElementById('exportMenu');
                                            menu.style.display = menu.style.display === 'block' ? 'none' : 'block';
                                        }}>
                                        <span>📥</span> Baixar Relatório ▽
                                    </button>
                                    <div id="exportMenu" className="dropdown-content" style={{display:'none', position:'absolute', right:0, background:'white', border:'1px solid #E2E8F0', borderRadius:'8px', marginTop:'5px', zIndex:100, width:'200px', boxShadow:'0 4px 12px rgba(0,0,0,0.1)'}}>
                                        <button onClick={() => { document.getElementById('exportMenu').style.display='none'; exportSuggestionExcel(); }} style={{display:'flex', alignItems:'center', gap:'8px', width:'100%', padding:'12px 15px', textAlign:'left', border:'none', background:'transparent', color:'#1A202C', cursor:'pointer', borderBottom:'1px solid #E2E8F0', fontWeight:'700'}}>
                                            <span style={{color:'var(--success)', fontSize:'1.2rem'}}>📊</span> Excel (.xlsx)
                                        </button>
                                        <button onClick={() => { document.getElementById('exportMenu').style.display='none'; exportSuggestionPDF(); }} style={{display:'flex', alignItems:'center', gap:'8px', width:'100%', padding:'12px 15px', textAlign:'left', border:'none', background:'transparent', color:'#1A202C', cursor:'pointer', fontWeight:'700'}}>
                                            <span style={{color:'var(--danger)', fontSize:'1.2rem'}}>📄</span> Relatório PDF
                                        </button>
                                    </div>
                                </div>
                                <label style={{display:'flex', alignItems:'center', gap:'10px', cursor:'pointer', fontSize:'0.9rem', fontWeight:'800', color:'var(--sunny-blue)', background:'#EBF8FF', padding:'10px 16px', borderRadius:'10px', border:'1px solid #BEE3F8', transition:'0.2s', boxShadow:'0 2px 4px rgba(0,0,0,0.05)'}}>
                                    <input 
                                        type="checkbox" 
                                        checked={onlyStock} 
                                        onChange={(e) => setOnlyStock(e.target.checked)}
                                        style={{width:'18px', height:'18px', cursor:'pointer'}}
                                    />
                                    Pronta Entrega
                                </label>
                            </div>
                        </div>

                        {brandFilter && (
                            <div style={{marginBottom:'20px', display:'flex', alignItems:'center', gap:'12px', background:'var(--sunny-blue)', color:'white', padding:'8px 16px', borderRadius:'10px', alignSelf:'flex-start', fontSize:'0.8rem'}}>
                                <span>📍 Filtrando por Marca: <strong>{brandFilter}</strong></span>
                                <button 
                                    style={{background:'rgba(255,255,255,0.2)', border:0, color:'white', padding:'2px 8px', borderRadius:'4px', cursor:'pointer', fontSize:'0.7rem', fontWeight:'800'}}
                                    onClick={() => setBrandFilter(null)}
                                >
                                    LIMPAR FILTRO [X]
                                </button>
                            </div>
                        )}

                        <div style={{overflowX:'auto', borderRadius:'12px', border:'1px solid #E2E8F0'}}>
                            <table className="analy-table" style={{minWidth:'1200px', borderCollapse:'collapse'}}>
                                <thead style={{background:'#F8FAFC'}}>
                                    <tr>
                                        <th style={{padding:'16px', textAlign:'center', width:'60px'}}>Foto</th>
                                        <th style={{padding:'16px', textAlign:'left', minWidth:'280px'}}>Produto / SKUs</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Status / Saldo</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Ciclo / Compra</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Sugestão (Pack)</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>PV Sunny</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Sug. PDV</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Preço Médio</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Última Compra</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Desconto</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Status de Giro</th>
                                        <th style={{padding:'16px', textAlign:'right'}}>Faturamento</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Ação</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {(data.topProdutos || [])
                                        .filter(p => !onlyStock || p.saldo > 0)
                                        .filter(p => !brandFilter || p.marca === brandFilter)
                                        .sort((a, b) => {
                                            if (onlyStock) {
                                                if (b.total !== a.total) return b.total - a.total;
                                                return b.saldo - a.saldo;
                                            }
                                            return b.total - a.total;
                                        })
                                        .map((p, i) => {
                                            const diff = (p.preco_ultima || 0) - (p.pv || 0);
                                            const isHigher = diff > 0.05;
                                            const isLower = diff < -0.05;
                                            const isSelected = preOrder.find(item => item.nome === p.nome);
                                            const percDesconto = p.pv > 0 ? (((p.preco_ultima / p.pv) - 1) * 100) : 0;

                                            // Cálculo da Sugestão (Pack) e Ciclo
                                            const diasUltimaCompra = p.ultima_data ? Math.floor((new Date('2026-04-24') - new Date(p.ultima_data)) / (1000*60*60*24)) : 0;
                                            const diasDesdePrimeira = p.primeira_data ? Math.floor((new Date('2026-04-24') - new Date(p.primeira_data)) / (1000 * 60 * 60 * 24)) : 30;
                                            let cicloSku = 30;
                                            if (p.total_pedidos_sku > 1) {
                                                cicloSku = Math.max(1, Math.floor(diasDesdePrimeira / p.total_pedidos_sku));
                                            }
                                            
                                            const mediaCompra = p.total_pedidos_sku > 0 ? (p.qtd / p.total_pedidos_sku) : p.qtd;
                                            const packNum = parseInt(p.pack) || 1;
                                            let sugestaoUn = Math.ceil(mediaCompra / packNum) * packNum;
                                            let isLimited = false;

                                            if (sugestaoUn > p.saldo) {
                                                sugestaoUn = Math.floor(p.saldo / packNum) * packNum;
                                                isLimited = true;
                                            }

                                            let statusGiro = { cor: '#166534', bg: '#DCFCE7', label: 'Estoque Saudável' };
                                            if (diasUltimaCompra > cicloSku) {
                                                statusGiro = { cor: '#991B1B', bg: '#FEE2E2', label: 'Provável Ruptura' };
                                            } else if (diasUltimaCompra >= (cicloSku * 0.7)) {
                                                statusGiro = { cor: '#92400E', bg: '#FEF3C7', label: 'Compra Iminente' };
                                            }

                                            return (
                                                <tr key={i} style={{
                                                    background: isSelected ? '#F0F9FF' : (i % 2 === 0 ? '#FFFFFF' : '#FBFDFF'),
                                                    borderBottom:'1px solid #EDF2F7',
                                                    transition:'0.15s'
                                                }}>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div className="product-thumb-wrap">
                                                            {p.image_url ? (
                                                                <img src={p.image_url} alt={p.nome} className="product-thumb" loading="lazy" onError={(e) => {e.target.onerror = null; e.target.style.display='none'; e.target.parentElement.innerHTML = '<div class="img-placeholder">🚫</div>'}} />
                                                            ) : (
                                                                <div className="img-placeholder">🖼️</div>
                                                            )}
                                                        </div>
                                                    </td>
                                                    <td style={{padding:'16px'}}>
                                                        <div style={{display:'flex', alignItems:'baseline', gap:'10px'}}>
                                                            <span style={{fontWeight:'900', color:'var(--sunny-blue)', fontSize:'1rem'}}>{p.produto_id}</span>
                                                            <span style={{fontWeight:'800', fontSize:'0.9rem', color:'#1A202C'}}>{p.nome}</span>
                                                        </div>
                                                        <div style={{fontSize:'0.8rem', color:'#4A5568', marginTop:'4px', fontWeight:'600'}}>
                                                            {p.marca} {p.pack ? `• Pack: ${p.pack}` : ''} {p.sortimento ? `• Sort: ${p.sortimento}` : ''}
                                                        </div>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div style={{fontWeight:'900', fontSize:'1.1rem', marginBottom:'6px', color: p.saldo > 10 ? '#166534' : (p.saldo > 0 ? '#CA8A04' : '#991B1B')}}>
                                                            {p.saldo || 0} un
                                                        </div>
                                                        {p.saldo > 50 && (
                                                            <span style={{background:'var(--sunny-yellow)', color:'black', fontSize:'0.7rem', fontWeight:'900', padding:'4px 10px', borderRadius:'6px', display:'inline-block', boxShadow:'0 2px 4px rgba(0,0,0,0.1)'}}>
                                                                ⚡ PRONTA ENTREGA
                                                            </span>
                                                        )}
                                                        {p.saldo <= 0 && p.previsao && (
                                                            <span style={{background:'#E9D8FD', color:'#553C9A', fontSize:'0.7rem', fontWeight:'900', padding:'4px 10px', borderRadius:'6px', display:'inline-block', border:'1px solid #D6BCFA'}}>
                                                                🚚 PREV: {p.previsao}
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div style={{color:'#1A202C', fontSize:'0.85rem', fontWeight:'700'}}>Ciclo: {cicloSku}d</div>
                                                        <div style={{color:'#718096', fontSize:'0.75rem'}}>Compra/med: {mediaCompra.toFixed(0)}u</div>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        {sugestaoUn > 0 ? (
                                                            <div style={{display:'flex', flexDirection:'column', gap:'2px', alignItems:'center'}}>
                                                                <strong style={{color:'var(--sunny-blue)', fontSize:'1rem'}}>{sugestaoUn} un</strong>
                                                                <small style={{color:'#4A5568', fontSize:'0.75rem'}}>({sugestaoUn / packNum} packs)</small>
                                                                {isLimited && <span style={{background:'#FED7D7', color:'#9B2C2C', padding:'2px 6px', borderRadius:'4px', fontSize:'0.65rem', fontWeight:'700', marginTop:'2px'}}>Estoque Limitado</span>}
                                                            </div>
                                                        ) : (
                                                            <span style={{color:'#CBD5E1'}}>-</span>
                                                        )}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center', fontWeight:'800', color:'var(--sunny-blue)', fontSize:'0.95rem'}}>

                                                        {formatCurrency(p.pv)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center', fontWeight:'700', color:'#4A5568', fontSize:'0.9rem'}}>
                                                        {formatCurrency(p.pdv)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center', fontWeight:'700', color:'#1A202C', fontSize:'0.9rem'}}>
                                                        {formatCurrency(p.preco_medio)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div style={{fontWeight:'800', color:'#4A5568', fontSize:'0.9rem', display:'flex', alignItems:'center', justifyContent:'center', gap:'6px'}}>
                                                            {formatCurrency(p.preco_ultima)}
                                                            {isHigher && <span style={{color:'#E53E3E', fontWeight:'900', fontSize:'1.1rem'}} title="Mais caro que o PV">↑</span>}
                                                            {isLower && <span style={{color:'#38A169', fontWeight:'900', fontSize:'1.1rem'}} title="Mais barato que o PV">↓</span>}
                                                        </div>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <span style={{
                                                            fontWeight:'800', 
                                                            fontSize:'0.85rem',
                                                            color: percDesconto < 0 ? '#E53E3E' : (percDesconto > 0 ? 'var(--sunny-blue)' : '#718096')
                                                        }}>
                                                            {percDesconto.toFixed(2)}%
                                                        </span>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <span style={{
                                                            background: statusGiro.bg, 
                                                            color: statusGiro.cor, 
                                                            padding:'4px 10px', 
                                                            borderRadius:'8px', 
                                                            fontSize:'0.75rem', 
                                                            fontWeight:'800',
                                                            display:'inline-block',
                                                            border: `1px solid ${statusGiro.cor}40`
                                                        }}>
                                                            {statusGiro.label}<br/>
                                                            <small style={{opacity:0.8}}>{diasUltimaCompra} dias atrás</small>
                                                        </span>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'right', fontWeight:'900', color:'var(--sunny-blue)', fontSize:'1rem'}}>
                                                        {formatCurrency(p.total)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <button 
                                                            onClick={() => togglePreOrder(p)}
                                                            style={{
                                                                width:'40px', height:'40px', borderRadius:'10px', border:'2px solid',
                                                                borderColor: isSelected ? 'var(--sunny-blue)' : '#E2E8F0',
                                                                background: isSelected ? 'var(--sunny-blue)' : 'white',
                                                                color: isSelected ? 'white' : '#718096',
                                                                cursor:'pointer', transition:'0.2s', display:'flex', alignItems:'center', justifyContent:'center', fontSize:'1.2rem'
                                                            }}
                                                        >
                                                            {isSelected ? '✔️' : '➕'}
                                                        </button>
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}

                {activeTab === 'oportunidades' && (
                    <div className="fade-in">
                        <div style={{background:'white', padding:'24px', borderRadius:'24px', boxShadow:'0 4px 20px rgba(0,0,0,0.04)', border:'1px solid #E2E8F0', marginBottom:'24px'}}>
                            <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'20px', flexWrap:'wrap', gap:'12px'}}>
                                <div>
                                    <h3 style={{margin:0, color:'var(--sunny-blue)', fontSize:'1.3rem'}}>🚀 Gap de Mercado (Oportunidades de Expansão)</h3>
                                    <p style={{margin:'4px 0 0 0', color:'var(--text-muted)', fontSize:'0.9rem'}}>Top 50 produtos mais vendidos na Sunny que este cliente ainda não possui no mix.</p>
                                </div>
                                <div style={{display:'flex', alignItems:'center', gap:'12px', flexWrap:'wrap'}}>
                                    <div className="filter-select-v2" style={{minWidth:'200px', display:'flex', gap:'8px', alignItems:'center'}}>
                                        <div style={{position:'relative', flex:1}}>
                                            <span className="select-icon">📅</span>
                                            <select 
                                                value={localFilters.periodo} 
                                                onChange={(e) => handlePeriodChange(e.target.value)}
                                                style={{width:'100%', padding:'8px 12px 8px 36px', borderRadius:'10px', border:'1px solid #E2E8F0', fontWeight:'700', fontSize:'0.85rem'}}
                                            >
                                                <option value="Mês Atual">Mês Atual</option>
                                                <option value="Últimos 30 Dias">Últimos 30 Dias</option>
                                                <option value="Últimos 6 Meses">Últimos 6 Meses</option>
                                                <option value="Ano Atual">Ano Atual</option>
                                                <option value="Personalizado">Período Personalizado</option>
                                            </select>
                                        </div>
                                        {localFilters.periodo === 'Personalizado' && (
                                            <div style={{display:'flex', gap:'4px', alignItems:'center'}}>
                                                <input type="date" value={localFilters.start} onChange={e => setLocalFilters({...localFilters, start: e.target.value})} style={{padding:'6px', borderRadius:'8px', border:'1px solid #E2E8F0', fontSize:'0.75rem'}} />
                                                <input type="date" value={localFilters.end} onChange={e => setLocalFilters({...localFilters, end: e.target.value})} style={{padding:'6px', borderRadius:'8px', border:'1px solid #E2E8F0', fontSize:'0.75rem'}} />
                                            </div>
                                        )}
                                    </div>
                                    <label style={{display:'flex', alignItems:'center', gap:'8px', fontSize:'0.85rem', fontWeight:'700', color:'var(--sunny-blue)', cursor:'pointer', background:'#F7FAFC', padding:'8px 14px', borderRadius:'10px', border:'1px solid #E2E8F0'}}>
                                        <input 
                                            type="checkbox" 
                                            checked={hideAssortment} 
                                            onChange={(e) => setHideAssortment(e.target.checked)}
                                            style={{width:'16px', height:'16px'}}
                                        />
                                        Ocultar Sortimento
                                    </label>
                                    <button className="btn-sec" onClick={exportDossiePDF} style={{whiteSpace:'nowrap'}}>
                                        📄 Exportar Dossiê (PDF)
                                    </button>
                                </div>
                            </div>

                            <table className="analy-table" style={{width:'100%', borderCollapse:'separate', borderSpacing:'0'}}>
                                <thead style={{position:'sticky', top:0, background:'#F8FAFC', zIndex:10}}>
                                    <tr>
                                        <th style={{padding:'16px', textAlign:'center', width:'60px'}}>Foto</th>
                                        <th style={{padding:'16px'}}>Produto (Código / Descrição)</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Disponível</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>PV Sunny</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Sug. PDV</th>
                                        <th style={{padding:'16px', textAlign:'right'}}>Giro Nacional</th>
                                        <th style={{padding:'16px', textAlign:'center'}}>Ações</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    {data.marketGap
                                        ?.filter(p => !hideAssortment || p.sortimento === 'Não')
                                        ?.map((p, i) => {
                                            const isSelected = preOrder.find(item => item.produto_id === p.produto_id);
                                            return (
                                                <tr key={i} style={{
                                                    background: isSelected ? '#F0F9FF' : (i % 2 === 0 ? '#FFFFFF' : '#FBFDFF'),
                                                    borderBottom:'1px solid #EDF2F7',
                                                    transition:'0.15s'
                                                }}>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div className="product-thumb-wrap">
                                                            {p.image_url ? (
                                                                <img src={p.image_url} alt={p.nome} className="product-thumb" loading="lazy" onError={(e) => {e.target.onerror = null; e.target.style.display='none'; e.target.parentElement.innerHTML = '<div class="img-placeholder">🚫</div>'}} />
                                                            ) : (
                                                                <div className="img-placeholder">🖼️</div>
                                                            )}
                                                        </div>
                                                    </td>
                                                    <td style={{padding:'16px'}}>
                                                        <div style={{display:'flex', alignItems:'baseline', gap:'10px', flexWrap:'wrap'}}>
                                                            <span style={{fontWeight:'900', color:'var(--sunny-blue)', fontSize:'1rem'}}>{p.produto_id}</span>
                                                            <span style={{fontWeight:'800', fontSize:'0.9rem', color:'#1A202C'}}>{p.nome}</span>
                                                            {p.rank_nacional <= 10 && (
                                                                <span style={{background:'#FEF08A', color:'#713F12', fontSize:'0.7rem', padding:'2px 8px', borderRadius:'12px', fontWeight:'900', border:'1px solid #FDE047'}}>
                                                                    ⭐ Top 10 Nacional
                                                                </span>
                                                            )}
                                                        </div>
                                                        <div style={{fontSize:'0.8rem', color:'#4A5568', marginTop:'4px', fontWeight:'600'}}>
                                                            {toTitleCase(p.marca)} {p.pack ? `• Pack: ${p.pack}` : ''} {p.sortimento !== 'Não' ? `• Sortimento: ${p.sortimento}` : ''}
                                                        </div>
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <div style={{fontWeight:'900', fontSize:'1.1rem', marginBottom:'6px', color: p.saldo > 10 ? '#166534' : '#CA8A04'}}>
                                                            {p.saldo || 0} un
                                                        </div>
                                                        {p.saldo > 100 && (
                                                            <span style={{background:'var(--sunny-yellow)', color:'black', fontSize:'0.65rem', fontWeight:'900', padding:'3px 8px', borderRadius:'6px', display:'inline-block', boxShadow:'0 2px 4px rgba(0,0,0,0.1)'}}>
                                                                ⚡ PRONTA ENTREGA
                                                            </span>
                                                        )}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center', fontWeight:'800', color:'var(--sunny-blue)', fontSize:'0.95rem'}}>
                                                        {formatCurrency(p.pv)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center', fontWeight:'700', color:'#4A5568', fontSize:'0.9rem'}}>
                                                        {formatCurrency(p.pdv)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'right', fontWeight:'900', color:'var(--sunny-blue)', fontSize:'1rem'}}>
                                                        {formatCurrency(p.total_geral)}
                                                    </td>
                                                    <td style={{padding:'16px', textAlign:'center'}}>
                                                        <button 
                                                            onClick={() => togglePreOrder(p)}
                                                            style={{
                                                                width:'40px', height:'40px', borderRadius:'10px', border:'2px solid',
                                                                borderColor: isSelected ? 'var(--sunny-blue)' : '#E2E8F0',
                                                                background: isSelected ? 'var(--sunny-blue)' : 'white',
                                                                color: isSelected ? 'white' : '#718096',
                                                                cursor:'pointer', transition:'0.2s', display:'flex', alignItems:'center', justifyContent:'center', fontSize:'1.2rem'
                                                            }}
                                                        >
                                                            {isSelected ? '✔️' : '➕'}
                                                        </button>
                                                    </td>
                                                </tr>
                                            );
                                        })}
                                </tbody>
                            </table>
                        </div>
                    </div>
                )}

                {activeTab === 'pedidos' && (
                    <div className="v360-card-custom fade-in">
                        <h3>📦 Histórico de Pedidos Detalhado</h3>
                        <table className="analy-table">
                            <thead>
                                <tr>
                                    <th>Nº Documento</th>
                                    <th>Data</th>
                                    <th>Itens (SKUs)</th>
                                    <th>Qtd Total</th>
                                    <th>Faturamento</th>
                                    <th>Status</th>
                                    <th>Ações</th>
                                </tr>
                            </thead>
                            <tbody>
                                {data.pedidos?.map(p => (
                                    <tr key={p.id}>
                                        <td style={{fontWeight:'700'}}>#{p.id}</td>
                                        <td>{formatDate(p.data)}</td>
                                        <td>{p.itens} SKUs</td>
                                        <td>{p.qtdTotal} un</td>
                                        <td style={{fontWeight:'800', color:'var(--sunny-blue)'}}>{formatCurrency(p.valor)}</td>
                                        <td>
                                            <span style={{
                                                padding:'4px 8px', borderRadius:'6px', fontSize:'0.7rem', fontWeight:'800',
                                                background: p.status === 'Atendido' ? '#DCFCE7' : '#FEF3C7',
                                                color: p.status === 'Atendido' ? '#166534' : '#92400E'
                                            }}>
                                                {p.status}
                                            </span>
                                        </td>
                                        <td><button className="btn-sec-small">Detalhes</button></td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                )}
            </div>
        </div>
    );
};


const Importacao = ({ onStartTask, onFinishTask, activeTask }) => {
    const handleUpload = async (file, type, label) => {
        if (!file) return;
        
        onStartTask(label);

        const formData = new FormData();
        formData.append('file', file);

        try {
            const res = await fetch(`/api/import/${type}`, {
                method: 'POST',
                body: formData
            });
            const data = await res.json();
            
            if (res.ok) {
                onFinishTask(`✅ ${label}: ${data.detail || data.message}`);
            } else {
                onFinishTask(`❌ Erro em ${label}: ${data.error || 'Erro desconhecido'}`);
            }
        } catch (err) {
            onFinishTask(`❌ Erro de conexão em ${label}`);
        }
    };

    const [syncModal, setSyncModal] = React.useState({ open: false, loading: false, result: null });
    const [salesSyncModal, setSalesSyncModal] = React.useState({
        open: false,
        loading: false,
        mode: 'atualizar',
        startDate: '',
        endDate: '',
        result: null,
        error: null
    });

    const handleSyncPortal = async () => {
        setSyncModal({ open: true, loading: true, result: null });
        try {
            const res = await fetch('/api/sync-portal-sunny');
            const data = await res.json();
            setSyncModal({ open: true, loading: false, result: data });
        } catch (err) {
            setSyncModal({ open: true, loading: false, result: { success: false, error: 'Erro de conexão com o servidor', detail: err.message } });
        }
    };

    const handleSyncVendas = async () => {
        setSalesSyncModal(prev => ({ ...prev, loading: true, result: null, error: null }));
        try {
            const res = await fetch('/api/sync-nfs', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({
                    mode: salesSyncModal.mode,
                    startDate: salesSyncModal.startDate,
                    endDate: salesSyncModal.endDate
                })
            });
            const data = await res.json();
            if (res.ok && data.success) {
                setSalesSyncModal(prev => ({ ...prev, loading: false, result: data }));
            } else {
                setSalesSyncModal(prev => ({ ...prev, loading: false, error: data.error || 'Erro desconhecido na sincronização.' }));
            }
        } catch (err) {
            setSalesSyncModal(prev => ({ ...prev, loading: false, error: err.message || 'Erro de conexão com o servidor.' }));
        }
    };

    const modules = [
        { id: 'estoque', label: 'Estoque Disponível', icon: '📦', desc: 'Saldos, PV/PDV e Previsões.' },
        { id: 'vendas', label: 'Vendas (ERP Protheus)', icon: '📊', desc: 'Histórico de faturamento e mix.' },
        { id: 'galeria', label: 'Galeria de Fotos', icon: '🖼️', desc: 'Vínculo de imagens (URL) por SKU.' },
        { id: 'marcas', label: 'Cadastro de Marcas', icon: '🏷️', desc: 'Alimenta novas marcas e produtos não cadastrados.' },
        { id: 'clientes', label: 'Perfil de Clientes', icon: '🏢', desc: 'Atualização de perfis (Shopping, Rua, etc).' }
    ];

    return (
        <div className="fade-in">
            {/* Modal de Sincronização Portal Sunny */}
            {syncModal.open && (
                <div style={{
                    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', zIndex: 9999,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(6px)'
                }}>
                    <div style={{
                        background: '#fff', borderRadius: '20px', padding: '40px', maxWidth: '480px', width: '90%',
                        boxShadow: '0 25px 60px rgba(0,0,0,0.3)', textAlign: 'center', position: 'relative'
                    }}>
                        {syncModal.loading ? (
                            <>
                                <div style={{ fontSize: '3rem', marginBottom: '16px', animation: 'spin 1s linear infinite', display: 'inline-block' }}>🔄</div>
                                <h3 style={{ color: '#003087', fontSize: '1.3rem', marginBottom: '8px' }}>Conectando ao Portal Sunny</h3>
                                <p style={{ color: '#718096', fontSize: '0.9rem', marginBottom: '20px' }}>Autenticando e buscando dados de estoque...</p>
                                <div style={{
                                    height: '4px', background: '#E2E8F0', borderRadius: '2px', overflow: 'hidden'
                                }}>
                                    <div style={{
                                        height: '100%', width: '60%', background: 'linear-gradient(90deg, #003087, #FFD700)',
                                        borderRadius: '2px', animation: 'shimmer 1.5s ease-in-out infinite'
                                    }}/>
                                </div>
                            </>
                        ) : syncModal.result?.success ? (
                            <>
                                <div style={{ fontSize: '3.5rem', marginBottom: '16px' }}>✅</div>
                                <h3 style={{ color: '#276749', fontSize: '1.3rem', marginBottom: '8px' }}>Sincronização Concluída!</h3>
                                <div style={{
                                    display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px',
                                    margin: '20px 0', textAlign: 'left'
                                }}>
                                    <div style={{ background: '#EBF8FF', borderRadius: '12px', padding: '16px', borderLeft: '4px solid #3182CE' }}>
                                        <div style={{ fontSize: '2rem', fontWeight: '800', color: '#2B6CB0' }}>{syncModal.result.atualizados}</div>
                                        <div style={{ fontSize: '0.8rem', color: '#4A5568', marginTop: '4px' }}>Produtos Atualizados</div>
                                    </div>
                                    <div style={{ background: '#F0FFF4', borderRadius: '12px', padding: '16px', borderLeft: '4px solid #38A169' }}>
                                        <div style={{ fontSize: '2rem', fontWeight: '800', color: '#276749' }}>{syncModal.result.inseridos}</div>
                                        <div style={{ fontSize: '0.8rem', color: '#4A5568', marginTop: '4px' }}>Novos Inseridos</div>
                                    </div>
                                </div>
                                {syncModal.result.erros > 0 && (
                                    <div style={{ background: '#FFF5F5', borderRadius: '8px', padding: '10px', marginBottom: '12px', color: '#C53030', fontSize: '0.85rem' }}>
                                        ⚠️ {syncModal.result.erros} registros com erro (ignorados)
                                    </div>
                                )}
                                <p style={{ color: '#718096', fontSize: '0.8rem', marginBottom: '20px' }}>
                                    Endpoint: <code style={{ background: '#EDF2F7', padding: '2px 6px', borderRadius: '4px' }}>{syncModal.result.endpoint}</code> • Total: {syncModal.result.total} registros
                                </p>
                                <button onClick={() => setSyncModal({ open: false, loading: false, result: null })} style={{
                                    background: '#003087', color: '#fff', border: 'none', borderRadius: '10px',
                                    padding: '12px 32px', fontSize: '1rem', cursor: 'pointer', fontWeight: '600'
                                }}>Fechar</button>
                            </>
                        ) : (
                            <>
                                <div style={{ fontSize: '3.5rem', marginBottom: '16px' }}>❌</div>
                                <h3 style={{ color: '#C53030', fontSize: '1.3rem', marginBottom: '8px' }}>Erro na Sincronização</h3>
                                <div style={{ background: '#FFF5F5', borderRadius: '10px', padding: '16px', margin: '16px 0', textAlign: 'left' }}>
                                    <p style={{ color: '#C53030', fontSize: '0.9rem', margin: 0, fontFamily: 'monospace' }}>
                                        {syncModal.result?.error || 'Erro desconhecido'}
                                    </p>
                                    {syncModal.result?.detail && (
                                        <p style={{ color: '#718096', fontSize: '0.8rem', margin: '8px 0 0 0' }}>{syncModal.result.detail}</p>
                                    )}
                                </div>
                                <p style={{ color: '#718096', fontSize: '0.8rem', marginBottom: '20px' }}>Verifique as credenciais nas variáveis de ambiente do Railway</p>
                                <div style={{ display: 'flex', gap: '12px', justifyContent: 'center' }}>
                                    <button onClick={handleSyncPortal} style={{
                                        background: '#003087', color: '#fff', border: 'none', borderRadius: '10px',
                                        padding: '12px 24px', fontSize: '0.9rem', cursor: 'pointer', fontWeight: '600'
                                    }}>Tentar Novamente</button>
                                    <button onClick={() => setSyncModal({ open: false, loading: false, result: null })} style={{
                                        background: '#EDF2F7', color: '#4A5568', border: 'none', borderRadius: '10px',
                                        padding: '12px 24px', fontSize: '0.9rem', cursor: 'pointer', fontWeight: '600'
                                    }}>Fechar</button>
                                </div>
                            </>
                        )}
                    </div>
                </div>
            )}

            {/* Modal de Sincronização Protheus ERP */}
            {salesSyncModal.open && (
                <div style={{
                    position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.6)', zIndex: 9999,
                    display: 'flex', alignItems: 'center', justifyContent: 'center', backdropFilter: 'blur(6px)'
                }}>
                    <div style={{
                        background: '#fff', borderRadius: '20px', padding: '40px', maxWidth: '520px', width: '90%',
                        boxShadow: '0 25px 60px rgba(0,0,0,0.3)', position: 'relative'
                    }}>
                        {salesSyncModal.loading ? (
                            <div style={{ textAlign: 'center' }}>
                                <div style={{ fontSize: '3rem', marginBottom: '16px', animation: 'spin 1s linear infinite', display: 'inline-block' }}>🔄</div>
                                <h3 style={{ color: '#003087', fontSize: '1.3rem', marginBottom: '8px' }}>Sincronizando com Protheus ERP</h3>
                                <p style={{ color: '#718096', fontSize: '0.9rem', marginBottom: '20px' }}>
                                    {salesSyncModal.mode === 'carga' 
                                        ? `Buscando notas fiscais do período ${salesSyncModal.startDate.split('-').reverse().join('/')} a ${salesSyncModal.endDate.split('-').reverse().join('/')}...`
                                        : salesSyncModal.mode === 'revisar'
                                        ? 'Revisando notas fiscais dos últimos 30 dias...'
                                        : 'Atualizando notas fiscais (modo rápido)...'}
                                </p>
                                <p style={{ color: 'var(--text-muted)', fontSize: '0.8rem', marginBottom: '20px' }}>
                                    Isso pode levar alguns minutos dependendo do período selecionado.
                                </p>
                                <div style={{
                                    height: '4px', background: '#E2E8F0', borderRadius: '2px', overflow: 'hidden'
                                }}>
                                    <div style={{
                                        height: '100%', width: '70%', background: 'linear-gradient(90deg, #FFD700, #003087)',
                                        borderRadius: '2px', animation: 'shimmer 1.5s ease-in-out infinite'
                                    }}/>
                                </div>
                            </div>
                        ) : salesSyncModal.result ? (
                            <div style={{ textAlign: 'center' }}>
                                <div style={{ fontSize: '3.5rem', marginBottom: '16px' }}>✅</div>
                                <h3 style={{ color: '#276749', fontSize: '1.3rem', marginBottom: '8px' }}>Notas Fiscais Sincronizadas!</h3>
                                <div style={{
                                    display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px',
                                    margin: '20px 0', textAlign: 'left'
                                }}>
                                    <div style={{ background: '#EBF8FF', borderRadius: '12px', padding: '16px', borderLeft: '4px solid #3182CE' }}>
                                        <div style={{ fontSize: '2rem', fontWeight: '800', color: '#2B6CB0' }}>{salesSyncModal.result.updated}</div>
                                        <div style={{ fontSize: '0.8rem', color: '#4A5568', marginTop: '4px' }}>Notas Atualizadas/Revisadas</div>
                                    </div>
                                    <div style={{ background: '#F0FFF4', borderRadius: '12px', padding: '16px', borderLeft: '4px solid #38A169' }}>
                                        <div style={{ fontSize: '2rem', fontWeight: '800', color: '#276749' }}>{salesSyncModal.result.inserted}</div>
                                        <div style={{ fontSize: '0.8rem', color: '#4A5568', marginTop: '4px' }}>Novas Notas Inseridas</div>
                                    </div>
                                </div>
                                <p style={{ color: '#718096', fontSize: '0.85rem', marginBottom: '24px' }}>
                                    {salesSyncModal.result.detail || salesSyncModal.result.message}
                                </p>
                                <button onClick={() => setSalesSyncModal({ open: false, loading: false, mode: 'atualizar', startDate: '', endDate: '', result: null, error: null })} style={{
                                    background: '#003087', color: '#fff', border: 'none', borderRadius: '10px',
                                    padding: '12px 32px', fontSize: '1rem', cursor: 'pointer', fontWeight: '600', width: '100%'
                                }}>Fechar</button>
                            </div>
                        ) : salesSyncModal.error ? (
                            <div style={{ textAlign: 'center' }}>
                                <div style={{ fontSize: '3.5rem', marginBottom: '16px' }}>❌</div>
                                <h3 style={{ color: '#C53030', fontSize: '1.3rem', marginBottom: '8px' }}>Falha na Sincronização</h3>
                                <div style={{ background: '#FFF5F5', borderRadius: '10px', padding: '16px', margin: '16px 0', textAlign: 'left' }}>
                                    <p style={{ color: '#C53030', fontSize: '0.9rem', margin: 0, fontFamily: 'monospace' }}>
                                        {salesSyncModal.error}
                                    </p>
                                </div>
                                <div style={{ display: 'flex', gap: '12px', marginTop: '24px' }}>
                                    <button onClick={handleSyncVendas} style={{
                                        flex: 1, background: '#003087', color: '#fff', border: 'none', borderRadius: '10px',
                                        padding: '12px', fontSize: '0.95rem', cursor: 'pointer', fontWeight: '600'
                                    }}>Tentar Novamente</button>
                                    <button onClick={() => setSalesSyncModal(prev => ({ ...prev, error: null }))} style={{
                                        flex: 1, background: '#EDF2F7', color: '#4A5568', border: 'none', borderRadius: '10px',
                                        padding: '12px', fontSize: '0.95rem', cursor: 'pointer', fontWeight: '600'
                                    }}>Voltar</button>
                                </div>
                            </div>
                        ) : (
                            <div>
                                <h3 style={{ color: '#003087', fontSize: '1.4rem', margin: '0 0 8px 0', textAlign: 'center' }}>Sincronização de Notas Fiscais</h3>
                                <p style={{ color: '#718096', fontSize: '0.85rem', margin: '0 0 24px 0', textAlign: 'center' }}>
                                    Obtenha as notas fiscais (faturamentos reais) diretamente do Protheus ERP.
                                </p>

                                <div style={{ display: 'flex', flexDirection: 'column', gap: '12px', marginBottom: '24px' }}>
                                    {/* Card Atualizar */}
                                    <label style={{
                                        display: 'flex', alignItems: 'flex-start', gap: '12px', padding: '16px',
                                        borderRadius: '12px', border: `2px solid ${salesSyncModal.mode === 'atualizar' ? '#003087' : '#E2E8F0'}`,
                                        background: salesSyncModal.mode === 'atualizar' ? '#F0F5FF' : '#fff',
                                        cursor: 'pointer', transition: 'all 0.2s ease'
                                    }}>
                                        <input 
                                            type="radio" 
                                            name="syncMode" 
                                            value="atualizar"
                                            checked={salesSyncModal.mode === 'atualizar'}
                                            onChange={() => setSalesSyncModal(prev => ({ ...prev, mode: 'atualizar' }))}
                                            style={{ marginTop: '3px', accentColor: '#003087' }}
                                        />
                                        <div>
                                            <strong style={{ display: 'block', fontSize: '0.95rem', color: '#1A202C' }}>Atualização Rápida</strong>
                                            <span style={{ fontSize: '0.8rem', color: '#718096' }}>Busca novas notas emitidas a partir do faturamento mais recente.</span>
                                        </div>
                                    </label>

                                    {/* Card Revisar */}
                                    <label style={{
                                        display: 'flex', alignItems: 'flex-start', gap: '12px', padding: '16px',
                                        borderRadius: '12px', border: `2px solid ${salesSyncModal.mode === 'revisar' ? '#003087' : '#E2E8F0'}`,
                                        background: salesSyncModal.mode === 'revisar' ? '#F0F5FF' : '#fff',
                                        cursor: 'pointer', transition: 'all 0.2s ease'
                                    }}>
                                        <input 
                                            type="radio" 
                                            name="syncMode" 
                                            value="revisar"
                                            checked={salesSyncModal.mode === 'revisar'}
                                            onChange={() => setSalesSyncModal(prev => ({ ...prev, mode: 'revisar' }))}
                                            style={{ marginTop: '3px', accentColor: '#003087' }}
                                        />
                                        <div>
                                            <strong style={{ display: 'block', fontSize: '0.95rem', color: '#1A202C' }}>Revisar Últimos 30 Dias</strong>
                                            <span style={{ fontSize: '0.8rem', color: '#718096' }}>Sincroniza e atualiza notas do último mês para capturar cancelamentos ou correções.</span>
                                        </div>
                                    </label>

                                    {/* Card Carga Histórica */}
                                    <label style={{
                                        display: 'flex', alignItems: 'flex-start', gap: '12px', padding: '16px',
                                        borderRadius: '12px', border: `2px solid ${salesSyncModal.mode === 'carga' ? '#003087' : '#E2E8F0'}`,
                                        background: salesSyncModal.mode === 'carga' ? '#F0F5FF' : '#fff',
                                        cursor: 'pointer', transition: 'all 0.2s ease'
                                    }}>
                                        <input 
                                            type="radio" 
                                            name="syncMode" 
                                            value="carga"
                                            checked={salesSyncModal.mode === 'carga'}
                                            onChange={() => setSalesSyncModal(prev => ({ ...prev, mode: 'carga' }))}
                                            style={{ marginTop: '3px', accentColor: '#003087' }}
                                        />
                                        <div>
                                            <strong style={{ display: 'block', fontSize: '0.95rem', color: '#1A202C' }}>Carga Histórica Customizada</strong>
                                            <span style={{ fontSize: '0.8rem', color: '#718096' }}>Importar notas fiscais de um intervalo específico (Ex: um mês inteiro de 2025).</span>
                                        </div>
                                    </label>
                                </div>

                                {/* Form de Carga Histórica */}
                                {salesSyncModal.mode === 'carga' && (
                                    <div className="fade-in" style={{
                                        background: '#F7FAFC', border: '1px solid #E2E8F0', borderRadius: '12px',
                                        padding: '16px', marginBottom: '24px'
                                    }}>
                                        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '12px' }}>
                                            <div>
                                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: '700', marginBottom: '4px', color: '#4A5568' }}>Data Início:</label>
                                                <input 
                                                    type="date" 
                                                    value={salesSyncModal.startDate}
                                                    onChange={(e) => setSalesSyncModal(prev => ({ ...prev, startDate: e.target.value }))}
                                                    style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #CBD5E1', fontSize: '0.85rem' }}
                                                />
                                            </div>
                                            <div>
                                                <label style={{ display: 'block', fontSize: '0.75rem', fontWeight: '700', marginBottom: '4px', color: '#4A5568' }}>Data Fim:</label>
                                                <input 
                                                    type="date" 
                                                    value={salesSyncModal.endDate}
                                                    onChange={(e) => setSalesSyncModal(prev => ({ ...prev, endDate: e.target.value }))}
                                                    style={{ width: '100%', padding: '8px', borderRadius: '6px', border: '1px solid #CBD5E1', fontSize: '0.85rem' }}
                                                />
                                            </div>
                                        </div>

                                        <div style={{
                                            background: '#FFFBEB', borderLeft: '4px solid #D97706', padding: '10px 12px',
                                            borderRadius: '6px', fontSize: '0.75rem', color: '#92400E', lineHeight: '1.4'
                                        }}>
                                            <strong>⚠️ Limitação de Data Protheus:</strong> O Protheus ERP rejeita buscas antes de <strong>02/01/2025</strong>. 
                                            Para importar notas de 2020 a 2024, use a importação manual de planilhas Excel.
                                        </div>
                                    </div>
                                )}

                                <div style={{ display: 'flex', gap: '12px' }}>
                                    <button 
                                        onClick={handleSyncVendas} 
                                        disabled={
                                            salesSyncModal.mode === 'carga' && 
                                            (!salesSyncModal.startDate || !salesSyncModal.endDate || salesSyncModal.startDate < '2025-01-02' || salesSyncModal.startDate > salesSyncModal.endDate)
                                        }
                                        style={{
                                            flex: 1, background: '#003087', color: '#fff', border: 'none', borderRadius: '10px',
                                            padding: '12px', fontSize: '0.95rem', cursor: 'pointer', fontWeight: '600',
                                            opacity: (salesSyncModal.mode === 'carga' && (!salesSyncModal.startDate || !salesSyncModal.endDate || salesSyncModal.startDate < '2025-01-02' || salesSyncModal.startDate > salesSyncModal.endDate)) ? 0.5 : 1
                                        }}
                                    >
                                        Iniciar Sincronização
                                    </button>
                                    <button 
                                        onClick={() => setSalesSyncModal(prev => ({ ...prev, open: false }))} 
                                        style={{
                                            flex: 1, background: '#EDF2F7', color: '#4A5568', border: 'none', borderRadius: '10px',
                                            padding: '12px', fontSize: '0.95rem', cursor: 'pointer', fontWeight: '600'
                                        }}
                                    >
                                        Cancelar
                                    </button>
                                </div>
                            </div>
                        )}
                    </div>
                </div>
            )}

            <header style={{marginBottom:'32px'}}>
                <h2 style={{color:'var(--sunny-blue)', fontSize:'2rem', marginBottom:'8px'}}>📥 Importação de Dados</h2>
                <p style={{color:'var(--text-muted)'}}>Alimente a inteligência Sunny sem precisar parar seu trabalho.</p>
            </header>

            <div className="import-grid">
                {modules.map(mod => {
                    const isRunning = activeTask?.name === mod.label;
                    const isEstoque = mod.id === 'estoque';
                    const isVendas = mod.id === 'vendas';
                    return (
                        <div key={mod.id} className="import-card shadow-sm">
                            <div className="import-icon">{mod.icon}</div>
                            <h3>{mod.label}</h3>
                            <p>{mod.desc}</p>
                            
                            <div className="upload-zone" style={{opacity: isRunning ? 0.5 : 1, pointerEvents: isRunning ? 'none' : 'auto'}}>
                                <input 
                                    type="file" 
                                    accept=".xlsx,.xls" 
                                    onChange={(e) => handleUpload(e.target.files[0], mod.id, mod.label)}
                                    id={`file-${mod.id}`}
                                    style={{display:'none'}}
                                />
                                <label htmlFor={`file-${mod.id}`} style={{cursor:'pointer', display:'block', padding:'20px'}}>
                                    {isRunning ? '🚀 Processando...' : 'Clique para selecionar planilha'}
                                </label>
                            </div>

                            {/* Botão exclusivo do card de Estoque */}
                            {isEstoque && (
                                <button
                                    id="btn-sync-portal-sunny"
                                    onClick={handleSyncPortal}
                                    disabled={syncModal.loading}
                                    style={{
                                        marginTop: '12px', width: '100%', padding: '11px 16px',
                                        background: syncModal.loading
                                            ? '#E2E8F0'
                                            : 'linear-gradient(135deg, #003087 0%, #0056D2 100%)',
                                        color: syncModal.loading ? '#A0AEC0' : '#fff',
                                        border: '2px solid transparent',
                                        borderRadius: '10px', fontSize: '0.85rem', fontWeight: '700',
                                        cursor: syncModal.loading ? 'not-allowed' : 'pointer',
                                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                                        transition: 'all 0.2s ease', letterSpacing: '0.3px'
                                    }}
                                    onMouseEnter={e => { if (!syncModal.loading) { e.target.style.transform = 'translateY(-1px)'; e.target.style.boxShadow = '0 6px 20px rgba(0,48,135,0.4)'; }}}
                                    onMouseLeave={e => { e.target.style.transform = ''; e.target.style.boxShadow = ''; }}
                                >
                                    <span style={{ fontSize: '1rem' }}>{syncModal.loading ? '🔄' : '🌐'}</span>
                                    {syncModal.loading ? 'Sincronizando...' : 'Sincronizar via Portal Sunny'}
                                </button>
                            )}

                            {/* Botão exclusivo do card de Vendas */}
                            {isVendas && (
                                <button
                                    id="btn-sync-protheus-vendas"
                                    onClick={() => setSalesSyncModal(prev => ({ ...prev, open: true }))}
                                    style={{
                                        marginTop: '12px', width: '100%', padding: '11px 16px',
                                        background: 'linear-gradient(135deg, #FFD700 0%, #FFA500 100%)',
                                        color: '#001D54',
                                        border: '2px solid transparent',
                                        borderRadius: '10px', fontSize: '0.85rem', fontWeight: '700',
                                        cursor: 'pointer',
                                        display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px',
                                        transition: 'all 0.2s ease', letterSpacing: '0.3px'
                                    }}
                                    onMouseEnter={e => { e.target.style.transform = 'translateY(-1px)'; e.target.style.boxShadow = '0 6px 20px rgba(255,215,0,0.4)'; }}
                                    onMouseLeave={e => { e.target.style.transform = ''; e.target.style.boxShadow = ''; }}
                                >
                                    <span style={{ fontSize: '1rem' }}>🌐</span>
                                    Sincronizar via Protheus ERP
                                </button>
                            )}

                            {isRunning && (
                                <div className="progress-container">
                                    <div className="progress-bar-loading"></div>
                                    <div className="progress-text">Trabalhando em segundo plano...</div>
                                </div>
                            )}
                        </div>
                    );
                })}
            </div>

            <div className="v360-card-custom" style={{marginTop:'32px', background:'#F7FAFC'}}>
                <h4 style={{margin:'0 0 16px 0', color:'var(--sunny-blue)'}}>💡 Dicas de Importação</h4>
                <ul style={{fontSize:'0.85rem', color:'#4A5568', display:'grid', gridTemplateColumns:'1fr 1fr', gap:'12px'}}>
                    <li>✅ Você pode navegar para outras telas enquanto o upload termina.</li>
                    <li>✅ O sistema identifica colunas automaticamente (Busca Inteligente).</li>
                    <li>✅ Dados duplicados (mesma NF) são ignorados automaticamente.</li>
                    <li>✅ O código do produto é limpo (sanitizado) no momento da entrada.</li>
                </ul>
            </div>
        </div>
    );
};

const GestaoDados = () => {
    const [subTab, setSubTab] = useState('vendas');
    const [stats, setStats] = useState({ lastSale: '-', vendas: 0, estoque: 0, marcas: 0 });
    const [data, setData] = useState([]);
    const [loading, setLoading] = useState(false);
    const [filters, setFilters] = useState({ search: '', start: '', end: '', page: 1 });
    const [totalRegistros, setTotalRegistros] = useState(0);
    const [totalPages, setTotalPages] = useState(1);
    const [editingMarca, setEditingMarca] = useState(null);

    useEffect(() => {
        fetchStats();
        loadData();
    }, [subTab, filters.page]);

    const fetchStats = async () => {
        try {
            const res = await fetch('/api/admin/stats');
            const json = await res.json();
            setStats(json);
        } catch (e) { console.error(e); }
    };

    const loadData = async () => {
        setLoading(true);
        try {
            let url = `/api/admin/${subTab}?search=${filters.search}&start=${filters.start}&end=${filters.end}&page=${filters.page}`;
            const res = await fetch(url);
            const json = await res.json();
            setData(json.data || []);
            setTotalRegistros(json.total || 0);
            setTotalPages(Math.ceil((json.total || 0) / (json.limit || 50)));
        } catch (e) { console.error(e); }
        setLoading(false);
    };

    const handleExportarBase = async () => {
        setLoading(true);
        try {
            let url = `/api/admin/${subTab}?search=${filters.search}&start=${filters.start}&end=${filters.end}&export=true`;
            const res = await fetch(url);
            const json = await res.json();
            const exportData = json.data || [];
            
            if(!exportData.length) {
                alert('Nenhum dado encontrado para exportar.');
                setLoading(false);
                return;
            }
            
            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, subTab);
            XLSX.writeFile(wb, `Exportacao_Completa_${subTab}.xlsx`);
        } catch(e) {
            console.error(e);
            alert("Erro ao exportar base.");
        }
        setLoading(false);
    };

    const handleUpdateMarca = async (id, novoNome) => {
        if (!novoNome) return;
        const res = await fetch(`/api/admin/marcas/${id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ novoNome })
        });
        if (res.ok) {
            setEditingMarca(null);
            loadData();
            alert("Marca e produtos vinculados atualizados com sucesso!");
        }
    };

    const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val || 0);

    const handleClearVendas = async () => {
        if (!window.confirm("⚠️ ATENÇÃO: Isso apagará TODO o histórico de vendas importado. Deseja continuar?")) return;
        const res = await fetch(`/api/admin/vendas/clear`, { method: 'DELETE' });
        if (res.ok) {
            alert("Base de vendas limpa!");
            loadData();
            fetchStats();
        }
    };

    const handleClearSearch = () => {
        const newFilters = { ...filters, search: '', page: 1 };
        setFilters(newFilters);
        loadData(newFilters);
    };

    return (
        <div className="fade-in">
            <header style={{marginBottom:'24px', display:'flex', justifyContent:'space-between', alignItems:'flex-end'}}>
                <div>
                    <h2 style={{color:'var(--sunny-blue)', margin:0, fontSize:'1.8rem'}}>🛠️ Gerenciamento de Dados</h2>
                    <p style={{color:'var(--text-muted)', margin:'4px 0 0 0'}}>Auditoria e saneamento das bases importadas • <strong style={{color:'var(--sunny-blue)'}}>Total de Registros na Base: {totalRegistros}</strong></p>
                </div>
                <div style={{display:'flex', gap:'16px'}}>
                    {subTab === 'vendas' && (
                        <button 
                            className="btn-clear-minimal" 
                            onClick={handleClearVendas}
                            style={{color:'var(--danger)', borderColor:'var(--danger)', padding:'8px 16px'}}
                        >
                            🗑️ Limpar Base de Vendas
                        </button>
                    )}
                    <button 
                        className="btn-pri" 
                        onClick={handleExportarBase}
                        style={{display:'flex', gap:'8px', alignItems:'center', background:'#EBF8FF', color:'var(--sunny-blue)', border:'1px solid #BEE3F8', padding:'8px 16px', borderRadius:'10px', fontSize:'0.85rem', fontWeight:'800'}}
                    >
                        <span>📥</span> Exportar Base Atual (Excel)
                    </button>
                    <div className="kpi-card-v2" style={{padding:'10px 20px', minWidth:'auto', flexDirection:'row', gap:'12px', alignItems:'center'}}>
                        <span style={{fontSize:'1.2rem'}}>📅</span>
                        <div>
                            <small style={{display:'block', fontSize:'0.65rem', color:'var(--text-muted)'}}>Última Venda</small>
                            <strong style={{fontSize:'0.9rem'}}>{stats.lastSale}</strong>
                        </div>
                    </div>
                </div>
            </header>

            <div style={{display:'flex', gap:'12px', marginBottom:'24px', borderBottom:'1px solid #E2E8F0', paddingBottom:'12px'}}>
                <button className={`admin-nav-btn ${subTab==='vendas'?'active':''}`} onClick={()=>setSubTab('vendas')}>📦 Base de Vendas</button>
                <button className={`admin-nav-btn ${subTab==='estoque'?'active':''}`} onClick={()=>setSubTab('estoque')}>🏭 Gestão de Estoque</button>
                <button className={`admin-nav-btn ${subTab==='marcas'?'active':''}`} onClick={()=>setSubTab('marcas')}>🏷️ Gestão de Marcas</button>
                <button className={`admin-nav-btn ${subTab==='fotos'?'active':''}`} onClick={()=>setSubTab('fotos')}>🖼️ Gestão de Fotos</button>
                <button className={`admin-nav-btn ${subTab==='clientes'?'active':''}`} onClick={()=>setSubTab('clientes')}>🏢 Gestão de Clientes</button>
            </div>

            <div style={{display:'flex', gap:'16px', marginBottom:'20px', background:'white', padding:'16px', borderRadius:'16px', border:'1px solid #E2E8F0', alignItems:'center'}}>
                <div className="search-global-wrap" style={{flex:1, margin:0, position:'relative'}}>
                    <span className="lupa">🔍</span>
                    <input 
                        type="text" 
                        placeholder="Buscar por NF, Cliente, EAN ou Código..." 
                        style={{border:0, width:'100%', paddingRight:'35px'}}
                        value={filters.search}
                        onKeyDown={(e) => e.key === 'Enter' && loadData()}
                        onChange={e => setFilters({...filters, search: e.target.value})}
                    />
                    {filters.search && (
                        <button 
                            onClick={handleClearSearch}
                            style={{position:'absolute', right:'8px', top:'50%', transform:'translateY(-50%)', background:'transparent', border:0, cursor:'pointer', color:'#A0AEC0', fontSize:'1rem', fontWeight:'900', padding:'4px'}}
                        >
                            ✕
                        </button>
                    )}
                </div>
                <button className="btn-view" onClick={() => loadData()} style={{padding:'0 24px'}}>Buscar</button>
            </div>

            <div className="table-card-v21" style={{maxHeight:'550px', overflow:'auto'}}>
                {loading ? <div style={{padding:'60px', textAlign:'center', color:'var(--text-muted)'}}>Carregando base de dados...</div> : (
                    <table className="analy-table">
                        <thead style={{position:'sticky', top:0, zIndex:10, background:'white'}}>
                            {subTab === 'vendas' ? (
                                <tr>
                                    <th>NF</th><th>Emissão</th><th>CNPJ</th><th>Cliente</th><th>Produto</th><th>Descrição</th><th>Qtd</th><th>Total</th><th>Vendedor</th><th>Gerente</th><th>UF</th>
                                </tr>
                            ) : subTab === 'estoque' ? (
                                <tr>
                                    <th>Código</th><th>EAN</th><th>Descrição</th><th>Marca</th><th>Saldo</th><th>PV</th><th>Previsão</th>
                                </tr>
                            ) : subTab === 'clientes' ? (
                                <tr>
                                    <th>CÓDIGO</th><th>Cliente</th><th>Perfil</th><th>Ações</th>
                                </tr>
                            ) : subTab === 'fotos' ? (
                                <tr>
                                    <th>Thumbnail</th><th>CÓDIGO</th><th>Descrição</th><th>URL da Imagem</th><th>Ações</th>
                                </tr>
                            ) : (
                                <tr>
                                    <th>CÓDIGO</th><th>Produto</th><th>Marca</th><th>Ações</th>
                                </tr>
                            )}
                        </thead>
                        <tbody>
                            {data.length === 0 ? (
                                <tr><td colSpan="15" style={{padding:'40px', textAlign:'center'}}>Nenhum registro encontrado nesta base.</td></tr>
                            ) : data.map((row, i) => (
                                <tr key={i}>
                                    {subTab === 'vendas' ? (
                                        <>
                                            <td><strong>{row.num_docto}</strong></td>
                                            <td>{row.emissao}</td>
                                            <td>{row.cnpj || <span style={{color:'var(--danger)', fontSize:'0.7rem'}}>NÃO ENCONTRADO</span>}</td>
                                            <td style={{maxWidth:'180px', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}} title={row.nome_cliente}>
                                                {row.nome_cliente ? row.nome_cliente : <span style={{color:'var(--danger)'}}>{row.cliente_id} (SEM NOME)</span>}
                                            </td>
                                            <td>{row.produto_id}</td>
                                            <td style={{maxWidth:'200px', fontSize:'0.75rem', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}} title={row.descricao_produto}>{row.descricao_produto || <span style={{color:'var(--danger)'}}>NÃO ENCONTRADO</span>}</td>
                                            <td>{row.quantidade}</td>
                                            <td style={{fontWeight:'700'}}>{formatCurrency(row.valor_total)}</td>
                                            <td>{row.nome_vendedor || <span style={{color:'var(--danger)'}}>NÃO ENCONTRADO</span>}</td>
                                            <td>{row.gerente_id || '-'}</td>
                                            <td>{row.uf || <span style={{color:'var(--danger)', fontSize:'0.7rem'}}>NÃO ENCONTRADO</span>}</td>
                                        </>
                                    ) : subTab === 'estoque' ? (
                                        <>
                                            <td><strong>{row.cod_produto}</strong></td>
                                            <td>{row.ean}</td>
                                            <td>{row.descricao}</td>
                                            <td><span className="status-badge" style={{background:'#F0F7FF', color:'#3182CE'}}>{row.marca}</span></td>
                                            <td style={{color: row.saldo > 0 ? 'var(--success)' : 'var(--danger)', fontWeight:'700'}}>{row.saldo}</td>
                                            <td>{formatCurrency(row.pv)}</td>
                                            <td>{row.previsao || '-'}</td>
                                        </>
                                    ) : subTab === 'clientes' ? (
                                        <>
                                            <td style={{fontWeight:'800', color:'var(--sunny-blue)'}}>{row.cliente_id}</td>
                                            <td style={{fontSize:'0.85rem'}}>{row.nome_cliente}</td>
                                            <td><span className="status-badge" style={{background:'#EBF8FF', color:'#2B6CB0', fontWeight:'700'}}>{row.perfil || 'Rua'}</span></td>
                                            <td>
                                                <button className="btn-view" style={{padding:'4px 12px', fontSize:'0.75rem'}} onClick={()=>setEditingMarca({ ...row, cod_produto: row.cliente_id, descricao: row.nome_cliente, marca: row.perfil, isCliente: true })}>Alterar Perfil</button>
                                            </td>
                                        </>
                                    ) : subTab === 'fotos' ? (
                                        <>
                                            <td style={{textAlign:'center'}}>
                                                <div className="product-thumb-wrap">
                                                    {row.image_url ? (
                                                        <img src={row.image_url} alt={row.descricao} className="product-thumb" loading="lazy" onError={(e) => {e.target.onerror = null; e.target.style.display='none'; e.target.parentElement.innerHTML = '<div class="img-placeholder">🚫</div>'}} />
                                                    ) : (
                                                        <div className="img-placeholder">🖼️</div>
                                                    )}
                                                </div>
                                            </td>
                                            <td style={{fontWeight:'800', color:'var(--sunny-blue)'}}>{row.cod_produto}</td>
                                            <td style={{fontSize:'0.85rem'}}>{row.descricao}</td>
                                            <td style={{fontSize:'0.75rem', color:'var(--text-muted)', maxWidth:'250px', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}} title={row.image_url}>{row.image_url || 'Sem URL'}</td>
                                            <td>
                                                <button className="btn-view" style={{padding:'4px 12px', fontSize:'0.75rem'}} onClick={()=>setEditingMarca({ ...row, isFoto: true })}>Alterar Foto</button>
                                            </td>
                                        </>
                                    ) : (
                                        <>
                                            <td style={{fontWeight:'800', color:'var(--sunny-blue)'}}>{row.cod_produto}</td>
                                            <td style={{fontSize:'0.85rem'}}>{row.descricao}</td>
                                            <td><span className="status-badge" style={{background:'#EBF8FF', color:'#2B6CB0', fontWeight:'700'}}>{row.marca}</span></td>
                                            <td>
                                                <button className="btn-view" style={{padding:'4px 12px', fontSize:'0.75rem'}} onClick={()=>setEditingMarca(row)}>Alterar Marca</button>
                                            </td>
                                        </>
                                    )}
                                </tr>
                            ))}
                        </tbody>
                    </table>
                )}
            </div>

            {totalPages > 1 && (
                <div style={{display:'flex', gap:'8px', marginTop:'16px', justifyContent:'center', alignItems:'center', marginBottom:'16px'}}>
                    <button disabled={filters.page === 1} onClick={() => setFilters({...filters, page: filters.page - 1})} className="btn-sec" style={{padding:'6px 12px', fontSize:'0.85rem'}}>Anterior</button>
                    <span style={{fontSize:'0.85rem', color:'var(--text-muted)', fontWeight:'600'}}>Página {filters.page} de {totalPages}</span>
                    <button disabled={filters.page === totalPages} onClick={() => setFilters({...filters, page: filters.page + 1})} className="btn-sec" style={{padding:'6px 12px', fontSize:'0.85rem'}}>Próxima</button>
                </div>
            )}

            {editingMarca && (
                <div className="modal-overlay" style={{zIndex:2000}}>
                    <div className="modal-content" style={{width:'400px', padding:'32px'}}>
                        <h3 style={{margin:'0 0 8px 0', color:'var(--sunny-blue)'}}>
                            {editingMarca.isCliente ? '🏢 Alterar Perfil do Cliente' : editingMarca.isFoto ? '🖼️ Alterar URL da Foto' : '🏷️ Alterar Marca por SKU'}
                        </h3>
                        <p style={{fontSize:'0.8rem', color:'var(--text-muted)', marginBottom:'24px'}}>
                            <strong>{editingMarca.isCliente ? 'ID Cliente:' : 'SKU:'}</strong> {editingMarca.cod_produto}<br/>
                            {editingMarca.descricao}
                        </p>
                        <label style={{display:'block', marginBottom:'8px', fontSize:'0.8rem', fontWeight:'700'}}>
                            {editingMarca.isCliente ? 'Novo Perfil:' : editingMarca.isFoto ? 'Nova URL da Imagem:' : 'Nova Marca:'}
                        </label>
                        <input 
                            type="text" 
                            className="input-minimal" 
                            defaultValue={editingMarca.isFoto ? editingMarca.image_url : editingMarca.marca} 
                            id="editProductBrandInput"
                            autoFocus
                            style={{width:'100%', marginBottom:'24px', padding:'12px', border:'1px solid #E2E8F0', borderRadius:'8px'}}
                        />
                        <div style={{display:'flex', gap:'12px', justifyContent:'flex-end'}}>
                            <button className="btn-clear-minimal" onClick={()=>setEditingMarca(null)}>Cancelar</button>
                            <button className="btn-sunny-small" onClick={async () => {
                                const newVal = document.getElementById('editProductBrandInput').value;
                                const url = editingMarca.isCliente 
                                    ? `/api/admin/clientes/${editingMarca.cod_produto}/perfil`
                                    : editingMarca.isFoto
                                    ? `/api/admin/produtos/${editingMarca.cod_produto}/foto`
                                    : `/api/admin/produtos/${editingMarca.cod_produto}/marca`;
                                
                                const bodyData = editingMarca.isCliente 
                                    ? { novoPerfil: newVal } 
                                    : editingMarca.isFoto 
                                    ? { novaFoto: newVal } 
                                    : { novaMarca: newVal };

                                const res = await fetch(url, {
                                    method: 'PUT',
                                    headers: {'Content-Type':'application/json'},
                                    body: JSON.stringify(bodyData)
                                });
                                if (res.ok) {
                                    setEditingMarca(null);
                                    loadData();
                                }
                            }}>Salvar e Atualizar</button>
                        </div>
                    </div>
                </div>
            )}
        </div>
    );
};

const ViewProduct360 = ({ productCode, onBack, globalFilters }) => {
    const [data, setData] = useState(null);
    const [loading, setLoading] = useState(true);

    useEffect(() => {
        fetch(`/api/produtos/${productCode}?start=${globalFilters.start}&end=${globalFilters.end}`)
            .then(res => res.json())
            .then(d => { setData(d); setLoading(false); });
    }, [productCode]);

    if (loading) return <div style={{padding:'60px', textAlign:'center', color:'var(--sunny-blue)', fontWeight:'700'}}>Analisando raio-x do produto...</div>;

    const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val || 0);

    return (
        <div className="fade-in">
            <button className="btn-back" onClick={onBack} style={{marginBottom:'24px', background:'white', border:'1px solid #E2E8F0', padding:'8px 16px', borderRadius:'10px', fontWeight:'700', cursor:'pointer', color:'var(--sunny-blue)'}}>← Voltar ao Catálogo</button>
            
            <div className="v360-card-custom" style={{display:'flex', gap:'32px', alignItems:'center', padding:'32px', marginBottom:'32px'}}>
                <div style={{width:'160px', height:'160px', background:'#F7FAFC', borderRadius:'24px', display:'flex', alignItems:'center', justifyContent:'center', border:'1px solid #E2E8F0', padding:'10px'}}>
                    {data.product?.image_url ? <img src={data.product.image_url} style={{maxWidth:'100%', maxHeight:'100%', objectFit:'contain'}} /> : <span style={{fontSize:'3rem'}}>📦</span>}
                </div>
                <div style={{flex:1}}>
                    <h1 style={{color:'var(--sunny-blue)', margin:0, fontSize:'2rem'}}>{data.product?.descricao}</h1>
                    <p style={{fontSize:'1.1rem', color:'var(--text-muted)', margin:'8px 0'}}>SKU: <strong>{data.product?.cod_produto}</strong> • Marca: <strong>{data.product?.marca}</strong></p>
                    <div style={{display:'flex', gap:'12px', marginTop:'16px'}}>
                        <span className="status-badge" style={{background:'#EBF8FF', color:'#2B6CB0', fontSize:'0.85rem', padding:'6px 12px'}}>{data.product?.categoria || 'Geral'}</span>
                        <span className="status-badge" style={{background:'#F0FFF4', color:'#22543D', fontSize:'0.85rem', padding:'6px 12px'}}>Estoque Disponível: {data.product?.saldo || 0} un</span>
                        <span className="status-badge" style={{background:'#FFFBEB', color:'#92400E', fontSize:'0.85rem', padding:'6px 12px'}}>Preço Base: {formatCurrency(data.product?.pv)}</span>
                    </div>
                </div>
            </div>

            <div style={{display:'grid', gridTemplateColumns:'2fr 1fr', gap:'24px'}}>
                <div className="v360-card-custom">
                    <h3 style={{margin:'0 0 20px 0', color:'var(--sunny-blue)', display:'flex', alignItems:'center', gap:'10px'}}>
                        <span>👥</span> Maiores Compradores do Período
                    </h3>
                    <table className="analy-table">
                        <thead>
                            <tr>
                                <th>Cliente</th>
                                <th style={{textAlign:'right'}}>Faturamento</th>
                                <th style={{textAlign:'center'}}>Volume</th>
                                <th style={{textAlign:'center'}}>Última Compra</th>
                            </tr>
                        </thead>
                        <tbody>
                            {data.topCompradores.length === 0 ? <tr><td colSpan="4" style={{padding:'40px', textAlign:'center'}}>Nenhuma venda registrada para este produto no período.</td></tr> : 
                            data.topCompradores.map(c => (
                                <tr key={c.cnpj}>
                                    <td style={{fontWeight:'700', fontSize:'0.9rem'}}>{c.nome_cliente}</td>
                                    <td style={{textAlign:'right', fontWeight:'800', color:'var(--sunny-blue)'}}>{formatCurrency(c.total)}</td>
                                    <td style={{textAlign:'center', fontWeight:'700'}}>{c.qtd} un</td>
                                    <td style={{textAlign:'center', color:'var(--text-muted)', fontSize:'0.85rem'}}>{new Date(c.ultima_compra).toLocaleDateString('pt-BR')}</td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>

                <div className="v360-card-custom" style={{background:'linear-gradient(135deg, #003087 0%, #0056D2 100%)', color:'white'}}>
                    <h3 style={{margin:'0 0 20px 0', color:'white'}}>💡 Insight do Produto</h3>
                    <div style={{background:'rgba(255,255,255,0.1)', padding:'20px', borderRadius:'16px', border:'1px solid rgba(255,255,255,0.2)'}}>
                        <p style={{fontSize:'0.95rem', lineHeight:'1.6', margin:0}}>
                            Este item representa uma oportunidade de <strong>cross-sell</strong> em clientes do perfil <strong>{globalFilters.perfil || 'Shopping'}</strong>. 
                            Considere oferecer kits promocionais para aumentar o ticket médio.
                        </p>
                    </div>
                        <p style={{fontSize:'0.75rem', marginTop:'8px', opacity:0.9}}>Aderência de estoque nacional: 85%</p>
                </div>
            </div>
        </div>
    );
};

const ModuloProdutos = () => {
    const [activeSubTab, setActiveSubTab] = useState('performance');
    const [data, setData] = useState({ stats: {}, items: [] });
    const [loading, setLoading] = useState(true);
    const [selectedProduct, setSelectedProduct] = useState(null);
    const [filters, setFilters] = useState(() => {
        const initialDates = getDynamicDateRange('Mês Atual');
        return {
            periodo: 'Mês Atual',
            start: initialDates.start,
            end: initialDates.end,
            marca: 'Todos',
            apenasOportunidade: false
        };
    });

    useEffect(() => {
        setLoading(true);
        const url = new URL('/api/produtos', window.location.origin);
        Object.keys(filters).forEach(k => {
            if(filters[k] !== undefined && filters[k] !== null) url.searchParams.append(k, filters[k]);
        });
        fetch(url).then(res => res.json()).then(d => { 
            setData(d); 
            setLoading(false); 
        }).catch(e => {
            console.error(e);
            setLoading(false);
        });
    }, [filters]);

    const handlePeriodChange = (p) => {
        if (p === 'Personalizado') {
            setFilters(prev => ({ ...prev, periodo: p }));
            return;
        }
        const { start, end } = getDynamicDateRange(p);
        setFilters(prev => ({ ...prev, periodo: p, start, end }));
    };

    const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val || 0);

    const exportProdutosPDF = async () => {
        const doc = new jsPDF();
        doc.setFontSize(18);
        doc.setTextColor(0, 48, 135);
        doc.text("Dossiê Estratégico de Produtos - Sunny", 14, 20);
        
        doc.setFontSize(10);
        doc.setTextColor(50);
        doc.text(`Análise baseada no desempenho nacional de: ${filters.periodo}`, 14, 28);
        doc.text(`Filtro de Marca: ${filters.marca}`, 14, 34);
        doc.text(`Data da Exportação: ${new Date().toLocaleDateString('pt-BR')}`, 14, 40);

        const tableData = [];
        const images = {};

        // Limita a 50 itens para evitar PDF muito pesado se houver muitos
        const exportList = data.items.slice(0, 100);

        for (let i = 0; i < exportList.length; i++) {
            const p = exportList[i];
            tableData.push([
                '', // Foto
                `${p.id}\n${p.nome}`,
                p.marca,
                new Intl.NumberFormat('pt-BR').format(p.saldo || 0),
                new Intl.NumberFormat('pt-BR').format(p.giro_30d || 0),
                `${Math.round(p.doh || 0)} dias`,
                p.statusPrevisao
            ]);

            if (p.image_url) {
                const base64 = await imageUrlToBase64(p.image_url);
                if (base64) images[i] = base64;
            }
        }

        autoTable(doc, {
            startY: 48,
            head: [['Foto', 'Produto', 'Marca', 'Saldo', 'Giro', 'DOH', 'Previsão']],
            body: tableData,
            headStyles: { fillColor: [0, 48, 135], textColor: [255, 255, 255] },
            styles: { fontSize: 7, minCellHeight: 15, verticalAlign: 'middle' },
            columnStyles: {
                0: { cellWidth: 15 },
                1: { cellWidth: 60 },
                3: { halign: 'center' },
                4: { halign: 'center' },
                5: { halign: 'center' },
                6: { fontStyle: 'bold', cellWidth: 35 }
            },
            didDrawCell: (hookData) => {
                if (hookData.section === 'body' && hookData.column.index === 0 && images[hookData.row.index]) {
                    doc.addImage(images[hookData.row.index], 'JPEG', hookData.cell.x + 2, hookData.cell.y + 2, 11, 11);
                }
            },
            didParseCell: (hookData) => {
                if (hookData.section === 'body' && hookData.column.index === 6) {
                    const item = exportList[hookData.row.index];
                    if (item && item.previsaoAtrasada) {
                        hookData.cell.styles.textColor = [229, 62, 62];
                    }
                }
            }
        });

        doc.save(`Dossie_Produtos_${filters.periodo.replace(' ', '_')}.pdf`);
    };

    if (selectedProduct) return <ViewProduct360 productCode={selectedProduct} onBack={() => setSelectedProduct(null)} globalFilters={filters} />;

    return (
        <div className="fade-in">
            <header className="main-header" style={{marginBottom:'24px'}}>
                <div className="title-area">
                    <h1 style={{fontSize:'2.2rem'}}>Radar de Produtos</h1>
                    <p style={{fontSize:'1.1rem', color:'var(--text-muted)'}}>Performance de catálogo e análise técnica de sortimento</p>
                </div>
                <div className="context-layer">
                    {activeSubTab === 'ruptura' && (
                        <div className="context-group">
                            <label style={{display:'flex', alignItems:'center', gap:'8px', cursor:'pointer', fontSize:'0.9rem', fontWeight:'700', color:'var(--sunny-blue)'}}>
                                <input type="checkbox" checked={filters.apenasOportunidade} onChange={e => setFilters({...filters, apenasOportunidade: e.target.checked})} />
                                Oportunidade de Domínio (&lt; 45 dias)
                            </label>
                        </div>
                    )}
                    <div className={`context-group ${filters.periodo === 'Personalizado' ? 'expanded' : ''}`}>
                        <span className="ctx-icon">📅</span>
                        <select value={filters.periodo} onChange={e => handlePeriodChange(e.target.value)}>
                            <option>Mês Atual</option>
                            <option>Mês Anterior</option>
                            <option>Últimos 3 Meses</option>
                            <option>Ano Atual</option>
                            <option>Personalizado</option>
                        </select>
                        {filters.periodo === 'Personalizado' && (
                            <div className="date-inputs">
                                <input type="date" value={filters.start} onChange={e => setFilters({...filters, start: e.target.value})} />
                                <span>até</span>
                                <input type="date" value={filters.end} onChange={e => setFilters({...filters, end: e.target.value})} />
                            </div>
                        )}
                    </div>
                    <div className="context-group">
                        <span className="ctx-icon">🏷️</span>
                        <select value={filters.marca} onChange={e => setFilters({...filters, marca: e.target.value})}>
                            <option value="Todos">Todas as Marcas Licenciadas</option>
                            <option value="Patrulha Canina">Patrulha Canina</option>
                            <option value="Pokemon">Pokemon</option>
                            <option value="Sonic">Sonic</option>
                            <option value="Tartarugas Ninja">Tartarugas Ninja</option>
                            <option value="Hello Kitty">Hello Kitty</option>
                            <option value="Super Mario">Super Mario</option>
                            <option value="Batman">Batman</option>
                        </select>
                    </div>
                    <div className="date-range-badge">
                        {filters.start.split('-').reverse().join('/')} - {filters.end.split('-').reverse().join('/')}
                    </div>
                    <button className="btn-sec" onClick={exportProdutosPDF} style={{padding:'8px 16px', fontSize:'0.8rem', fontWeight:'800', display:'flex', gap:'8px', alignItems:'center'}}>
                        <span>📄</span> PDF Dossiê
                    </button>
                </div>
            </header>

            <div className="kpi-grid" style={{marginBottom:'32px'}}>
                <div className="kpi-card shadow-sm" style={{borderLeft:'4px solid var(--sunny-blue)'}}>
                    <label>Faturamento Total de Itens</label>
                    <strong>{formatCurrency(data.stats.faturamentoTotal)}</strong>
                </div>
                <div className="kpi-card shadow-sm" style={{borderLeft:'4px solid #4ADE80'}}>
                    <label>Qtd de SKUs Vendidos</label>
                    <strong>{data.stats.skusVendidos || 0}</strong>
                </div>
                <div className="kpi-card shadow-sm" style={{borderLeft:'4px solid #F6AD55'}}>
                    <label>Volume de Peças</label>
                    <strong>{new Intl.NumberFormat('pt-BR').format(data.stats.volumePecas || 0)} <small style={{fontSize:'0.8rem', opacity:0.7}}>un</small></strong>
                </div>
                <div className="kpi-card shadow-sm" style={{borderLeft:'4px solid #6366F1'}}>
                    <label>Ticket Médio (R$ / Peça)</label>
                    <strong>{formatCurrency(data.stats.ticketMedioSku)}</strong>
                </div>
            </div>

            <div style={{display:'flex', gap:'12px', marginBottom:'24px', borderBottom:'1px solid #E2E8F0', paddingBottom:'12px'}}>
                <button className={`admin-nav-btn ${activeSubTab==='performance'?'active':''}`} onClick={()=>setActiveSubTab('performance')}>📈 Performance Geral</button>
                <button className={`admin-nav-btn ${activeSubTab==='ruptura'?'active':''}`} onClick={()=>setActiveSubTab('ruptura')}>⚠️ Ruptura & Oportunidade</button>
                <button className={`admin-nav-btn ${activeSubTab==='lancamentos'?'active':''}`} onClick={()=>setActiveSubTab('lancamentos')}>🚀 Lançamentos</button>
            </div>

            <div className="table-card-v21">
                {activeSubTab === 'performance' ? (
                    <table className="analy-table">
                        <thead>
                            <tr>
                                <th>Produto</th>
                                <th style={{textAlign:'center'}}>Curva ABC</th>
                                <th style={{textAlign:'right'}}>Faturamento</th>
                                <th style={{textAlign:'center'}}>Qtd Vendida</th>
                                <th style={{textAlign:'center'}}>Nº Clientes Únicos</th>
                            </tr>
                        </thead>
                        <tbody>
                            {loading ? <tr><td colSpan="5" style={{padding:'40px', textAlign:'center', color:'var(--text-muted)'}}>Compilando indicadores de performance...</td></tr> : 
                            data.items.length === 0 ? <tr><td colSpan="5" style={{padding:'40px', textAlign:'center'}}>Nenhum produto encontrado com os filtros selecionados.</td></tr> :
                            data.items.map(p => (
                                <tr key={p.id} onClick={() => setSelectedProduct(p.id)} style={{cursor:'pointer'}}>
                                    <td style={{display:'flex', gap:'16px', alignItems:'center', padding:'12px 16px'}}>
                                        <div className="product-thumb-wrap" style={{width:'50px', height:'50px'}}>
                                            {p.image_url ? <img src={p.image_url} className="product-thumb" style={{width:'100%', height:'100%', objectFit:'contain'}} /> : <div className="img-placeholder" style={{fontSize:'1.2rem'}}>📦</div>}
                                        </div>
                                        <div>
                                            <div style={{fontWeight:'900', color:'var(--sunny-blue)', fontSize:'0.95rem'}}>{p.id}</div>
                                            <div style={{fontSize:'0.8rem', fontWeight:'600', color:'#4A5568', maxWidth:'300px', overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap'}}>{p.nome}</div>
                                            <div style={{fontSize:'0.7rem', color:'var(--text-muted)', fontWeight:'700', marginTop:'2px'}}>{p.marca}</div>
                                        </div>
                                    </td>
                                    <td style={{textAlign:'center'}}>
                                        <span style={{
                                            padding:'5px 12px', borderRadius:'8px', fontWeight:'900', fontSize:'0.85rem',
                                            background: p.classe_abc === 'A' ? '#D1FAE5' : p.classe_abc === 'B' ? '#FEF3C7' : '#F3F4F6',
                                            color: p.classe_abc === 'A' ? '#065F46' : p.classe_abc === 'B' ? '#92400E' : '#374151',
                                            boxShadow:'0 1px 2px rgba(0,0,0,0.05)'
                                        }}>
                                            Curva {p.classe_abc}
                                        </span>
                                    </td>
                                    <td style={{textAlign:'right', fontWeight:'900', color:'var(--sunny-blue)', fontSize:'1rem'}}>
                                        {formatCurrency(p.faturamento)}
                                    </td>
                                    <td style={{textAlign:'center', fontWeight:'700', color:'#4A5568'}}>
                                        {new Intl.NumberFormat('pt-BR').format(p.qtd_vendida)} <small style={{fontSize:'0.7rem'}}>un</small>
                                    </td>
                                    <td style={{textAlign:'center'}}>
                                        <div style={{fontWeight:'800', color:'var(--sunny-blue)'}}>{p.clientes_unicos}</div>
                                        <small style={{fontSize:'0.65rem', color:'var(--text-muted)', fontWeight:'700'}}>Pontos de Venda</small>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                ) : activeSubTab === 'ruptura' ? (
                    <table className="analy-table">
                        <thead>
                            <tr>
                                <th>Produto</th>
                                <th style={{textAlign:'center'}}>Saldo Global</th>
                                <th style={{textAlign:'center'}}>Giro no Período</th>
                                <th style={{textAlign:'center'}}>Dias de Estoque (DOH)</th>
                                <th style={{textAlign:'center'}}>Ação Rápida</th>
                            </tr>
                        </thead>
                        <tbody>
                            {loading ? <tr><td colSpan="5" style={{padding:'40px', textAlign:'center', color:'var(--text-muted)'}}>Calculando saúde de estoque...</td></tr> : 
                            data.items.length === 0 ? <tr><td colSpan="5" style={{padding:'40px', textAlign:'center'}}>Nenhum produto em alerta de ruptura.</td></tr> :
                            data.items.map(p => (
                                <tr key={p.id} style={{background: p.doh < 30 ? '#FFF5F5' : 'transparent'}}>
                                    <td style={{display:'flex', gap:'16px', alignItems:'center', padding:'12px 16px'}}>
                                        <div className="product-thumb-wrap" style={{width:'50px', height:'50px'}}>
                                            {p.image_url ? <img src={p.image_url} className="product-thumb" style={{width:'100%', height:'100%', objectFit:'contain'}} /> : <div className="img-placeholder" style={{fontSize:'1.2rem'}}>📦</div>}
                                        </div>
                                        <div>
                                            <div style={{fontWeight:'900', color:'var(--sunny-blue)', fontSize:'0.95rem'}}>{p.id}</div>
                                            <div style={{fontSize:'0.8rem', fontWeight:'600', color:'#4A5568'}}>{p.nome}</div>
                                        </div>
                                    </td>
                                    <td style={{textAlign:'center', fontWeight:'800', color: (p.saldo || 0) < 100 ? '#E53E3E' : 'inherit'}}>
                                        {new Intl.NumberFormat('pt-BR').format(p.saldo || 0)} <small>un</small>
                                    </td>
                                    <td style={{textAlign:'center', fontWeight:'700'}}>
                                        {new Intl.NumberFormat('pt-BR').format(p.giro_30d || 0)} <small>un</small>
                                    </td>
                                    <td style={{textAlign:'center'}}>
                                        <div style={{
                                            padding:'4px 12px', borderRadius:'20px', display:'inline-block', fontWeight:'900',
                                            background: (p.doh || 0) < 30 ? '#FEE2E2' : (p.doh || 0) < 60 ? '#FEF3C7' : '#D1FAE5',
                                            color: (p.doh || 0) < 30 ? '#991B1B' : (p.doh || 0) < 60 ? '#92400E' : '#065F46',
                                            border: (p.doh || 0) < 30 ? '1px solid #F87171' : 'none'
                                        }}>
                                            {(p.doh || 0) > 365 ? '+1 ano' : `${Math.round(p.doh || 0)} dias`}
                                        </div>
                                        {(p.doh || 0) < 30 && (p.doh || 0) > 0 && <div style={{fontSize:'0.65rem', color:'#E53E3E', fontWeight:'900', marginTop:'4px'}}>⚠️ ESTOQUE CRÍTICO</div>}
                                        <div style={{fontSize:'0.7rem', marginTop:'6px', color: p.previsaoAtrasada ? '#E53E3E' : 'var(--text-muted)', fontWeight: p.previsaoAtrasada ? '800' : '600'}}>
                                            {p.statusPrevisao}
                                        </div>
                                    </td>
                                    <td style={{textAlign:'center'}}>
                                        <button className="btn-v360-small" onClick={() => setSelectedProduct(p.id)} style={{background:'var(--sunny-blue)', color:'white', border:'none', padding:'6px 12px', borderRadius:'6px', fontWeight:'700', cursor:'pointer', fontSize:'0.75rem'}}>
                                            👥 Quem já compra?
                                        </button>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                ) : (
                    <div style={{padding:'80px', textAlign:'center', color:'var(--text-muted)'}}>
                        <div style={{fontSize:'3rem', marginBottom:'16px'}}>🚀</div>
                        <h3>Módulo de Lançamentos</h3>
                        <p>Análise de curva de adoção para novos SKUs em breve.</p>
                    </div>
                )}
            </div>
        </div>
    );
};

const App = () => {
  const [activeTab, setActiveTab] = useState('clientes');
  const [backgroundTask, setBackgroundTask] = useState(null);
  const [loading, setLoading] = useState(true);
  const [filters, setFilters] = useState(() => {
    const initialDates = getDynamicDateRange('Mês Atual');
    return { 
      search: '', uf: 'Todos', vendedor: 'Todos', status: 'Todos', gerente: 'Todos', 
      periodo: 'Mês Atual', start: initialDates.start, end: initialDates.end, compare: false 
    };
  });
  const [metaFiltros, setMetaFiltros] = useState({ vendedores: [], ufs: [], gerentes: [] });
  const [clientsData, setClientsData] = useState({ items: [], kpis: {}, grupos: { criticos: 0, alerta: 0, saudavel: 0, oportunidades: 0 } });
  const [selectedClient, setSelectedClient] = useState(null);
  const [clientDetails, setClientDetails] = useState(null);
  const [sortConfig, setSortConfig] = useState({ key: 'oportunidade', dir: 'desc' });

  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      dir: prev.key === key && prev.dir === 'desc' ? 'asc' : 'desc'
    }));
  };

  const sortedItems = [...(clientsData.items || [])].sort((a, b) => {
    const va = a[sortConfig.key];
    const vb = b[sortConfig.key];
    const na = parseFloat(va) || 0;
    const nb = parseFloat(vb) || 0;
    if (!isNaN(na) && !isNaN(nb)) return sortConfig.dir === 'desc' ? nb - na : na - nb;
    const sa = String(va || '').toLowerCase();
    const sb = String(vb || '').toLowerCase();
    return sortConfig.dir === 'desc' ? sb.localeCompare(sa) : sa.localeCompare(sb);
  });

  const SortTh = ({ col, label, align = 'left' }) => {
    const isActive = sortConfig.key === col;
    return (
      <th
        onClick={() => handleSort(col)}
        style={{
          cursor: 'pointer',
          userSelect: 'none',
          textAlign: align,
          background: isActive ? '#EBF4FF' : undefined,
          color: isActive ? 'var(--sunny-blue)' : undefined,
          whiteSpace: 'nowrap'
        }}
      >
        {label}{' '}
        <span style={{ opacity: isActive ? 1 : 0.3, fontSize: '0.75rem' }}>
          {isActive ? (sortConfig.dir === 'desc' ? '▼' : '▲') : '⇅'}
        </span>
      </th>
    );
  };

  const handleExport = () => {
    if (!clientsData.items?.length) return;
    const headers = ['ID', 'Cliente', 'Representante', 'UF', 'Faturamento', 'Variacao', 'Frequencia', 'Recencia', 'Status'];
    const rows = clientsData.items.map(c => [
        c.id, c.cliente, c.representante, c.uf, c.faturamento, c.variacao, c.frequencia, c.ultimoPedidoDias, c.scoreLabel
    ]);
    const csvContent = [headers, ...rows].map(e => e.join(";")).join("\n");
    const blob = new Blob(["\ufeff" + csvContent], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `sunny_clientes_${filters.periodo.toLowerCase()}.csv`;
    link.click();
  };

  const handlePeriodChange = (p) => {
    if (p === 'Personalizado') {
      setFilters(prev => ({ ...prev, periodo: p }));
      return;
    }
    const { start, end } = getDynamicDateRange(p);
    setFilters(prev => ({ ...prev, periodo: p, start, end }));
  };

  useEffect(() => {
    fetch(`/api/meta/filtros?gerente=${filters.gerente}&start=${filters.start}&end=${filters.end}`)
      .then(res => res.json())
      .then(setMetaFiltros)
      .catch(console.error);
  }, [filters.gerente, filters.start, filters.end]);

  useEffect(() => { fetchData(); }, [filters, activeTab]);

  const fetchData = () => {
    if (activeTab === 'importacao' || activeTab === 'produtos' || activeTab === 'gestao') {
        setLoading(false);
        return;
    }
    setLoading(true);
    let url = new URL(`/api/${activeTab}`, window.location.origin);
    Object.keys(filters).forEach(key => { if (filters[key] !== undefined) url.searchParams.append(key, filters[key]); });
    fetch(url).then(res => res.json()).then(resData => {
      if (activeTab === 'clientes') setClientsData(resData);
      setLoading(false);
    }).catch(console.error);
  };

  const formatCurrency = (val) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 }).format(val || 0);

  const getStatusColor = (status) => {
    switch (status) {
      case 'Crítico': return 'var(--danger)';
      case 'Atenção': return 'var(--warning)';
      case 'Oportunidade': return 'var(--info)';
      case 'Fora do Ciclo': return 'var(--indigo)';
      default: return 'var(--success)';
    }
  };

  return (
    <div className="sunny-root">
        <aside className="sunny-aside">
            <div className="brand-logo">Sunny</div>
            <nav className="side-nav">
                <button className={`nav-btn ${activeTab==='dashboard'?'active':''}`} onClick={()=>setActiveTab('dashboard')}>📊 Dashboard</button>
                <button className={`nav-btn ${activeTab==='clientes'?'active':''}`} onClick={()=>setActiveTab('clientes')}>👥 Clientes</button>
                <button className={`nav-btn ${activeTab==='produtos'?'active':''}`} onClick={()=>setActiveTab('produtos')}>📦 Produtos</button>
                <button className={`nav-btn ${activeTab==='importacao'?'active':''}`} onClick={()=>setActiveTab('importacao')}>📥 Importação</button>
                <button className={`nav-btn ${activeTab==='gestao'?'active':''}`} onClick={()=>setActiveTab('gestao')}>🛠️ Gestão</button>
                <button className="nav-btn">🏢 Rede</button>
                <button className="nav-btn">👔 Reps</button>
                <button className="nav-btn">🎯 Radar Completo</button>
            </nav>
            <div className="user-box">Tiago Santos<br/><small>Gerente Comercial</small></div>
        </aside>

        <main className="sunny-body">
            {activeTab === 'dashboard' ? (
                <>
                    <header className="main-header">
                        <div className="title-area">
                            <h1>Dashboard</h1>
                            <p>Visão Consolidada de Inteligência</p>
                        </div>
                    </header>
                </>
            ) : activeTab === 'importacao' ? (
                <Importacao 
                    onStartTask={(name) => setBackgroundTask({ name, status: 'processando', progress: 0 })}
                    onFinishTask={(msg) => {
                        setBackgroundTask(null);
                        alert(msg);
                    }}
                    activeTask={backgroundTask}
                />
            ) : activeTab === 'produtos' ? (
                <ModuloProdutos />
            ) : activeTab === 'gestao' ? (
                <GestaoDados />
            ) : (
                <>
                    <header className="main-header">
                        <div className="title-area">
                            <h1>Gestão de Clientes</h1>
                            <p>Análise técnica e priorização de base</p>
                        </div>
                        
                        {/* Camada de Contexto (Macro-Filtros) */}
                        <div className="context-layer">
                            <div className={`context-group ${filters.periodo === 'Personalizado' ? 'expanded' : ''}`}>
                                <span className="ctx-icon">📅</span>
                                <select value={filters.periodo} onChange={e => handlePeriodChange(e.target.value)}>
                                    <option>Mês Atual</option>
                                    <option>Mês Anterior</option>
                                    <option>Últimos 3 Meses</option>
                                    <option>Ano Atual</option>
                                    <option>Personalizado</option>
                                </select>
                                {filters.periodo === 'Personalizado' && (
                                    <div className="date-inputs">
                                        <input type="date" value={filters.start} onChange={e => setFilters({...filters, start: e.target.value})} />
                                        <span>até</span>
                                        <input type="date" value={filters.end} onChange={e => setFilters({...filters, end: e.target.value})} />
                                    </div>
                                )}
                                <div style={{marginLeft:'12px', paddingLeft:'12px', borderLeft:'1px solid #E2E8F0', display:'flex', alignItems:'center', gap:'8px'}}>
                                    <input type="checkbox" id="chkComp" checked={filters.compare} onChange={e => setFilters({...filters, compare: e.target.checked})} />
                                    <label htmlFor="chkComp" style={{fontSize:'0.7rem', fontWeight:'700', color:'var(--sunny-blue)', cursor:'pointer'}}>Comparar YoY</label>
                                </div>
                            </div>
                            <div className="context-group">
                                <span className="ctx-icon">🏢</span>
                                <select value={filters.gerente} onChange={e => setFilters({...filters, gerente: e.target.value})}>
                                    <option value="Todos">Todos os Gerentes</option>
                                    {metaFiltros?.gerentes?.map(g => <option key={g.id} value={g.id}>{g.nome}</option>)}
                                </select>
                            </div>
                            <div className="date-range-badge">
                                {filters.start.split('-').reverse().join('/')} - {filters.end.split('-').reverse().join('/')}
                            </div>
                        </div>
                    </header>

                    {/* Camada de Operação (Barra de Filtros Principal) */}
                    <div className="filters-bar-v2">
                        <div className="search-global-wrap">
                            <span className="lupa">🔍</span>
                            <input 
                                type="text" 
                                placeholder="Buscar cliente por nome, CNPJ ou EAN..." 
                                value={filters.search || ''}
                                onChange={e => setFilters({...filters, search: e.target.value})}
                            />
                        </div>

                        <div className={`filter-select-v2 ${filters.uf !== 'Todos' ? 'active' : ''}`}>
                            <span className="select-icon">📍</span>
                            <select value={filters.uf} onChange={e => setFilters({...filters, uf: e.target.value})}>
                                <option value="Todos">Todas as Regiões</option>
                                {metaFiltros.ufs?.map(uf => <option key={uf} value={uf}>{uf}</option>)}
                            </select>
                            {filters.uf !== 'Todos' && <span className="filter-badge">1</span>}
                        </div>

                        <div className={`filter-select-v2 vendedor ${filters.vendedor !== 'Todos' ? 'active' : ''}`}>
                            <span className="select-icon">👤</span>
                            <select value={filters.vendedor} onChange={e => setFilters({...filters, vendedor: e.target.value})}>
                                <option value="Todos">Todos os Vendedores</option>
                                {metaFiltros.vendedores?.map(v => (
                                    <option key={v.nome} value={v.nome}>{v.nome} ({v.total})</option>
                                ))}
                            </select>
                            {filters.vendedor !== 'Todos' && <span className="filter-badge">1</span>}
                        </div>

                        <div className={`filter-select-v2 ${filters.perfil && filters.perfil !== 'Todos' ? 'active' : ''}`}>
                            <span className="select-icon">🏢</span>
                            <select value={filters.perfil || 'Todos'} onChange={e => setFilters({...filters, perfil: e.target.value})}>
                                <option value="Todos">Todos os Perfis</option>
                                <option value="Shopping">Shopping</option>
                                <option value="Rua">Rua</option>
                                <option value="Magazine">Magazine</option>
                                <option value="Especializada">Especializada</option>
                            </select>
                            {filters.perfil && filters.perfil !== 'Todos' && <span className="filter-badge">1</span>}
                        </div>

                        <div style={{flex: 1, display:'flex', justifyContent:'flex-end', gap:'12px', alignItems:'center'}}>
                            <button className="btn-clear-minimal" onClick={() => {
                                const d = getDynamicDateRange('Mês Atual');
                                setFilters({search: '', uf:'Todos', vendedor:'Todos', perfil:'Todos', status:'Todos', gerente:'Todos', periodo:'Mês Atual', start: d.start, end: d.end, compare: false});
                            }}>
                                Limpar Filtros
                            </button>
                            <button className="btn-sec" onClick={handleExport} style={{padding:'8px 12px', fontSize:'0.75rem'}}>📥 Exportar</button>
                        </div>
                    </div>

                    <div style={{display:'flex', alignItems:'center', gap:'16px', marginBottom:'24px'}}>
                        <div className="insight-bar" style={{margin:0, flex:1}}>
                            <div className="insight-pill" onClick={() => setFilters({...filters, uf: 'SP', status: 'Crítico'})}>
                                <span>🚩</span> <strong>Queda em SP</strong> (-15%)
                            </div>
                            <div className="insight-pill" onClick={() => setFilters({...filters, status: 'Crítico'})}>
                                <span>⚠️</span> <strong>Recência Crítica</strong>
                            </div>
                            <div className="insight-pill" onClick={() => setFilters({...filters, search: 'SUPERLEGAL'})}>
                                <span>🏢</span> <strong>Rede:</strong> Superlegal
                            </div>
                        </div>
                        
                        <div className="status-pills" style={{border:0, margin:0, padding:0}}>
                            {['Todos', 'Crítico', 'Alerta', 'Saudável'].map(st => (
                                <div 
                                    key={st} 
                                    className={`status-pill ${st.toLowerCase()} ${filters.status === st ? 'active' : ''}`}
                                    onClick={() => setFilters({...filters, status: st})}
                                >
                                    {st}
                                </div>
                            ))}
                        </div>
                    </div>

                        <KPIStatsCards kpis={clientsData.kpis} compare={filters.compare} />

                    <div className="table-card-v21">
                        <table className="analy-table">
                            <thead>
                                <tr>
                                    <SortTh col="cliente" label="Identificação do Cliente" />
                                    <SortTh col="faturamento" label="Faturamento" />
                                    <SortTh col="variacao" label="Desvio (%)" />
                                    <SortTh col="ciclo" label="Frequência" />
                                    <SortTh col="ultimoPedidoDias" label="Recência" />
                                    <SortTh col="score" label="Health Score" />
                                    <SortTh col="oportunidade" label="Oportunidade Estimada (R$)" />
                                    <th>Ações</th>
                                </tr>
                            </thead>
                            <tbody>
                                {loading ? <tr><td colSpan="8" style={{padding:'60px',textAlign:'center', color:'var(--text-muted)'}}>Processando indicadores analíticos...</td></tr> : 
                                 sortedItems.map(c => (
                                    <tr key={c.unique_id}>
                                        <td>
                                            <div className="cli-cell">
                                                <span className="status-dot" style={{ background: getStatusColor(c.status) }}></span>
                                                <strong>{c.cliente}</strong><br/>
                                                <small>ID: {c.id} | {c.representante} | {c.uf} | <strong style={{color:'var(--sunny-blue)'}}>{c.perfil}</strong></small>
                                            </div>
                                        </td>
                                        <td>
                                            <div className="fat-cell" style={{fontWeight:'600'}}>{formatCurrency(c.faturamento)}</div>
                                            {filters.compare && <div style={{fontSize:'0.65rem', color:'var(--text-muted)'}}>{formatCurrency(c.fatComp)} (YoY)</div>}
                                        </td>
                                        <td className={parseFloat(c.variacao) < 0 ? 'txt-danger' : 'txt-success'}>
                                            <span style={{fontWeight:'700'}}>{parseFloat(c.variacao) > 0 ? '+' : ''}{c.variacao}%</span>
                                        </td>
                                        <td><div className="freq-cell" style={{fontSize:'0.8rem'}}>{c.frequencia}<br/><small>Ciclo: {c.ciclo}d</small></div></td>
                                        <td><div style={{fontWeight:'700', color: 'var(--text-main)'}}>{c.ultimoPedidoDias} dias</div></td>
                                        <td><ScoreBadge score={c.score} label={c.scoreLabel} /></td>
                                        <td><div style={{color: 'var(--text-main)', fontWeight: '700'}}>{formatCurrency(c.oportunidade)}</div></td>
                                        <td><button className="btn-view" onClick={() => setSelectedClient(c)}>Analisar 360°</button></td>
                                    </tr>
                                ))}
                            </tbody>
                        </table>
                    </div>
                </>
            )}
        </main>

        {selectedClient && (
            <View360 
                client={selectedClient} 
                onBack={() => setSelectedClient(null)} 
                globalFilters={filters}
            />
        )}

        {/* Indicador Global de Segundo Plano */}
        {backgroundTask && (
            <div className="global-task-indicator fade-in">
                <div className="spinner-small"></div>
                <div style={{flex:1}}>
                    <div style={{fontSize:'0.85rem', fontWeight:'900', color:'white', marginBottom:'2px'}}>
                        {backgroundTask.name}
                    </div>
                    <div className="progress-bar-loading"></div>
                    <div className="progress-text" style={{color:'rgba(255,255,255,0.8)', textAlign:'left'}}>
                        Processando em segundo plano...
                    </div>
                </div>
            </div>
        )}
    </div>
  );
};

export default App;
