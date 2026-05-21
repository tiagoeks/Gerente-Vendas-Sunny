const express = require('express');
const cors = require('cors');
const { Pool } = require('pg');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const app = express();
const port = process.env.PORT || 3003;
const upload = multer({ dest: 'uploads/' });

app.use(cors());
app.use(express.json());

// Log GLOBAL para rastrear qualquer requisição
app.use((req, res, next) => {
    console.log(`[${new Date().toLocaleTimeString()}] ${req.method} ${req.url}`);
    next();
});

const dbPath = path.resolve(__dirname, 'sunny.sqlite');
let db = null;
let pgPool = null;
const usePostgres = !!process.env.DATABASE_URL;

if (usePostgres) {
    console.log("Conectando ao banco de dados PostgreSQL...");
    pgPool = new Pool({
        connectionString: process.env.DATABASE_URL,
        ssl: {
            rejectUnauthorized: false
        }
    });
} else {
    console.log("Conectando ao banco de dados SQLite local...");
    const sqlite3 = require('sqlite3').verbose();
    db = new sqlite3.Database(dbPath);
    // Ativa Modo Multitarefa (WAL) - Permite ler enquanto escreve (apenas SQLite)
    db.run('PRAGMA journal_mode = WAL');
}

const query = (sql, params = []) => {
    return new Promise((resolve, reject) => {
        if (usePostgres) {
            // Converte placeholders "?" do SQLite para "$1", "$2" etc do Postgres
            let index = 1;
            const pgSql = sql.replace(/\?/g, () => `$${index++}`);
            pgPool.query(pgSql, params, (err, res) => {
                if (err) reject(err);
                else resolve(res.rows);
            });
        } else {
            db.all(sql, params, (err, rows) => {
                if (err) reject(err);
                else resolve(rows);
            });
        }
    });
};

// Inicializador de tabelas (para Postgres e SQLite)
const initDb = async () => {
    if (usePostgres) {
        // Inicializar tabelas no PostgreSQL
        await query(`
            CREATE TABLE IF NOT EXISTS vendas (
                id SERIAL PRIMARY KEY,
                num_docto TEXT,
                emissao TEXT,
                mes TEXT,
                uf TEXT,
                status TEXT,
                cliente_id TEXT,
                loja TEXT,
                nome_cliente TEXT,
                cnpj TEXT,
                pedido_cliente TEXT,
                quantidade REAL,
                valor_unitario REAL,
                valor_total REAL,
                almox TEXT,
                produto_id TEXT,
                descricao_produto TEXT,
                vendedor_id TEXT,
                nome_vendedor TEXT,
                gerente_id TEXT,
                marca TEXT,
                ean TEXT,
                natureza_operacao TEXT
            )
        `);
        await query(`
            CREATE TABLE IF NOT EXISTS estoque (
                id SERIAL PRIMARY KEY,
                cod_produto TEXT UNIQUE,
                descricao TEXT,
                unidade TEXT,
                marca TEXT,
                saldo REAL,
                pack TEXT,
                sortimento TEXT,
                pv REAL,
                pdv REAL,
                ean TEXT,
                previsao TEXT,
                image_url TEXT,
                categoria TEXT
            )
        `);
        
        // Garante que pv, pdv e saldo sejam tipos flutuantes (DOUBLE PRECISION) se a tabela foi criada anteriormente como INTEGER
        try {
            await query(`ALTER TABLE estoque ALTER COLUMN pv TYPE DOUBLE PRECISION USING pv::double precision`);
            await query(`ALTER TABLE estoque ALTER COLUMN pdv TYPE DOUBLE PRECISION USING pdv::double precision`);
            await query(`ALTER TABLE estoque ALTER COLUMN saldo TYPE DOUBLE PRECISION USING saldo::double precision`);
            console.log('[Db Init] Colunas pv, pdv e saldo da tabela estoque verificadas/migradas para DOUBLE PRECISION.');
        } catch (alterErr) {
            console.log('[Db Init] Não foi necessário alterar colunas da tabela estoque:', alterErr.message);
        }

        // Garante que a coluna natureza_operacao exista na tabela vendas do Postgres
        try {
            await query(`ALTER TABLE vendas ADD COLUMN IF NOT EXISTS natureza_operacao TEXT`);
            console.log('[Db Init] Coluna natureza_operacao na tabela vendas verificada/adicionada no PostgreSQL.');
        } catch (alterErr) {
            console.log('[Db Init] Não foi necessário alterar coluna natureza_operacao no PostgreSQL:', alterErr.message);
        }

        await query(`
            CREATE TABLE IF NOT EXISTS clientes_perfil (
                cliente_id TEXT PRIMARY KEY,
                nome_cliente TEXT,
                perfil TEXT
            )
        `);
        await query(`
            CREATE TABLE IF NOT EXISTS marcas_mestre (
                id SERIAL PRIMARY KEY,
                nome TEXT UNIQUE
            )
        `);
        await query(`
            CREATE TABLE IF NOT EXISTS marcas_legado (
                id SERIAL PRIMARY KEY,
                marca TEXT,
                produto_id TEXT
            )
        `);
        // Criar índices para performance se não existirem
        await query(`CREATE INDEX IF NOT EXISTS idx_cliente ON vendas(cliente_id)`);
        await query(`CREATE INDEX IF NOT EXISTS idx_vendedor ON vendas(vendedor_id)`);
        await query(`CREATE INDEX IF NOT EXISTS idx_status ON vendas(status)`);
        await query(`CREATE INDEX IF NOT EXISTS idx_emissao ON vendas(emissao)`);
        await query(`CREATE INDEX IF NOT EXISTS idx_vendas_doc_prod ON vendas(num_docto, produto_id)`);
    } else {
        // Tabela de Perfil de Clientes no SQLite
        db.run(`CREATE TABLE IF NOT EXISTS clientes_perfil (
            cliente_id TEXT PRIMARY KEY,
            nome_cliente TEXT,
            perfil TEXT
        )`);
        // Migrações no SQLite
        db.run("ALTER TABLE vendas ADD COLUMN valor_unitario REAL", () => {});
        db.run("ALTER TABLE vendas ADD COLUMN natureza_operacao TEXT", () => {});
        db.run("ALTER TABLE estoque ADD COLUMN image_url TEXT", () => {});
        db.run("ALTER TABLE estoque ADD COLUMN categoria TEXT", () => {});
        db.run("CREATE INDEX IF NOT EXISTS idx_vendas_doc_prod ON vendas(num_docto, produto_id)", () => {});
    }
};

initDb().catch(console.error);

// --- HELPER: SANITIZAÇÃO UNIVERSAL ---
const sanitize = (val) => {
    if (val === undefined || val === null) return '';
    // Remove espaços e zeros à esquerda
    return String(val).trim().replace(/^0+/, '');
};

const toTitleCaseClean = (str) => {
    if (!str) return '';
    return String(str)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "") // remove acentos
        .replace(/\s+/g, " ") // remove espaços múltiplos
        .trim()
        .toLowerCase()
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
};

const parseMoney = (val) => {
    if (val === undefined || val === null || val === '') return 0;
    if (typeof val === 'number') return val;
    let s = String(val).replace('R$', '').replace(/\s/g, '').trim();
    if (s.includes(',') && s.includes('.')) {
        return parseFloat(s.replace(/\./g, '').replace(',', '.'));
    } else if (s.includes(',')) {
        return parseFloat(s.replace(',', '.'));
    }
    return parseFloat(s) || 0;
};

const excelDate = (val) => {
    if (!val) return '';
    if (!isNaN(val) && parseFloat(val) > 30000) {
        const date = new Date(Math.round((val - 25569) * 86400 * 1000));
        return date.toISOString().split('T')[0];
    }
    return String(val).trim();
};

const BASE_DATE = new Date().toISOString().split('T')[0];
const FIXED_WHERE = "(status = '5' OR status = '6') AND almox = '20'";
const CLIENT_KEY = "cliente_id || ' - ' || nome_cliente";

const getDefaultDates = (start, end) => {
    if (start && end) return { dStart: start, dEnd: end };
    
    const today = new Date();
    const y = today.getFullYear();
    const m = today.getMonth(); // 0-indexed
    const pad = (n) => String(n).padStart(2, '0');
    
    const dStart = start || `${y}-${pad(m + 1)}-01`;
    const dEnd = end || `${y}-${pad(m + 1)}-${pad(new Date(y, m + 1, 0).getDate())}`;
    return { dStart, dEnd };
};

function getWhereAndParams(filters) {
    const { periodo, gerente, representante, cliente } = filters;
    let whereClause = `WHERE ${FIXED_WHERE}`; 
    let params = [];

    const today = new Date();
    const y = today.getFullYear();
    const m = today.getMonth(); // 0-indexed
    const pad = (n) => String(n).padStart(2, '0');

    if (periodo === 'Mês Atual') {
        const startOfMonth = `${y}-${pad(m + 1)}-01`;
        whereClause += ` AND emissao >= '${startOfMonth}'`;
    } else if (periodo === 'Último Mês' || periodo === 'Mês Anterior') {
        const prevMonthDate = new Date(y, m - 1, 1);
        const prevY = prevMonthDate.getFullYear();
        const prevM = prevMonthDate.getMonth();
        const startOfPrevMonth = `${prevY}-${pad(prevM + 1)}-01`;
        const endOfPrevMonth = `${prevY}-${pad(prevM + 1)}-${pad(new Date(prevY, prevM + 1, 0).getDate())}`;
        whereClause += ` AND emissao BETWEEN '${startOfPrevMonth}' AND '${endOfPrevMonth}'`;
    }

    if (gerente && gerente !== 'Todos') { whereClause += " AND gerente_id = ?"; params.push(gerente); }
    if (representante && representante !== 'Todos') { whereClause += " AND vendedor_id = ?"; params.push(representante); }
    if (cliente && cliente !== 'Todos') { whereClause += " AND cliente_id = ?"; params.push(cliente); }

    return { whereClause, params };
}

// Endpoint para metadados de filtros (Vendedores com contagem e UFs)
app.get('/api/meta/filtros', async (req, res) => {
    try {
        const { gerente, start, end } = req.query;
        const { dStart, dEnd } = getDefaultDates(start, end);
        let vWhere = `emissao BETWEEN '${dStart}' AND '${dEnd}'`;
        
        if (gerente && gerente !== 'Todos') {
            vWhere += ` AND gerente_id = ?`;
        }
        const vParams = gerente && gerente !== 'Todos' ? [gerente] : [];

        const gerentes = await query(`SELECT DISTINCT gerente_id as id FROM vendas WHERE gerente_id != '' ORDER BY id`);
        const ufs = await query(`SELECT DISTINCT uf FROM vendas WHERE ${vWhere} ORDER BY uf`, vParams);
        const vendedores = await query(`
            SELECT DISTINCT nome_vendedor as nome, COUNT(DISTINCT ${CLIENT_KEY}) as total 
            FROM vendas 
            WHERE ${vWhere} 
            GROUP BY nome_vendedor 
            ORDER BY nome
        `, vParams);

        res.json({ 
            vendedores, 
            ufs: ufs.map(u => u.uf),
            gerentes: gerentes.map(g => ({ id: g.id, nome: `Gerente ${g.id}` }))
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/clientes', async (req, res) => {
    try {
        const { search, uf, vendedor, status, gerente, start, end, compare } = req.query;
        const baseDate = new Date(BASE_DATE);
        
        const { dStart, dEnd } = getDefaultDates(start, end);

        // Determinar Datas de Comparação (YoY)
        const compStart = dStart.replace(/^(\d{4})/, (match, p1) => String(parseInt(p1) - 1));
        const compEnd = dEnd.replace(/^(\d{4})/, (match, p1) => String(parseInt(p1) - 1));

        // Build WHERE using v. prefix directly — safe, no fragile .replace()
        let where = `v.emissao BETWEEN ? AND ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20'`;
        const params = [dStart, dEnd];

        if (search) {
            where += ` AND (v.nome_cliente LIKE ? OR v.cnpj LIKE ?)`;
            params.push(`%${search}%`, `%${search}%`);
        }
        if (uf && uf !== 'Todos') {
            where += ` AND v.uf = ?`;
            params.push(uf);
        }
        if (vendedor && vendedor !== 'Todos') {
            where += ` AND v.nome_vendedor = ?`;
            params.push(vendedor);
        }
        if (gerente && gerente !== 'Todos') {
            where += ` AND v.gerente_id = ?`;
            params.push(gerente);
        }
        
        let perfilFilter = '';
        const { perfil } = req.query;
        if (perfil && perfil !== 'Todos') {
            perfilFilter = `AND COALESCE(cp.perfil, 'Rua') = ?`;
            params.push(perfil);
        }

        const current = await query(`
            SELECT 
                (v.cliente_id || ' - ' || v.nome_cliente) as unique_id,
                v.cliente_id, 
                v.nome_cliente, 
                MAX(v.nome_vendedor) as representante, 
                MAX(v.uf) as uf,
                MAX(cp.perfil) as perfil,
                SUM(v.valor_total) as faturamento, 
                COUNT(DISTINCT v.num_docto) as pedidos, 
                SUM(v.quantidade) as itens_totais, 
                MAX(v.emissao) as ultima_compra, 
                MAX(v.cnpj) as cnpj
            FROM vendas v
            LEFT JOIN clientes_perfil cp ON cp.cliente_id = v.cliente_id
            WHERE ${where} ${perfilFilter}
            GROUP BY v.cliente_id, v.nome_cliente
        `, params);

        const mediasRows = await query(`
            SELECT 
                v.uf, 
                COALESCE(cp.perfil, 'Rua') as perfil, 
                SUM(v.valor_total) / COUNT(DISTINCT v.cliente_id) as media_faturamento
            FROM vendas v
            LEFT JOIN clientes_perfil cp ON cp.cliente_id = v.cliente_id
            WHERE v.emissao BETWEEN ? AND ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20'
            GROUP BY v.uf, COALESCE(cp.perfil, 'Rua')
        `, [dStart, dEnd]);
        
        const ufAverages = await query(`
            SELECT v.uf, SUM(v.valor_total) / COUNT(DISTINCT v.cliente_id) as media
            FROM vendas v
            WHERE v.emissao BETWEEN ? AND ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20'
            GROUP BY v.uf
        `, [dStart, dEnd]);

        const mediasMap = {};
        mediasRows.forEach(r => mediasMap[r.uf + '_' + r.perfil] = r.media_faturamento);
        const ufAvgMap = Object.fromEntries(ufAverages.map(u => [u.uf, u.media]));
        
        // Query de Comparação (Apenas se solicitado)
        let compMap = {};
        if (compare === 'true') {
            const previous = await query(`
                SELECT (cliente_id || ' - ' || nome_cliente) as unique_id, SUM(valor_total) as faturamento 
                FROM vendas 
                WHERE emissao BETWEEN ? AND ? AND (status = '5' OR status = '6') AND almox = '20'
                GROUP BY cliente_id, nome_cliente
            `, [compStart, compEnd]);
            compMap = Object.fromEntries(previous.map(p => [p.unique_id, p.faturamento]));
        }

        const history = await query(`SELECT (cliente_id || ' - ' || nome_cliente) as unique_id, COUNT(DISTINCT num_docto) as t_ped, MIN(emissao) as start FROM vendas WHERE (status = '5' OR status = '6') AND almox = '20' GROUP BY cliente_id, nome_cliente`);
        const hMap = Object.fromEntries(history.map(h => [h.unique_id, { ...h, t_ped: parseInt(h.t_ped) || 0 }]));

        let results = current.map(c => {
            const fatComp = compMap[c.unique_id] || 0;
            const hist = hMap[c.unique_id] || { t_ped: c.pedidos, start: c.ultima_compra };
            const d1 = new Date(hist.start);
            const d2 = new Date(BASE_DATE);
            const diffMonths = Math.max(1, (d2.getFullYear() - d1.getFullYear()) * 12 + (d2.getMonth() - d1.getMonth() + 1));
            const freq = hist.t_ped / diffMonths;
            const cicloDias = Math.max(1, Math.round(30.4 / Math.max(0.1, freq)));
            const dRec = Math.floor((d2 - new Date(c.ultima_compra))/(1000*60*60*24));
            
            // Variação YoY ou Periodo Anterior
            const variacao = fatComp > 0 ? ((c.faturamento - fatComp)/fatComp*100) : 0;

            let sRec = 10 * Math.exp(-dRec / (cicloDias * 1.5));
            let sGrowth = 5 + (variacao / 10);
            sGrowth = Math.min(10, Math.max(0, sGrowth));
            const finalScore = (sRec * 0.6) + (sGrowth * 0.4);
            
            let category = "Saudável";
            if (variacao < -25 || dRec > cicloDias * 2) category = "Crítico";
            else if (dRec > cicloDias) category = "Atenção";
            else if (variacao > 15) category = "Oportunidade";

            const perfilFinal = c.perfil || 'Rua';
            const mediaBenchmark = mediasMap[c.uf + '_' + perfilFinal] || ufAvgMap[c.uf] || 0;
            const oportunidade = Math.max(0, mediaBenchmark - c.faturamento);

            return {
                id: c.cliente_id, unique_id: c.unique_id, cnpj: c.cnpj, cliente: c.nome_cliente, representante: c.representante, uf: c.uf,
                perfil: perfilFinal, oportunidade, benchmark: mediaBenchmark,
                faturamento: parseFloat(c.faturamento) || 0,
                fatComp: parseFloat(fatComp) || 0,
                pedidos: parseInt(c.pedidos) || 0,
                itens: parseFloat(c.itens_totais) || 0,
                variacao: variacao ? variacao.toFixed(1) : '0',
                frequencia: `${freq.toFixed(1)}/mês`, ciclo: cicloDias, ultimoPedidoDias: dRec,
                status: category, 
                score: finalScore, 
                scoreLabel: finalScore >= 8 ? 'Excelente' : finalScore >= 6.5 ? 'Saudável' : finalScore >= 4 ? 'Atenção' : 'Crítico',
                potencial: oportunidade
            };
        });

        if (status && status !== 'Todos') {
            const mappedStatus = status.toLowerCase() === 'alerta' ? 'atenção' : status.toLowerCase();
            results = results.filter(r => r.scoreLabel.toLowerCase() === mappedStatus);
        }

        // Cálculo de KPIs Agregados (Dinâmico conforme filtros)
        const kpis = {
            faturamento: results.reduce((a, b) => a + b.faturamento, 0),
            fatComp: results.reduce((a, b) => a + b.fatComp, 0),
            pedidos: results.reduce((a, b) => a + b.pedidos, 0),
            pedidosComp: compare === 'true' ? results.reduce((a, b) => a + (compMap[b.unique_id] ? b.pedidos : 0), 0) : 0, // Simplificado
            itens: results.reduce((a, b) => a + b.itens, 0),
        };

        // Adição de Ticket Médio e Variações
        kpis.ticketMedio = kpis.pedidos > 0 ? kpis.faturamento / kpis.pedidos : 0;
        kpis.ticketComp = kpis.pedidosComp > 0 ? kpis.fatComp / kpis.pedidosComp : 0;
        
        kpis.varFat = kpis.fatComp > 0 ? ((kpis.faturamento - kpis.fatComp)/kpis.fatComp*100).toFixed(1) : 0;
        kpis.varPed = kpis.pedidosComp > 0 ? ((kpis.pedidos - kpis.pedidosComp)/kpis.pedidosComp*100).toFixed(1) : 0;
        kpis.varTicket = kpis.ticketComp > 0 ? ((kpis.ticketMedio - kpis.ticketComp)/kpis.ticketComp*100).toFixed(1) : 0;

        results.sort((a, b) => b.potencial - a.potencial);
        res.json({ 
            items: results, 
            kpis,
            grupos: { 
                criticos: results.filter(r => r.status === 'Crítico').length,
                alerta: results.filter(r => r.status === 'Atenção').length,
                saudavel: results.filter(r => r.status === 'Saudável').length,
                oportunidades: results.filter(r => r.status === 'Oportunidade').length
            }
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/clientes/:id', async (req, res) => {
    try {
        const { id } = req.params; // CNPJ
        const { start, end } = req.query;
        const { dStart, dEnd } = getDefaultDates(start, end);
        
        console.log(`[Busca 360] Analisando indicadores para CNPJ: "${id}" | Filtro: ${dStart} a ${dEnd}`);

        const dateFilter = `AND emissao BETWEEN '${dStart}' AND '${dEnd}'`;
        const dateFilterStrict = `AND emissao BETWEEN '${dStart}' AND '${dEnd}' AND (status = '5' OR status = '6') AND almox = '20'`;

        // Obter Perfil e UF do Cliente
        const clientInfo = await query(`
            SELECT v.uf, cp.perfil 
            FROM vendas v 
            LEFT JOIN clientes_perfil cp ON cp.cliente_id = v.cliente_id 
            WHERE v.cnpj = ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20' LIMIT 1
        `, [id]);
        const clientUf = clientInfo[0]?.uf || 'SP';
        const clientPerfil = clientInfo[0]?.perfil || 'Rua';

        // 1. Dados Vitalícios (Lifetime) - Ignora filtros de data, mas respeita status
        const lifetime = await query(`
            SELECT 
                SUM(valor_total) as total,
                MIN(emissao) as primeira_compra,
                COUNT(DISTINCT num_docto) as total_pedidos,
                COUNT(DISTINCT produto_id) as skus_comprados
            FROM vendas WHERE cnpj = ? AND (status = '5' OR status = '6') AND almox = '20'
        `, [id]);

        const activeSkusQuery = await query(`SELECT COUNT(cod_produto) as total FROM estoque WHERE saldo > 0`);
        const totalActiveSkus = parseInt(activeSkusQuery[0]?.total) || 1;
        const skusComprados = parseInt(lifetime[0]?.skus_comprados) || 0;
        const coberturaMix = (skusComprados / totalActiveSkus) * 100;

        // 2. Faturamento no Período Selecionado (apenas pedidos atendidos)
        const totalPeriodo = await query(`SELECT SUM(valor_total) as total FROM vendas WHERE cnpj = ? ${dateFilterStrict}`, [id]);
        const fatTotal = parseFloat(totalPeriodo[0]?.total) || 1;

        // 2.1 Benchmark UF + Perfil
        const benchmarkQuery = await query(`
            SELECT SUM(v.valor_total) / NULLIF(COUNT(DISTINCT v.cliente_id), 0) as media
            FROM vendas v
            LEFT JOIN clientes_perfil cp ON cp.cliente_id = v.cliente_id
            WHERE v.uf = ? AND COALESCE(cp.perfil, 'Rua') = ? ${dateFilterStrict.replace(/status/g, 'v.status').replace(/almox/g, 'v.almox').replace(/emissao/g, 'v.emissao')}
        `, [clientUf, clientPerfil]);
        const benchmarkValor = parseFloat(benchmarkQuery[0]?.media) || 0;

        // 3. Histórico de Pedidos (Rápido com Índice)
        const pedidosBase = await query(`
            SELECT num_docto, MAX(emissao) as emissao, SUM(valor_total) as total, COUNT(DISTINCT descricao_produto) as skus, SUM(quantidade) as qtd_total, MAX(status) as status_id
            FROM vendas WHERE cnpj = ? AND (status = '5' OR status = '6') AND almox = '20'
            GROUP BY num_docto ORDER BY MAX(emissao) DESC
        `, [id]);

        // 4. Mix Completo (RESPEITA FILTRO)
        const produtosMix = await query(`
            SELECT 
                v.ean, v.descricao_produto as nome, COALESCE(e.marca, v.marca) as marca, v.produto_id,
                SUM(v.quantidade) as qtd, SUM(v.valor_total) as total,
                (SUM(v.valor_total) / NULLIF(SUM(v.quantidade), 0)) as preco_medio,
                (SELECT COALESCE(v2.valor_unitario, v2.valor_total / NULLIF(v2.quantidade, 0), 0) FROM vendas v2 
                 WHERE v2.cnpj = ? AND v2.produto_id = v.produto_id 
                 AND (v2.status = '5' OR v2.status = '6') AND v2.almox = '20'
                 ORDER BY v2.emissao DESC LIMIT 1) as preco_ultima,
                MAX(v.emissao) as ultima_data,
                MIN(v.emissao) as primeira_data,
                COUNT(DISTINCT v.num_docto) as total_pedidos_sku,
                ((SUM(v.valor_total) * 100.0) / ?) as participacao,
                MAX(e.saldo) as saldo, MAX(e.pv) as pv, MAX(e.pdv) as pdv,
                MAX(e.previsao) as previsao, MAX(e.pack) as pack,
                MAX(e.sortimento) as sortimento, MAX(e.image_url) as image_url
            FROM vendas v
            LEFT JOIN estoque e ON v.produto_id = e.cod_produto
            WHERE v.cnpj = ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20' ${dateFilter}
            GROUP BY v.ean, v.descricao_produto, v.produto_id, COALESCE(e.marca, v.marca)
            ORDER BY total DESC
        `, [id, fatTotal, id]);

        // 5. Oportunidades (Gap de Mercado - Itens mais vendidos da Sunny Nacionalmente que o cliente NÃO tem no período)
        const marketGap = await query(`
            SELECT 
                v.produto_id, v.ean, v.descricao_produto as nome, COALESCE(e.marca, v.marca) as marca, 
                SUM(v.valor_total) as total_geral,
                MAX(e.saldo) as saldo, MAX(e.pv) as pv, MAX(e.pdv) as pdv,
                MAX(e.previsao) as previsao, MAX(e.pack) as pack, MAX(e.sortimento) as sortimento,
                RANK() OVER (ORDER BY SUM(v.valor_total) DESC) as rank_nacional,
                MAX(e.image_url) as image_url
            FROM vendas v
            LEFT JOIN estoque e ON v.produto_id = e.cod_produto
            WHERE e.saldo > 0 AND (v.status = '5' OR v.status = '6') AND v.almox = '20' ${dateFilter.replace(/emissao/g, 'v.emissao')}
            GROUP BY v.produto_id, v.ean, v.descricao_produto, COALESCE(e.marca, v.marca)
            HAVING SUM(CASE WHEN v.cnpj = ? THEN 1 ELSE 0 END) = 0
            ORDER BY total_geral DESC LIMIT 50
        `, [id]);


        // 6. Share de Marcas e Timeline (RESPEITA FILTRO)
        const [brandShare, timelineRaw] = await Promise.all([
            query(`
                SELECT e.marca, SUM(v.valor_total) as total 
                FROM vendas v
                LEFT JOIN estoque e ON v.produto_id = e.cod_produto
                WHERE v.cnpj = ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20' ${dateFilter}
                GROUP BY e.marca 
                ORDER BY total DESC
            `, [id]),
            query(`
                SELECT mes, SUM(atual) as atual, SUM(anterior) as anterior
                FROM (
                    SELECT 
                        LPAD(EXTRACT(MONTH FROM emissao::date)::text, 2, '0') as mes,
                        CASE WHEN emissao >= '${new Date().getFullYear()}-01-01' THEN valor_total ELSE 0 END as atual,
                        CASE WHEN emissao >= '${new Date().getFullYear() - 1}-01-01' AND emissao < '${new Date().getFullYear()}-01-01' THEN valor_total ELSE 0 END as anterior
                    FROM vendas 
                    WHERE cnpj = ? AND (status = '5' OR status = '6') AND almox = '20'
                ) t
                GROUP BY mes ORDER BY mes
            `, [id])
        ]);

        const pedidosFull = pedidosBase.map(p => {
            let statusText = p.status_id;
            if (p.status_id == 5 || p.status_id == 6) statusText = "Atendido";
            else if (p.status_id == 1) statusText = "Pendente";
            return { id: p.num_docto, data: p.emissao, valor: p.total, itens: p.skus, qtdTotal: p.qtd_total, status: statusText || 'Atendido' };
        });

        res.json({
            records: { 
                ultimoPedido: pedidosBase[0] || { total: 0, emissao: '-' }, 
                maiorPedido: [...pedidosBase].sort((a, b) => b.total - a.total)[0] || { total: 0, emissao: '-' } 
            },
            lifetime: lifetime[0],
            coberturaMix,
            benchmark: {
                valor: benchmarkValor,
                uf: clientUf,
                perfil: clientPerfil
            },
            faturamentoPeriodo: fatTotal,
            topProdutos: produtosMix,
            marketGap,
            brandShare,
            timeline: timelineRaw,
            pedidos: pedidosFull
        });
    } catch (err) { 
        console.error(`[Busca 360] Erro fatal:`, err);
        res.status(500).json({ error: err.message }); 
    }
});

app.listen(port, () => {
    console.log(`Sunny Backend V4.0 (Analytics 360 + Import) na porta ${port}`);
});

// --- MÓDULO DE IMPORTAÇÃO REFINADO ---

app.post('/api/import/estoque', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Arquivo não enviado");
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);
        
        console.log(`[Import Estoque] Processando ${data.length} linhas.`);
        if (data.length > 0) console.log(`[Import Estoque] Exemplo de colunas encontradas:`, Object.keys(data[0]));

        let count = 0;
        let newBrandsCount = 0;

        for (const row of data) {
            // Busca dinâmica de chaves (case-insensitive e flexível)
            const getVal = (keys) => {
                const foundKey = Object.keys(row).find(k => keys.some(target => k.toLowerCase().trim().includes(target.toLowerCase())));
                return foundKey ? row[foundKey] : null;
            };

            const cod = sanitize(getVal(['cód. produto', 'cod. produto', 'produto', 'codigo']));
            const desc = String(getVal(['descrição', 'descric', 'nome']) || '').trim();
            const un = String(getVal(['unidade', 'un.']) || '').trim();
            const marca = String(getVal(['marca', 'fabricante']) || '').trim();
            const saldo = parseMoney(getVal(['disponível', 'disponivel', 'saldo', 'estoque']));
            
            const pack = String(getVal(['pack']) || '').trim();
            const sortimento = String(getVal(['sortimento', 'sort.']) || '').trim();
            const pv = parseMoney(getVal(['pv', 'preço venda', 'preco venda']));
            const pdv = parseMoney(getVal(['pdv', 'sugerido', 'preço sugerido']));
            const ean = sanitize(getVal(['ean', 'barra']));
            const previsao = String(getVal(['previsão', 'previsao', 'chegada']) || '').trim();

            if (cod) {
                const descClean = toTitleCaseClean(desc);
                const marcaClean = toTitleCaseClean(marca);

                const existing = await query('SELECT cod_produto FROM estoque WHERE cod_produto = ?', [cod]);
                if (existing.length > 0) {
                    await query(`UPDATE estoque SET 
                                 descricao=COALESCE(?, descricao), 
                                 unidade=COALESCE(?, unidade), 
                                 marca=COALESCE(?, marca), 
                                 saldo=?, pack=?, sortimento=?, pv=?, pdv=?, 
                                 ean=COALESCE(?, ean), previsao=? 
                                 WHERE cod_produto = ?`, 
                                 [descClean || null, un || null, marcaClean || null, saldo, pack, sortimento, pv, pdv, ean || null, previsao, cod]);
                } else {
                    await query(`INSERT INTO estoque (cod_produto, descricao, unidade, marca, saldo, pack, sortimento, pv, pdv, ean, previsao) 
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`, 
                                 [cod, descClean, un, marcaClean, saldo, pack, sortimento, pv, pdv, ean, previsao]);
                }
                
                if (marcaClean) {
                    const insertMarcaQuery = usePostgres 
                        ? 'INSERT INTO marcas_mestre (nome) VALUES (?) ON CONFLICT (nome) DO NOTHING'
                        : 'INSERT OR IGNORE INTO marcas_mestre (nome) VALUES (?)';
                    await query(insertMarcaQuery, [marcaClean]);
                }
                count++;
            }
        }
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({ 
            message: `Processamento concluído.`, 
            detail: `Estoque atualizado para ${count} produtos. ${newBrandsCount} novas marcas identificadas e vinculadas.`
        });
    } catch (err) { 
        console.error(`[Import Estoque] Erro:`, err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: err.message }); 
    }
});

// --- SINCRONIZAÇÃO VIA PORTAL SUNNY ---
// Credenciais armazenadas em variáveis de ambiente (segurança)
const PORTAL_BASE_URL = process.env.SUNNY_PORTAL_URL || 'https://site-sunny.com.br/rest2';
const PORTAL_USER = process.env.SUNNY_PORTAL_USER || 'trocha';
const PORTAL_PASS = process.env.SUNNY_PORTAL_PASS || '123';

async function fetchPortalSunny(endpoint) {
    const basicAuth = 'BASIC ' + Buffer.from(`${PORTAL_USER}:${PORTAL_PASS}`).toString('base64');
    const url = `${PORTAL_BASE_URL}${endpoint}`;
    
    // Salva o estado original e ignora a validação estrita de SSL temporariamente
    const originalReject = process.env.NODE_TLS_REJECT_UNAUTHORIZED;
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

    try {
        console.log(`[Portal Sync] Fazendo fetch em: ${url}`);
        
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'Authorization': basicAuth,
                'Content-Type': 'application/json',
                'Accept': 'application/json, text/plain, */*',
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
                'Accept-Language': 'pt-BR,pt;q=0.9,en;q=0.8',
                'Cache-Control': 'no-cache',
                'Pragma': 'no-cache',
                'Connection': 'keep-alive'
            }
        });

        // Restaura a configuração TLS original
        if (originalReject === undefined) {
            delete process.env.NODE_TLS_REJECT_UNAUTHORIZED;
        } else {
            process.env.NODE_TLS_REJECT_UNAUTHORIZED = originalReject;
        }

        if (response.status === 401) {
            throw new Error('Credenciais inválidas (401)');
        }
        if (response.status === 404) {
            throw new Error(`Endpoint não encontrado (404): ${endpoint}`);
        }
        if (!response.ok) {
            throw new Error(`Erro HTTP ${response.status}: ${response.statusText}`);
        }

        const text = await response.text();
        try {
            return { status: response.status, data: JSON.parse(text) };
        } catch (e) {
            throw new Error(`Resposta do portal não é um JSON válido: ${text.substring(0, 200)}`);
        }
    } catch (err) {
        // Garante a restauração do TLS mesmo em caso de erro na requisição
        if (originalReject === undefined) {
            delete process.env.NODE_TLS_REJECT_UNAUTHORIZED;
        } else {
            process.env.NODE_TLS_REJECT_UNAUTHORIZED = originalReject;
        }
        throw err;
    }
}

app.get('/api/sync-portal-sunny', async (req, res) => {
    try {
        console.log('[Portal Sync] Iniciando sincronização com portal Sunny...');
        
        // Diagnóstico de colunas e triggers no PostgreSQL
        if (usePostgres) {
            try {
                const cols = await query(`
                    SELECT column_name, data_type 
                    FROM information_schema.columns 
                    WHERE table_name = 'estoque'
                `);
                console.log('[Portal Sync Debug Schema] Colunas da tabela estoque:', JSON.stringify(cols, null, 2));
                
                const triggers = await query(`
                    SELECT trigger_name, event_manipulation, action_statement 
                    FROM information_schema.triggers 
                    WHERE event_object_table = 'estoque'
                `);
                console.log('[Portal Sync Debug Schema] Triggers da tabela estoque:', JSON.stringify(triggers, null, 2));
            } catch (schemaErr) {
                console.log('[Portal Sync Debug Schema] Erro ao diagnosticar estrutura:', schemaErr.message);
            }
        }
        
        // Tentar múltiplos endpoints possíveis para tabela de preço/estoque
        const endpointsParaTentar = [
            '/listaprodutos?take=10000',
            '/listaprodutos?take=9999',
            '/listaprodutos',
            '/produtos',
            '/estoque',
            '/tabelapreco',
            '/catalogo',
            '/products',
            '/itens',
            '/sunny/produtos',
        ];
        
        let produtos = null;
        let endpointUsado = null;
        let lastError = null;
        
        for (const ep of endpointsParaTentar) {
            try {
                console.log(`[Portal Sync] Tentando endpoint: ${ep}`);
                const result = await fetchPortalSunny(ep);
                const data = result.data;
                // Aceitar arrays ou objetos com array dentro
                const arr = Array.isArray(data) ? data : (data?.data || data?.items || data?.produtos || data?.produtos_list || null);
                if (arr && arr.length > 0) {
                    produtos = arr;
                    endpointUsado = ep;
                    console.log(`[Portal Sync] ✅ Endpoint ${ep} retornou ${arr.length} produtos`);
                    
                    // Se o endpoint retornou uma lista grande (mais de 20 itens), consideramos bem sucedido
                    // Se retornou poucos itens (ex: 10) e ainda temos endpoints de paginação para tentar,
                    // continuamos tentando para ver se achamos um que retorne a lista inteira.
                    if (arr.length > 20) {
                        break;
                    } else {
                        console.log(`[Portal Sync] Endpoint ${ep} retornou apenas ${arr.length} itens. Continuaremos buscando por um endpoint completo...`);
                    }
                }
            } catch (e) {
                lastError = e;
                console.log(`[Portal Sync] Endpoint ${ep} falhou: ${e.message}`);
            }
        }
        
        if (!produtos) {
            throw new Error(`Nenhum endpoint retornou dados válidos. Último erro: ${lastError?.message}`);
        }
        
        // Log de debug do primeiro produto para verificar a estrutura de chaves
        if (produtos.length > 0) {
            console.log('[Portal Sync] Exemplo da estrutura do primeiro produto retornado pelo portal:');
            console.log(JSON.stringify(produtos[0], null, 2));
            console.log(`[Portal Sync] Total de produtos carregados na lista: ${produtos.length}`);
        }
        
        // Mapeamento flexível de colunas (igual ao import manual)
        const getVal = (obj, keys) => {
            const foundKey = Object.keys(obj).find(k => 
                keys.some(t => k.toLowerCase().trim().replace(/[_\s]/g, '').includes(t.toLowerCase().replace(/[_\s]/g, '')))
            );
            return foundKey ? obj[foundKey] : null;
        };
        
        let atualizados = 0;
        let inseridos = 0;
        let erros = 0;
        
        for (const row of produtos) {
            try {
                // Mapeia chaves Protheus b1_cod, etc. e chaves comuns
                const cod = sanitize(
                    getVal(row, ['b1_cod', 'cod_produto', 'codigo', 'code', 'produto_id', 'sku', 'codproduto', 'id'])
                );
                if (!cod) continue;
                
                const desc = String(getVal(row, ['b1_desc', 'descricao', 'description', 'nome', 'name', 'produto']) || '').trim();
                const marca = String(getVal(row, ['bm_desc', 'marca', 'brand', 'fabricante']) || '').trim();
                const saldo = parseMoney(getVal(row, ['disponivel', 'saldo', 'disponível', 'estoque', 'qtd', 'quantidade', 'stock']));
                const pv = parseMoney(getVal(row, ['da1_prcven', 'pv', 'preco_venda', 'price', 'valor', 'preco']));
                const pdv = parseMoney(getVal(row, ['da1_xprcsu', 'pdv', 'preco_sugerido', 'retail_price', 'preco_pdv']));
                const previsao = String(getVal(row, ['b1_xprevis', 'b1_xdatche', 'previsao', 'previsão', 'arrival', 'chegada', 'data_previsao']) || '').trim();
                const ean = sanitize(getVal(row, ['b1_codbar', 'ean', 'barcode', 'barra', 'codigo_barras']));
                const unidade = String(getVal(row, ['b1_um', 'unidade', 'unit', 'un']) || '').trim();
                const pack = String(getVal(row, ['pack', 'embalagem']) || '').trim();
                const sortimento = String(getVal(row, ['b1_xsortim', 'sortimento', 'sort']) || '').trim();
                const image_url = String(getVal(row, ['image_url', 'imagem', 'foto', 'image']) || '').trim();
                
                const descClean = toTitleCaseClean(desc);
                const marcaClean = toTitleCaseClean(marca);
                
                const existing = await query('SELECT cod_produto FROM estoque WHERE cod_produto = ?', [cod]);
                
                if (existing.length > 0) {
                    // UPDATE: atualiza saldo e previsão (campos principais) + outros se disponíveis
                    await query(`UPDATE estoque SET 
                        saldo = ?,
                        previsao = CASE WHEN ? != '' THEN ? ELSE previsao END,
                        pv = CASE WHEN CAST(? AS DOUBLE PRECISION) > 0 THEN CAST(? AS DOUBLE PRECISION) ELSE pv END,
                        pdv = CASE WHEN CAST(? AS DOUBLE PRECISION) > 0 THEN CAST(? AS DOUBLE PRECISION) ELSE pdv END,
                        descricao = CASE WHEN ? != '' THEN ? ELSE descricao END,
                        marca = CASE WHEN ? != '' THEN ? ELSE marca END,
                        ean = CASE WHEN ? != '' THEN ? ELSE ean END,
                        image_url = CASE WHEN ? != '' THEN ? ELSE image_url END
                        WHERE cod_produto = ?`,
                        [saldo,
                         previsao, previsao,
                         pv, pv,
                         pdv, pdv,
                         descClean, descClean,
                         marcaClean, marcaClean,
                         ean, ean,
                         image_url, image_url,
                         cod]);
                    atualizados++;
                } else {
                    // INSERT: novo produto descoberto
                    await query(`INSERT INTO estoque (cod_produto, descricao, unidade, marca, saldo, pack, sortimento, pv, pdv, ean, previsao, image_url)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                        [cod, descClean, unidade, marcaClean, saldo, pack, sortimento, pv, pdv, ean, previsao, image_url]);
                    inseridos++;
                    
                    // Registrar marca nova
                    if (marcaClean) {
                        const insertMarcaQuery = usePostgres 
                            ? 'INSERT INTO marcas_mestre (nome) VALUES (?) ON CONFLICT (nome) DO NOTHING'
                            : 'INSERT OR IGNORE INTO marcas_mestre (nome) VALUES (?)';
                        await query(insertMarcaQuery, [marcaClean]);
                    }
                }
            } catch (rowErr) {
                erros++;
                console.error(`[Portal Sync] Erro na linha:`, rowErr.message);
            }
        }
        
        console.log(`[Portal Sync] ✅ Concluído: ${atualizados} atualizados, ${inseridos} inseridos, ${erros} erros`);
        res.json({
            success: true,
            endpoint: endpointUsado,
            total: produtos.length,
            atualizados,
            inseridos,
            erros,
            message: `Sincronização concluída via ${endpointUsado}`,
            detail: `${atualizados} produtos atualizados | ${inseridos} novos produtos inseridos${erros > 0 ? ` | ${erros} erros` : ''}`
        });
        
    } catch (err) {
        console.error('[Portal Sync] Erro fatal:', err.message);
        res.status(500).json({ 
            success: false,
            error: err.message,
            detail: 'Verifique as credenciais e disponibilidade do portal Sunny'
        });
    }
});

// --- IMPORTAÇÃO: GALERIA DE FOTOS (URL por SKU) ---
app.post('/api/import/galeria', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Arquivo não enviado");
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);

        console.log(`[Import Galeria] Processando ${data.length} linhas.`);
        if (data.length > 0) console.log(`[Import Galeria] Colunas:`, Object.keys(data[0]));

        const normalize = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").trim();
        const getV = (row, targets) => {
            const normTargets = targets.map(normalize);
            const col = Object.keys(row).find(k => normTargets.some(t => normalize(k) === t));
            return col ? String(row[col] || '').trim() : '';
        };

        let updated = 0;
        let notFound = 0;

        for (const row of data) {
            // Aceita variações de nome de coluna
            const rawCod = getV(row, ['codigo', 'codigo do produto', 'cod', 'cod produto', 'sku', 'produto', 'cod_produto']);
            const url    = getV(row, ['url', 'url da imagem', 'imagem', 'foto', 'image_url', 'link']);

            if (!rawCod || !url) continue;

            // Higienização: remove zeros à esquerda
            const cod = String(rawCod).trim().replace(/^0+/, '');

            const existing = await query('SELECT cod_produto FROM estoque WHERE cod_produto = ?', [cod]);
            if (existing.length > 0) {
                await query('UPDATE estoque SET image_url = ? WHERE cod_produto = ?', [url, cod]);
                updated++;
            } else {
                notFound++;
                console.log(`[Import Galeria] SKU não encontrado: ${cod} (original: ${rawCod})`);
            }
        }

        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({
            message: `Galeria atualizada com sucesso.`,
            detail: `${updated} imagens vinculadas. ${notFound} SKUs não encontrados na base.`
        });
    } catch (err) {
        console.error(`[Import Galeria] Erro:`, err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: err.message });
    }
});

app.post('/api/import/vendas', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Arquivo não enviado");
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);
        
        let added = 0;
        let ignored = 0;

        const normalize = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "").trim();

        // Salva a primeira linha para debug real
        if (data.length > 0) {
            fs.writeFileSync(path.join(__dirname, 'debug_importacao.json'), JSON.stringify({
                original_keys: Object.keys(data[0]),
                normalized_keys: Object.keys(data[0]).map(normalize),
                first_row: data[0]
            }, null, 2));
        }

        if (usePostgres) {
            const client = await pgPool.connect();
            try {
                await client.query("BEGIN");
                const insertQuery = `INSERT INTO vendas (num_docto, emissao, cnpj, nome_cliente, produto_id, ean, descricao_produto, quantidade, valor_total, valor_unitario, nome_vendedor, uf, marca, status, almox, gerente_id, vendedor_id, cliente_id, natureza_operacao) VALUES ($1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19)`;
                
                for (const row of data) {
                    const getV = (targets) => {
                        const normTargets = targets.map(normalize);
                        const col = Object.keys(row).find(k => normTargets.some(t => normalize(k) === t)) ||
                                    Object.keys(row).find(k => normTargets.some(t => normalize(k).includes(t)));
                        return col ? row[col] : null;
                    };

                    const numDocto = sanitize(getV(['numdocto', 'nffiscal', 'documento']));
                    if (!numDocto) continue;

                    const rawAlmox = getV(['almox', 'almoxarifado']);
                    const almoxVal = rawAlmox !== null && rawAlmox !== undefined ? String(rawAlmox).trim() : '';
                    const rawStatus = getV(['status', 'situa']);
                    const statusVal = rawStatus !== null && rawStatus !== undefined ? String(rawStatus).trim() : '';
                    
                    // Filtrar: almox deve ser '20' ou '0020' se presente. Se ausente, assume '20'.
                    if (almoxVal !== '' && almoxVal !== '20' && almoxVal !== '0020') {
                        ignored++;
                        continue;
                    }
                    // Filtrar: status deve ser '5' ou '6' (ou '05', '06') se presente. Se ausente, assume '5'.
                    if (statusVal !== '' && statusVal !== '5' && statusVal !== '6' && statusVal !== '05' && statusVal !== '06') {
                        ignored++;
                        continue;
                    }
                    
                    const finalAlmox = almoxVal !== '' ? (almoxVal === '0020' ? '20' : almoxVal) : '20';
                    const finalStatus = statusVal !== '' ? (statusVal === '05' ? '5' : (statusVal === '06' ? '6' : statusVal)) : '5';

                    const dataVenda = excelDate(getV(['emissao', 'data']));
                    const cnpj = sanitize(getV(['cnpj', 'cgccpf']));
                    const clienteNome = String(getV(['nome', 'nomecliente', 'razaosocial']) || '').trim();
                    const clienteId = sanitize(getV(['cliente', 'codcliente']));
                    const produtoId = sanitize(getV(['produto', 'codproduto']));
                    const ean = sanitize(getV(['ean', 'codbarras']));
                    const descricao = String(getV(['descricao', 'descproduto']) || '').trim();
                    const valorRaw = getV(['vlrtotal', 'valorTotal', 'valorliq']);
                    const valorTotal = parseMoney(valorRaw);
                    const qtd = parseFloat(getV(['quantidade', 'qtd']) || 0);
                    const vUnitRaw = getV(['vlrunitario', 'unitario', 'precounitario']);
                    const valorUnitario = vUnitRaw ? parseMoney(vUnitRaw) : (qtd > 0 ? valorTotal / qtd : 0);
                    const nomeVendedor = String(getV(['nomevendedor', 'vendedornome']) || '').trim();
                    const vendedorId = sanitize(getV(['vendedor', 'codvendedor']));
                    const uf = String(getV(['uf', 'estado']) || '').trim().toUpperCase();
                    const marca = String(getV(['marca', 'fabricante', 'circana']) || '').trim();
                    const gerenteId = sanitize(getV(['gerente', 'codgerente']));
                    const natOperacaoVal = String(getV(['natureza', 'operacao', 'naturezadeoperacao']) || 'Venda').trim();

                    await client.query(insertQuery, [
                        numDocto, dataVenda, cnpj, clienteNome, produtoId, ean, descricao, 
                        qtd, valorTotal, valorUnitario, nomeVendedor, uf, marca, finalStatus, finalAlmox, gerenteId, vendedorId, clienteId, natOperacaoVal
                    ]);
                    added++;
                }
                await client.query("COMMIT");
            } catch (err) {
                await client.query("ROLLBACK");
                throw err;
            } finally {
                client.release();
            }
        } else {
            await new Promise((resolve, reject) => {
                db.serialize(() => {
                    db.run("BEGIN TRANSACTION");
                    
                    const stmt = db.prepare(`INSERT OR REPLACE INTO vendas (num_docto, emissao, cnpj, nome_cliente, produto_id, ean, descricao_produto, quantidade, valor_total, valor_unitario, nome_vendedor, uf, marca, status, almox, gerente_id, vendedor_id, cliente_id, natureza_operacao) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`);

                    for (const row of data) {
                        const getV = (targets) => {
                            const normTargets = targets.map(normalize);
                            const col = Object.keys(row).find(k => normTargets.some(t => normalize(k) === t)) ||
                                        Object.keys(row).find(k => normTargets.some(t => normalize(k).includes(t)));
                            return col ? row[col] : null;
                        };

                        const numDocto = sanitize(getV(['numdocto', 'nffiscal', 'documento']));
                        if (!numDocto) continue;

                        const rawAlmox = getV(['almox', 'almoxarifado']);
                        const almoxVal = rawAlmox !== null && rawAlmox !== undefined ? String(rawAlmox).trim() : '';
                        const rawStatus = getV(['status', 'situa']);
                        const statusVal = rawStatus !== null && rawStatus !== undefined ? String(rawStatus).trim() : '';
                        
                        // Filtrar: almox deve ser '20' ou '0020' se presente. Se ausente, assume '20'.
                        if (almoxVal !== '' && almoxVal !== '20' && almoxVal !== '0020') {
                            ignored++;
                            continue;
                        }
                        // Filtrar: status deve ser '5' ou '6' (ou '05', '06') se presente. Se ausente, assume '5'.
                        if (statusVal !== '' && statusVal !== '5' && statusVal !== '6' && statusVal !== '05' && statusVal !== '06') {
                            ignored++;
                            continue;
                        }
                        
                        const finalAlmox = almoxVal !== '' ? (almoxVal === '0020' ? '20' : almoxVal) : '20';
                        const finalStatus = statusVal !== '' ? (statusVal === '05' ? '5' : (statusVal === '06' ? '6' : statusVal)) : '5';

                        const dataVenda = excelDate(getV(['emissao', 'data']));
                        const cnpj = sanitize(getV(['cnpj', 'cgccpf']));
                        const clienteNome = String(getV(['nome', 'nomecliente', 'razaosocial']) || '').trim();
                        const clienteId = sanitize(getV(['cliente', 'codcliente']));
                        const produtoId = sanitize(getV(['produto', 'codproduto']));
                        const ean = sanitize(getV(['ean', 'codbarras']));
                        const descricao = String(getV(['descricao', 'descproduto']) || '').trim();
                        const valorRaw = getV(['vlrtotal', 'valorTotal', 'valorliq']);
                        const valorTotal = parseMoney(valorRaw);
                        const qtd = parseFloat(getV(['quantidade', 'qtd']) || 0);
                        const vUnitRaw = getV(['vlrunitario', 'unitario', 'precounitario']);
                        const valorUnitario = vUnitRaw ? parseMoney(vUnitRaw) : (qtd > 0 ? valorTotal / qtd : 0);
                        const nomeVendedor = String(getV(['nomevendedor', 'vendedornome']) || '').trim();
                        const vendedorId = sanitize(getV(['vendedor', 'codvendedor']));
                        const uf = String(getV(['uf', 'estado']) || '').trim().toUpperCase();
                        const marca = String(getV(['marca', 'fabricante', 'circana']) || '').trim();
                        const gerenteId = sanitize(getV(['gerente', 'codgerente']));
                        const natOperacaoVal = String(getV(['natureza', 'operacao', 'naturezadeoperacao']) || 'Venda').trim();

                        stmt.run([
                            numDocto, dataVenda, cnpj, clienteNome, produtoId, ean, descricao, 
                            qtd, valorTotal, valorUnitario, nomeVendedor, uf, marca, finalStatus, finalAlmox, gerenteId, vendedorId, clienteId, natOperacaoVal
                        ]);
                        
                        added++;
                    }

                    stmt.finalize();
                    db.run("COMMIT", (err) => {
                        if (err) reject(err);
                        else resolve();
                    });
                });
            });
        }


        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({ 
            message: `Importação concluída com sucesso.`, 
            detail: `${added} novos registros adicionados, ${ignored} duplicados ignorados. Total: ${data.length}`
        });
    } catch (err) { 
        console.error(`[Import Vendas] Erro:`, err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: err.message }); 
    }
});

// Helper para dividir intervalo de datas em blocos de X dias
function getChunks(startStr, endStr, chunkSizeDays = 15) {
    const start = new Date(startStr + 'T00:00:00');
    const end = new Date(endStr + 'T00:00:00');
    const chunks = [];
    
    let currentStart = new Date(start);
    while (currentStart <= end) {
        let currentEnd = new Date(currentStart);
        currentEnd.setDate(currentEnd.getDate() + chunkSizeDays - 1);
        if (currentEnd > end) {
            currentEnd = new Date(end);
        }
        
        const yStart = currentStart.getFullYear();
        const mStart = String(currentStart.getMonth() + 1).padStart(2, '0');
        const dStart = String(currentStart.getDate()).padStart(2, '0');
        
        const yEnd = currentEnd.getFullYear();
        const mEnd = String(currentEnd.getMonth() + 1).padStart(2, '0');
        const dEnd = String(currentEnd.getDate()).padStart(2, '0');
        
        chunks.push({
            startStr: `${yStart}-${mStart}-${dStart}`,
            endStr: `${yEnd}-${mEnd}-${dEnd}`,
            apiStart: `${yStart}${mStart}${dStart}`,
            apiEnd: `${yEnd}${mEnd}${dEnd}`
        });
        
        currentStart = new Date(currentEnd);
        currentStart.setDate(currentStart.getDate() + 1);
    }
    return chunks;
}

// Endpoint de Sincronização direta com a API do Protheus
app.post('/api/sync-nfs', async (req, res) => {
    const { mode, startDate, endDate } = req.body;
    console.log(`[Sync NFs] Iniciado. Modo: ${mode}, Range recebido: ${startDate} a ${endDate}`);
    
    const originalReject = process.env.NODE_TLS_REJECT_UNAUTHORIZED;
    process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';
    
    try {
        let finalStart = '';
        let finalEnd = '';
        
        const todayStr = new Date().toISOString().split('T')[0];
        
        if (mode === 'revisar') {
            const d = new Date();
            d.setDate(d.getDate() - 30);
            finalStart = d.toISOString().split('T')[0];
            finalEnd = todayStr;
        } else if (mode === 'atualizar') {
            const maxDateRow = await query("SELECT MAX(emissao) as max_date FROM vendas");
            const maxDate = maxDateRow[0]?.max_date;
            if (maxDate && /^\d{4}-\d{2}-\d{2}$/.test(maxDate)) {
                finalStart = maxDate;
            } else {
                finalStart = '2025-01-02';
            }
            finalEnd = todayStr;
        } else if (mode === 'carga') {
            if (!startDate || !endDate) {
                throw new Error("Parâmetros startDate e endDate são obrigatórios no modo de carga.");
            }
            finalStart = startDate;
            finalEnd = endDate;
        } else {
            throw new Error(`Modo de sincronização inválido: ${mode}`);
        }
        
        // Limita a data de início para Protheus API (mínimo de 2025-01-02)
        const minAllowedDate = '2025-01-02';
        if (finalStart < minAllowedDate) {
            console.log(`[Sync NFs] Data de início ajustada de ${finalStart} para ${minAllowedDate} devido a limitações da API.`);
            finalStart = minAllowedDate;
        }
        
        if (finalStart > finalEnd) {
            return res.json({
                success: true,
                inserted: 0,
                updated: 0,
                message: "A data de início da sincronização é posterior à data de fim. Nada a sincronizar.",
                detail: `Período verificado: ${finalStart} até ${finalEnd}`
            });
        }
        
        const chunks = getChunks(finalStart, finalEnd, 15);
        console.log(`[Sync NFs] Período dividido em ${chunks.length} blocos.`);
        
        // Cache de Vendedor -> Gerente
        const vToGMap = {};
        try {
            const mapRows = await query("SELECT DISTINCT vendedor_id, gerente_id FROM vendas WHERE vendedor_id IS NOT NULL AND vendedor_id != ''");
            mapRows.forEach(r => {
                if (r.gerente_id) vToGMap[r.vendedor_id] = r.gerente_id;
            });
        } catch (dbErr) {
            console.log("[Sync NFs] Erro ao carregar mapa de gerentes:", dbErr.message);
        }
        
        let totalInserted = 0;
        let totalUpdated = 0;
        
        const user = 'trocha';
        const pass = '123';
        const basicAuth = 'Basic ' + Buffer.from(`${user}:${pass}`).toString('base64');
        
        for (const chunk of chunks) {
            console.log(`[Sync NFs] Processando bloco: ${chunk.startStr} a ${chunk.endStr}`);
            
            let skip = 0;
            const take = 100;
            let hasMore = true;
            
            while (hasMore) {
                const filterObj = [
                    ["f2_emissao", ">=", `'${chunk.apiStart}`],
                    "and",
                    ["f2_emissao", "<=", `'${chunk.apiEnd}`],
                    "and",
                    ["d2_local", "=", "20"],
                    "and",
                    [
                        ["f2_xstatus", "=", "5"],
                        "or",
                        ["f2_xstatus", "=", "6"]
                    ]
                ];
                const filterStr = JSON.stringify(filterObj);
                const param = `usuario=000114&group=null&filter=${encodeURIComponent(filterStr)}`;
                const url = `https://site-sunny.com.br:8087/rest/nfs/lista?${param}&take=${take}&skip=${skip}`;
                
                let response = null;
                let retries = 3;
                let delayMs = 1000;
                
                while (retries > 0) {
                    try {
                        response = await fetch(url, {
                            method: 'GET',
                            headers: {
                                'Authorization': basicAuth,
                                'Content-Type': 'application/json',
                                'Accept': 'application/json'
                            }
                        });
                        
                        if (response.ok) {
                            break;
                        }
                        
                        if (response.status >= 500 || response.status === 429) {
                            console.log(`[Sync NFs] Resposta HTTP ${response.status} ao chamar API. Tentando novamente em ${delayMs}ms... (${retries - 1} tentativas restantes)`);
                            retries--;
                            if (retries === 0) {
                                const text = await response.text();
                                throw new Error(`Erro na API Protheus (HTTP ${response.status}): ${text.substring(0, 200)}`);
                            }
                            await new Promise(resolve => setTimeout(resolve, delayMs));
                            delayMs *= 2;
                        } else {
                            const text = await response.text();
                            throw new Error(`Erro na API Protheus (HTTP ${response.status}): ${text.substring(0, 200)}`);
                        }
                    } catch (fetchErr) {
                        retries--;
                        if (retries === 0) {
                            throw fetchErr;
                        }
                        console.log(`[Sync NFs] Falha na requisição: ${fetchErr.message}. Tentando novamente em ${delayMs}ms... (${retries} tentativas restantes)`);
                        await new Promise(resolve => setTimeout(resolve, delayMs));
                        delayMs *= 2;
                    }
                }
                
                const parsed = await response.json();
                const items = parsed.data || parsed.items || [];
                
                if (items.length === 0) {
                    hasMore = false;
                    break;
                }
                
                if (usePostgres) {
                    const client = await pgPool.connect();
                    try {
                        await client.query("BEGIN");
                        for (const item of items) {
                            const numDocto = sanitize(item.f2_doc);
                            const produtoId = sanitize(item.d2_cod);
                            if (!numDocto || !produtoId) continue;
                            
                            let formattedEmissao = '';
                            if (item.f2_emissao && item.f2_emissao.length === 8) {
                                formattedEmissao = `${item.f2_emissao.substring(0, 4)}-${item.f2_emissao.substring(4, 6)}-${item.f2_emissao.substring(6, 8)}`;
                            } else {
                                formattedEmissao = todayStr;
                            }
                            
                            const cnpj = sanitize(item.a1_cgc);
                            const clienteNome = String(item.a1_nome || '').trim();
                            const clienteId = sanitize(item.a1_cod);
                            const ean = sanitize(item.b1_codbar);
                            const descricao = String(item.b1_desc || '').trim();
                            const qtd = parseFloat(item.d2_quant) || 0;
                            const valorUnitario = parseFloat(item.d2_prcven) || 0;
                            const valorTotal = parseFloat(item.d2_total) || 0;
                            const nomeVendedor = String(item.a3_nreduz || '').trim();
                            const vendedorId = sanitize(item.a3_cod);
                            const uf = String(item.a1_est || '').trim().toUpperCase();
                            const marca = String(item.bm_desc || '').trim();
                            const status = sanitize(item.f2_xstatus) || '5';
                            const almox = '20';
                            const gerenteId = vToGMap[vendedorId] || '';
                            const pedidoCliente = sanitize(item.d2_pedido);
                            const loja = sanitize(item.a1_loja);
                            const naturezaOperacao = 'Venda';
                            
                            const existRes = await client.query(
                                "SELECT id FROM vendas WHERE num_docto = $1 AND produto_id = $2",
                                [numDocto, produtoId]
                            );
                            
                            if (existRes.rows.length > 0) {
                                await client.query(
                                    `UPDATE vendas SET 
                                        emissao = $1, cnpj = $2, nome_cliente = $3, ean = $4, 
                                        descricao_produto = $5, quantidade = $6, valor_unitario = $7, 
                                        valor_total = $8, nome_vendedor = $9, uf = $10, marca = $11, 
                                        status = $12, almox = $13, gerente_id = $14, vendedor_id = $15, 
                                        cliente_id = $16, pedido_cliente = $17, loja = $18, natureza_operacao = $19
                                     WHERE num_docto = $20 AND produto_id = $21`,
                                    [
                                        formattedEmissao, cnpj, clienteNome, ean, descricao, qtd, 
                                        valorUnitario, valorTotal, nomeVendedor, uf, marca, status, 
                                        almox, gerenteId, vendedorId, clienteId, pedidoCliente, loja, 
                                        naturezaOperacao, numDocto, produtoId
                                    ]
                                );
                                totalUpdated++;
                            } else {
                                await client.query(
                                    `INSERT INTO vendas (
                                        num_docto, emissao, cnpj, nome_cliente, produto_id, ean, 
                                        descricao_produto, quantidade, valor_unitario, valor_total, 
                                        nome_vendedor, uf, marca, status, almox, gerente_id, vendedor_id, 
                                        cliente_id, pedido_cliente, loja, natureza_operacao
                                     ) VALUES (
                                        $1, $2, $3, $4, $5, $6, $7, $8, $9, $10, $11, $12, $13, $14, $15, $16, $17, $18, $19, $20, $21
                                     )`,
                                    [
                                        numDocto, formattedEmissao, cnpj, clienteNome, produtoId, ean, 
                                        descricao, qtd, valorUnitario, valorTotal, nomeVendedor, uf, 
                                        marca, status, almox, gerenteId, vendedorId, clienteId, 
                                        pedidoCliente, loja, naturezaOperacao
                                    ]
                                );
                                totalInserted++;
                            }
                        }
                        await client.query("COMMIT");
                    } catch (err) {
                        await client.query("ROLLBACK");
                        throw err;
                    } finally {
                        client.release();
                    }
                } else {
                    // SQLite
                    await new Promise((resolve, reject) => {
                        db.serialize(async () => {
                            db.run("BEGIN TRANSACTION");
                            try {
                                for (const item of items) {
                                    const numDocto = sanitize(item.f2_doc);
                                    const produtoId = sanitize(item.d2_cod);
                                    if (!numDocto || !produtoId) continue;
                                    
                                    let formattedEmissao = '';
                                    if (item.f2_emissao && item.f2_emissao.length === 8) {
                                        formattedEmissao = `${item.f2_emissao.substring(0, 4)}-${item.f2_emissao.substring(4, 6)}-${item.f2_emissao.substring(6, 8)}`;
                                    } else {
                                        formattedEmissao = todayStr;
                                    }
                                    
                                    const cnpj = sanitize(item.a1_cgc);
                                    const clienteNome = String(item.a1_nome || '').trim();
                                    const clienteId = sanitize(item.a1_cod);
                                    const ean = sanitize(item.b1_codbar);
                                    const descricao = String(item.b1_desc || '').trim();
                                    const qtd = parseFloat(item.d2_quant) || 0;
                                    const valorUnitario = parseFloat(item.d2_prcven) || 0;
                                    const valorTotal = parseFloat(item.d2_total) || 0;
                                    const nomeVendedor = String(item.a3_nreduz || '').trim();
                                    const vendedorId = sanitize(item.a3_cod);
                                    const uf = String(item.a1_est || '').trim().toUpperCase();
                                    const marca = String(item.bm_desc || '').trim();
                                    const status = sanitize(item.f2_xstatus) || '5';
                                    const almox = '20';
                                    const gerenteId = vToGMap[vendedorId] || '';
                                    const pedidoCliente = sanitize(item.d2_pedido);
                                    const loja = sanitize(item.a1_loja);
                                    const naturezaOperacao = 'Venda';
                                    
                                    const existing = await new Promise((resSelect) => {
                                        db.get("SELECT id FROM vendas WHERE num_docto = ? AND produto_id = ?", [numDocto, produtoId], (e, r) => {
                                            resSelect(r);
                                        });
                                    });
                                    
                                    if (existing) {
                                        db.run(
                                            `UPDATE vendas SET 
                                                emissao = ?, cnpj = ?, nome_cliente = ?, ean = ?, 
                                                descricao_produto = ?, quantidade = ?, valor_unitario = ?, 
                                                valor_total = ?, nome_vendedor = ?, uf = ?, marca = ?, 
                                                status = ?, almox = ?, gerente_id = ?, vendedor_id = ?, 
                                                cliente_id = ?, pedido_cliente = ?, loja = ?, natureza_operacao = ?
                                             WHERE num_docto = ? AND produto_id = ?`,
                                            [
                                                formattedEmissao, cnpj, clienteNome, ean, descricao, qtd, 
                                                valorUnitario, valorTotal, nomeVendedor, uf, marca, status, 
                                                almox, gerenteId, vendedorId, clienteId, pedidoCliente, loja, 
                                                naturezaOperacao, numDocto, produtoId
                                            ]
                                        );
                                        totalUpdated++;
                                    } else {
                                        db.run(
                                            `INSERT INTO vendas (
                                                num_docto, emissao, cnpj, nome_cliente, produto_id, ean, 
                                                descricao_produto, quantidade, valor_unitario, valor_total, 
                                                nome_vendedor, uf, marca, status, almox, gerente_id, vendedor_id, 
                                                cliente_id, pedido_cliente, loja, natureza_operacao
                                             ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                                            [
                                                numDocto, formattedEmissao, cnpj, clienteNome, produtoId, ean, 
                                                descricao, qtd, valorUnitario, valorTotal, nomeVendedor, uf, 
                                                marca, status, almox, gerenteId, vendedorId, clienteId, 
                                                pedidoCliente, loja, naturezaOperacao
                                            ]
                                        );
                                        totalInserted++;
                                    }
                                }
                                db.run("COMMIT", (commitErr) => {
                                    if (commitErr) reject(commitErr);
                                    else resolve();
                                });
                            } catch (sqliteErr) {
                                db.run("ROLLBACK");
                                reject(sqliteErr);
                            }
                        });
                    });
                }
                
                if (items.length < take) {
                    hasMore = false;
                } else {
                    skip += take;
                    await new Promise(resolve => setTimeout(resolve, 200));
                }
            }
        }
        
        res.json({
            success: true,
            inserted: totalInserted,
            updated: totalUpdated,
            message: `Sincronização concluída com sucesso.`,
            detail: `Período: ${finalStart} a ${finalEnd} | Inseridos: ${totalInserted} | Atualizados: ${totalUpdated}`
        });
        
    } catch (err) {
        console.error(`[Sync NFs] Erro fatal:`, err);
        res.status(500).json({
            success: false,
            error: err.message,
            detail: 'Erro interno ao sincronizar notas fiscais com Protheus API.'
        });
    } finally {
        if (originalReject === undefined) {
            delete process.env.NODE_TLS_REJECT_UNAUTHORIZED;
        } else {
            process.env.NODE_TLS_REJECT_UNAUTHORIZED = originalReject;
        }
    }
});

// --- MÓDULO DE PRODUTOS ---

app.get('/api/produtos', async (req, res) => {
    try {
        const { start, end, marca, apenasOportunidade } = req.query;
        const { dStart, dEnd } = getDefaultDates(start, end);

        let where = `v.emissao BETWEEN ? AND ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20'`;
        const params = [dStart, dEnd];

        if (marca && marca !== 'Todos') {
            where += ` AND COALESCE(e.marca, v.marca) = ?`;
            params.push(marca);
        }

        const d1 = new Date(dStart);
        const d2 = new Date(dEnd);
        const diffDays = Math.ceil(Math.abs(d2 - d1) / (1000 * 60 * 60 * 24)) + 1;
        const periodDays = Math.max(1, diffDays);

        // Listagem de Performance com Métricas de Ruptura Otimizada
        const items = await query(`
            SELECT 
                v.produto_id as id,
                MAX(v.ean) as ean,
                MAX(v.descricao_produto) as nome,
                MAX(COALESCE(e.marca, v.marca)) as marca,
                MAX(e.image_url) as image_url,
                MAX(COALESCE(e.saldo, 0)) as saldo,
                MAX(e.previsao) as previsao,
                SUM(v.valor_total) as faturamento,
                SUM(v.quantidade) as qtd_vendida,
                COUNT(DISTINCT v.cnpj) as clientes_unicos,
                MAX(g.giro_periodo) as giro_periodo
            FROM vendas v
            LEFT JOIN estoque e ON v.produto_id = e.cod_produto
            LEFT JOIN (
                SELECT produto_id, SUM(quantidade) as giro_periodo
                FROM vendas
                WHERE emissao BETWEEN ? AND ?
                AND (status = '5' OR status = '6')
                GROUP BY produto_id
            ) g ON g.produto_id = v.produto_id
            WHERE ${where}
            GROUP BY v.produto_id
            ORDER BY faturamento DESC
        `, [dStart, dEnd, ...params]);

        // Processamento de DOH e Filtros Avançados
        let processedItems = items.map(item => {
            const fat = parseFloat(item.faturamento) || 0;
            const qtd = parseFloat(item.qtd_vendida) || 0;
            const sld = parseFloat(item.saldo) || 0;
            const gp = parseFloat(item.giro_periodo) || 0;
            const giroDiario = gp / periodDays;
            const doh = giroDiario > 0 ? (sld / giroDiario) : (sld > 0 ? 999 : 0);

            // Lógica de Previsão de Chegada
            let statusPrevisao = 'Sem Previsão';
            let previsaoAtrasada = false;
            if (item.previsao && item.previsao !== 'null') {
                statusPrevisao = `📅 Entrada: ${item.previsao}`;
                
                // Converter string DD/MM/AAAA para Date
                let pDate;
                if (item.previsao.includes('/')) {
                    const [d, m, y] = item.previsao.split('/');
                    pDate = new Date(y, m - 1, d);
                } else {
                    pDate = new Date(item.previsao);
                }

                if (!isNaN(pDate)) {
                    const esgotamento = new Date(BASE_DATE);
                    esgotamento.setDate(esgotamento.getDate() + Math.round(doh));
                    const limite = new Date(esgotamento);
                    limite.setDate(limite.getDate() + 30);
                    if (pDate > limite) previsaoAtrasada = true;
                }
            }

            return { 
                ...item, 
                faturamento: fat,
                qtd_vendida: qtd,
                saldo: sld,
                giro_30d: gp,
                doh,
                statusPrevisao,
                previsaoAtrasada
            };
        });

        if (apenasOportunidade === 'true') {
            processedItems = processedItems.filter(i => i.doh < 45 && i.saldo > 0);
        }

        // Cálculo ABC
        const totalFat = processedItems.reduce((acc, i) => acc + i.faturamento, 0);
        const totalPecas = processedItems.reduce((acc, i) => acc + i.qtd_vendida, 0);
        let cumulative = 0;
        const itemsABC = processedItems.map(item => {
            cumulative += item.faturamento;
            const perc = (cumulative / (totalFat || 1)) * 100;
            let abc = 'C';
            if (perc <= 70) abc = 'A';
            else if (perc <= 90) abc = 'B';
            return { ...item, classe_abc: abc };
        });

        // Dashboard Stats
        const stats = {
            faturamentoTotal: totalFat,
            skusVendidos: processedItems.length,
            volumePecas: totalPecas,
            ticketMedioSku: totalFat / (totalPecas || 1)
        };

        res.json({ stats, items: itemsABC });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/produtos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { start, end } = req.query;
        const { dStart, dEnd } = getDefaultDates(start, end);

        const product = await query(`SELECT cod_produto, descricao, marca, categoria, image_url, saldo, pv FROM estoque WHERE cod_produto = ?`, [id]);
        
        const topCompradores = await query(`
            SELECT 
                v.cnpj, 
                v.nome_cliente, 
                SUM(v.valor_total) as total,
                SUM(v.quantidade) as qtd,
                MAX(v.emissao) as ultima_compra
            FROM vendas v
            WHERE v.produto_id = ? AND v.emissao BETWEEN ? AND ? AND (v.status = '5' OR v.status = '6') AND v.almox = '20'
            GROUP BY v.cnpj, v.nome_cliente
            ORDER BY total DESC
            LIMIT 10
        `, [id, dStart, dEnd]);

        res.json({ product: product[0], topCompradores });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// --- MÓDULO DE GESTÃO E AUDITORIA (ADMIN) ---

// Estatísticas Rápidas e Última Venda
app.get('/api/admin/stats', async (req, res) => {
    try {
        const lastSale = await query(`SELECT MAX(emissao) as data FROM vendas`);
        const totalVendas = await query(`SELECT COUNT(*) as total FROM vendas`);
        const totalEstoque = await query(`SELECT COUNT(*) as total FROM estoque`);
        const totalMarcas = await query(`SELECT COUNT(*) as total FROM marcas_mestre`);
        
        res.json({
            lastSale: lastSale[0]?.data || '-',
            vendas: totalVendas[0]?.total,
            estoque: totalEstoque[0]?.total,
            marcas: totalMarcas[0]?.total
        });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// Listagem de Vendas Brutas com Filtros
app.get('/api/admin/vendas', async (req, res) => {
    try {
        const { search, start, end, page = 1, export: isExport } = req.query;
        const limit = 50;
        const offset = (page - 1) * limit;

        let where = "WHERE 1=1";
        let params = [];

        if (search) {
            where += " AND (num_docto LIKE ? OR cnpj LIKE ? OR nome_cliente LIKE ?)";
            params.push(`%${search}%`, `%${search}%`, `%${search}%`);
        }
        if (start && end) {
            where += " AND emissao BETWEEN ? AND ?";
            params.push(start, end);
        }

        const limitClause = isExport === 'true' ? "" : `LIMIT ${limit} OFFSET ${offset}`;
        const rows = await query(`SELECT * FROM vendas ${where} ORDER BY emissao DESC ${limitClause}`, params);
        const total = await query(`SELECT COUNT(*) as total FROM vendas ${where}`, params);

        res.json({ data: rows, total: total[0].total, limit });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// Listagem de Estoque Bruto
app.get('/api/admin/estoque', async (req, res) => {
    try {
        const { search, page = 1, export: isExport } = req.query;
        const limit = 50;
        const offset = (page - 1) * limit;

        let where = "";
        let params = [];

        if (search) {
            where = "WHERE cod_produto LIKE ? OR ean LIKE ? OR descricao LIKE ?";
            params.push(`%${search}%`, `%${search}%`, `%${search}%`);
        }

        const limitClause = isExport === 'true' ? "" : `LIMIT ${limit} OFFSET ${offset}`;
        const rows = await query(`SELECT * FROM estoque ${where} ORDER BY cod_produto ASC ${limitClause}`, params);
        const totalResult = await query(`SELECT COUNT(*) as total FROM estoque ${where}`, params);
        res.json({ data: rows, total: totalResult[0].total, limit });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// Gestão de Marcas (Visão por SKU e Update em Cascata)
app.get('/api/admin/marcas', async (req, res) => {
    try {
        const { search, page = 1, export: isExport } = req.query;
        const limit = 50;
        const offset = (page - 1) * limit;
        let where = "";
        if (search) where = `WHERE cod_produto LIKE '%${search}%' OR descricao LIKE '%${search}%' OR marca LIKE '%${search}%'`;
        
        const limitClause = isExport === 'true' ? "" : `LIMIT ${limit} OFFSET ${offset}`;
        const data = await query(`
            SELECT 
                cod_produto, 
                descricao, 
                marca
            FROM estoque 
            ${where}
            ORDER BY cod_produto ASC
            ${limitClause}
        `);
        const totalResult = await query(`SELECT COUNT(*) as total FROM estoque ${where}`);
        res.json({ data, total: totalResult[0].total, limit });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.get('/api/admin/clientes', async (req, res) => {
    try {
        const { search, page = 1, export: isExport } = req.query;
        const limit = 50;
        const offset = (page - 1) * limit;
        let where = "";
        let params = [];
        if (search) {
            where = `WHERE cliente_id LIKE ? OR nome_cliente LIKE ? OR perfil LIKE ?`;
            params.push(`%${search}%`, `%${search}%`, `%${search}%`);
        }
        
        const limitClause = isExport === 'true' ? "" : `LIMIT ${limit} OFFSET ${offset}`;
        const data = await query(`
            SELECT cliente_id, nome_cliente, perfil
            FROM clientes_perfil 
            ${where}
            ORDER BY nome_cliente ASC
            ${limitClause}
        `, params);
        
        const totalResult = await query(`SELECT COUNT(*) as total FROM clientes_perfil ${where}`, params);
        res.json({ data, total: totalResult[0].total, limit });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

// Gestão de Fotos (Thumbnail, Código, URL)
app.get('/api/admin/fotos', async (req, res) => {
    try {
        const { search, page = 1, export: isExport } = req.query;
        const limit = 50;
        const offset = (page - 1) * limit;
        let where = "";
        let params = [];
        if (search) {
            where = `WHERE cod_produto LIKE ? OR descricao LIKE ? OR image_url LIKE ?`;
            params.push(`%${search}%`, `%${search}%`, `%${search}%`);
        }
        
        const limitClause = isExport === 'true' ? "" : `LIMIT ${limit} OFFSET ${offset}`;
        const data = await query(`
            SELECT cod_produto, descricao, image_url
            FROM estoque 
            ${where}
            ORDER BY cod_produto ASC
            ${limitClause}
        `, params);
        
        const totalResult = await query(`SELECT COUNT(*) as total FROM estoque ${where}`, params);
        res.json({ data, total: totalResult[0].total, limit });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put('/api/admin/produtos/:id/foto', async (req, res) => {
    try {
        const { id } = req.params;
        const { novaFoto } = req.body;
        await query('UPDATE estoque SET image_url = ? WHERE cod_produto = ?', [novaFoto, id]);
        res.json({ message: "URL da foto atualizada!" });
    } catch (err) { res.status(500).json({ error: err.message }); }
});


app.put('/api/admin/clientes/:id/perfil', async (req, res) => {
    try {
        const { id } = req.params;
        const { novoPerfil } = req.body;
        await query('UPDATE clientes_perfil SET perfil = ? WHERE cliente_id = ?', [novoPerfil, id]);
        res.json({ message: "Perfil atualizado!" });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.delete('/api/admin/vendas/clear', async (req, res) => {
    try {
        await query('DELETE FROM vendas');
        res.json({ message: "Base de vendas limpa com sucesso!" });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.put('/api/admin/produtos/:id/marca', async (req, res) => {
    const { id } = req.params;
    const { novaMarca } = req.body;
    try {
        await query('UPDATE estoque SET marca = ? WHERE cod_produto = ?', [novaMarca, id]);
        await query('UPDATE vendas SET marca = ? WHERE produto_id = ?', [novaMarca, id]);
        res.json({ message: "Marca do produto atualizada com sucesso!" });
    } catch (err) { res.status(500).json({ error: err.message }); }
});

app.post('/api/import/marcas', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Arquivo não enviado");
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);
        
        let added = 0;
        let updated = 0;
        
        // Função de normalização para busca de colunas
        const normalize = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

        for (const row of data) {
            const getV = (targets) => {
                const normTargets = targets.map(normalize);
                const col = Object.keys(row).find(k => normTargets.some(t => normalize(k) === t)) ||
                            Object.keys(row).find(k => normTargets.some(t => normalize(k).includes(t)));
                return col ? row[col] : null;
            };

            const codigo = sanitize(getV(['codigo', 'sku', 'codproduto', 'cód. produto']));
            const descricao = toTitleCaseClean(getV(['descricao', 'nome']));
            const marca = toTitleCaseClean(getV(['marca']));
            
            if (codigo && descricao && marca) {
                // Verifica se já existe
                const existing = await query('SELECT cod_produto FROM estoque WHERE cod_produto = ?', [codigo]);
                if (existing.length === 0) {
                    await query('INSERT INTO estoque (cod_produto, descricao, marca, saldo, pv, pdv) VALUES (?, ?, ?, 0, 0, 0)', [codigo, descricao, marca]);
                    added++;
                } else {
                    await query('UPDATE estoque SET descricao = COALESCE(?, descricao), marca = COALESCE(?, marca) WHERE cod_produto = ?', [descricao, marca, codigo]);
                    updated++;
                }
            }
        }
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({ message: `Sucesso: ${added} novas marcas cadastradas. Atualizados: ${updated} itens já existentes.`, added, updated, total: data.length });
    } catch (err) { 
        console.error(`[Import Marcas] Erro:`, err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: err.message }); 
    }
});

app.post('/api/import/clientes', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) throw new Error("Arquivo não enviado");
        const workbook = xlsx.readFile(req.file.path);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);
        
        let added = 0;
        let updated = 0;
        
        const normalize = (s) => String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
        const validProfiles = ["Shopping", "Rua", "Magazine", "Especializada"];

        for (const row of data) {
            const getV = (targets) => {
                const normTargets = targets.map(normalize);
                const col = Object.keys(row).find(k => normTargets.some(t => normalize(k) === t)) ||
                            Object.keys(row).find(k => normTargets.some(t => normalize(k).includes(t)));
                return col ? row[col] : null;
            };

            const clienteId = sanitize(getV(['id', 'cliente', 'codigo']));
            const nome = toTitleCaseClean(getV(['nome', 'razao', 'cliente']));
            let perfilRaw = toTitleCaseClean(getV(['perfil', 'tipo', 'loja']));
            
            // Format to exact match
            let perfil = "Rua"; // default
            if (perfilRaw) {
                const found = validProfiles.find(p => p.toLowerCase() === perfilRaw.toLowerCase());
                if (found) perfil = found;
                else if (perfilRaw.includes("shop")) perfil = "Shopping";
                else if (perfilRaw.includes("mag")) perfil = "Magazine";
                else if (perfilRaw.includes("esp")) perfil = "Especializada";
                else perfil = "Rua"; // default fallback for unmatched
            }

            if (clienteId) {
                const existing = await query('SELECT cliente_id FROM clientes_perfil WHERE cliente_id = ?', [clienteId]);
                if (existing.length === 0) {
                    await query('INSERT INTO clientes_perfil (cliente_id, nome_cliente, perfil) VALUES (?, ?, ?)', [clienteId, nome, perfil]);
                    added++;
                } else {
                    await query('UPDATE clientes_perfil SET perfil = ?, nome_cliente = COALESCE(?, nome_cliente) WHERE cliente_id = ?', [perfil, nome, clienteId]);
                    updated++;
                }
            }
        }
        if (fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.json({ message: `Sucesso: ${added} perfis criados. Atualizados: ${updated} clientes já existentes.`, added, updated, total: data.length });
    } catch (err) { 
        console.error(`[Import Clientes] Erro:`, err);
        if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
        res.status(500).json({ error: err.message }); 
    }
});

// Servir arquivos estáticos do frontend (React) em produção
const frontendDistPath = path.resolve(__dirname, '../frontend/dist');
if (fs.existsSync(frontendDistPath)) {
    app.use(express.static(frontendDistPath));
    app.get(/.*/, (req, res, next) => {
        if (req.path.startsWith('/api')) {
            return next();
        }
        res.sendFile(path.join(frontendDistPath, 'index.html'));
    });
}

