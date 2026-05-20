const sqlite3 = require('sqlite3').verbose();
const { Pool } = require('pg');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const dbPath = path.resolve(__dirname, 'sunny.sqlite');
const pgUrl = process.argv[2] || process.env.DATABASE_URL;

if (!pgUrl) {
    console.error("\n❌ ERRO: DATABASE_URL não fornecida!");
    console.log("Por favor, forneça a URL de conexão do PostgreSQL como argumento ou no arquivo .env");
    console.log("Exemplo: node migrate_sqlite_to_pg.js \"postgresql://user:pass@host:port/dbname\"\n");
    process.exit(1);
}

if (!fs.existsSync(dbPath)) {
    console.error(`\n❌ ERRO: Banco SQLite local não encontrado em: ${dbPath}`);
    process.exit(1);
}

const sqliteDb = new sqlite3.Database(dbPath);
const pgPool = new Pool({
    connectionString: pgUrl,
    ssl: { rejectUnauthorized: false }
});

const getSqliteData = (query, params = []) => {
    return new Promise((resolve, reject) => {
        sqliteDb.all(query, params, (err, rows) => {
            if (err) reject(err);
            else resolve(rows);
        });
    });
};

async function migrateTable(tableName, pgInsertQuery, conflictResolution = '', pgClient) {
    console.log(`\n📦 Migrando tabela: ${tableName}...`);
    
    // Obter total de registros
    const countRes = await getSqliteData(`SELECT count(*) as count FROM ${tableName}`).catch(() => [{ count: 0 }]);
    const totalRows = countRes[0].count;
    
    if (totalRows === 0) {
        console.log(`⚠️ Tabela ${tableName} está vazia ou não existe localmente. Pulando.`);
        return;
    }
    
    console.log(`   Total de registros locais: ${totalRows}`);
    
    // Obter todas as colunas da tabela do SQLite para fazer o SELECT correto
    const sqliteRows = await getSqliteData(`SELECT * FROM ${tableName}`);
    
    // Preparar inserção em lotes (batching) de 500 registros para alta performance
    const batchSize = 500;
    let successCount = 0;
    
    for (let i = 0; i < sqliteRows.length; i += batchSize) {
        const batch = sqliteRows.slice(i, i + batchSize);
        
        await pgClient.query('BEGIN');
        try {
            for (const row of batch) {
                const keys = Object.keys(row);
                // Filtrar colunas geradas automaticamente (como id/serial) se houver
                const filteredKeys = keys.filter(k => k !== 'id');
                const values = filteredKeys.map(k => row[k]);
                
                const placeholders = filteredKeys.map((_, idx) => `$${idx + 1}`).join(', ');
                const columns = filteredKeys.join(', ');
                
                const queryStr = `
                    INSERT INTO ${tableName} (${columns}) 
                    VALUES (${placeholders}) 
                    ${conflictResolution}
                `;
                
                await pgClient.query(queryStr, values);
                successCount++;
            }
            await pgClient.query('COMMIT');
        } catch (err) {
            await pgClient.query('ROLLBACK');
            console.error(`❌ Erro no lote da tabela ${tableName}:`, err.message);
            throw err;
        }
        
        const progress = Math.min(100, ((i + batch.length) / totalRows) * 100);
        process.stdout.write(`   Progresso: ${progress.toFixed(1)}% (${successCount}/${totalRows})\r`);
    }
    console.log(`\n✅ Tabela ${tableName} migrada com sucesso! (${successCount} registros inseridos)`);
}

async function initPgDb(pgClient) {
    console.log("🛠️ Inicializando tabelas no PostgreSQL se não existirem...");
    await pgClient.query(`
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
            ean TEXT
        )
    `);
    await pgClient.query(`
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
    await pgClient.query(`
        CREATE TABLE IF NOT EXISTS clientes_perfil (
            cliente_id TEXT PRIMARY KEY,
            nome_cliente TEXT,
            perfil TEXT
        )
    `);
    await pgClient.query(`
        CREATE TABLE IF NOT EXISTS marcas_mestre (
            id SERIAL PRIMARY KEY,
            nome TEXT UNIQUE
        )
    `);
    await pgClient.query(`
        CREATE TABLE IF NOT EXISTS marcas_legado (
            id SERIAL PRIMARY KEY,
            marca TEXT,
            produto_id TEXT
        )
    `);
    
    await pgClient.query(`CREATE INDEX IF NOT EXISTS idx_cliente ON vendas(cliente_id)`);
    await pgClient.query(`CREATE INDEX IF NOT EXISTS idx_vendedor ON vendas(vendedor_id)`);
    await pgClient.query(`CREATE INDEX IF NOT EXISTS idx_status ON vendas(status)`);
    await pgClient.query(`CREATE INDEX IF NOT EXISTS idx_emissao ON vendas(emissao)`);
    console.log("✅ Tabelas inicializadas no PostgreSQL.");
}

async function startMigration() {
    console.log("🚀 Iniciando migração de dados do SQLite para PostgreSQL...");
    
    const pgClient = await pgPool.connect();
    
    try {
        // Garantir criação das tabelas antes da migração
        await initPgDb(pgClient);

        // 1. Marcas Mestre
        await migrateTable(
            'marcas_mestre',
            null,
            'ON CONFLICT (nome) DO NOTHING',
            pgClient
        ).catch(err => console.log("   (Aviso marcas_mestre):", err.message));
        
        // 2. Marcas Legado
        await migrateTable(
            'marcas_legado',
            null,
            '',
            pgClient
        ).catch(err => console.log("   (Aviso marcas_legado):", err.message));

        // 3. Clientes Perfil
        await migrateTable(
            'clientes_perfil',
            null,
            'ON CONFLICT (cliente_id) DO UPDATE SET nome_cliente = EXCLUDED.nome_cliente, perfil = EXCLUDED.perfil',
            pgClient
        ).catch(err => console.log("   (Aviso clientes_perfil):", err.message));

        // 4. Estoque
        await migrateTable(
            'estoque',
            null,
            'ON CONFLICT (cod_produto) DO UPDATE SET saldo = EXCLUDED.saldo, pv = EXCLUDED.pv, pdv = EXCLUDED.pdv, image_url = EXCLUDED.image_url, categoria = EXCLUDED.categoria',
            pgClient
        ).catch(err => console.log("   (Aviso estoque):", err.message));

        // 5. Vendas (Pode ser muito grande)
        await migrateTable(
            'vendas',
            null,
            '',
            pgClient
        ).catch(err => console.log("   (Aviso vendas):", err.message));

        console.log("\n🎉 MIGRAÇÃO CONCLUÍDA COM SUCESSO!");
        
    } catch (err) {
        console.error("\n❌ ERRO FATAL NA MIGRAÇÃO:", err);
    } finally {
        sqliteDb.close();
        pgClient.release();
        await pgPool.end();
    }
}

startMigration();
