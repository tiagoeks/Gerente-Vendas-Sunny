const xlsx = require('xlsx');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const path = require('path');

const excelPath = path.resolve(__dirname, '../../Bando de Dados/Base_Padrão (Sell_In).xlsx');
const dbPath = path.resolve(__dirname, 'sunny.sqlite');

if (!fs.existsSync(excelPath)) {
    console.error("Excel file not found at " + excelPath);
    process.exit(1);
}

const workbook = xlsx.readFile(excelPath);
const sheetName = 'Base_Padrão (Sell_In)';

if (!workbook.Sheets[sheetName]) {
    console.error("Sheet not found!");
    process.exit(1);
}

// Use raw: true to get serial dates, then convert them properly
const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { raw: true });

const db = new sqlite3.Database(dbPath);

// Helper function to convert Excel serial date to YYYY-MM-DD
function excelDateToISO(serial) {
    if (!serial) return '1970-01-01';
    
    // If it's already a string like "01/04/2026"
    if (typeof serial === 'string') {
        const parts = serial.split('/');
        if (parts.length === 3) {
            // Convert DD/MM/YYYY to YYYY-MM-DD
            const day = parts[0].padStart(2, '0');
            const month = parts[1].padStart(2, '0');
            const year = parts[2];
            return `${year}-${month}-${day}`;
        }
        const d = new Date(serial);
        if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
        return '1970-01-01';
    }
    
    // If it's a serial number
    const utc_days  = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    return date_info.toISOString().split('T')[0];
}

db.serialize(() => {
    db.run(`DROP TABLE IF EXISTS vendas`);
    db.run(`
        CREATE TABLE vendas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
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

    db.run(`CREATE INDEX idx_cliente ON vendas(cliente_id)`);
    db.run(`CREATE INDEX idx_status ON vendas(status)`);
    db.run(`CREATE INDEX idx_vendedor ON vendas(vendedor_id)`);
    db.run(`CREATE INDEX idx_gerente ON vendas(gerente_id)`);
    db.run(`CREATE INDEX idx_emissao ON vendas(emissao)`);

    const stmt = db.prepare(`
        INSERT INTO vendas (
            num_docto, emissao, mes, uf, status, cliente_id, loja, nome_cliente,
            cnpj, pedido_cliente, quantidade, valor_unitario, valor_total,
            almox, produto_id, descricao_produto, vendedor_id, nome_vendedor,
            gerente_id, marca, ean
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    let count = 0;
    let ignored = 0;
    
    for (const row of data) {
        let status = String(row['STATUS'] || '').trim();
        let almox = String(row['Almox.'] || '').trim();
        
        if ((status === '5' || status === '6') && almox === '20') {
            const vTotal = parseFloat(row['Vlr.Total']) || 0;
            const vUnit = parseFloat(row['Vlr.Unitario']) || 0;
            const qtd = parseFloat(row['Quantidade']) || 0;
            const rawEmissao = row['Emissao'];
            const isoEmissao = excelDateToISO(rawEmissao);
            const mesStr = isoEmissao.substring(0, 7);

            stmt.run([
                row['Num. Docto.'], isoEmissao, mesStr, row['U.F.'], status, row['Cliente'], row['Loja'], row['Nome'],
                row['C.N.P.J.'], row['Pedido Cliente'], qtd, vUnit, vTotal, almox, row['Produto'], row['Descricao'],
                row['Vendedor'], row['Nome Vendedor'], String(row['Gerente'] || ''), row['Marca'], row['EAN']
            ]);
            count++;
        } else {
            ignored++;
            if (ignored < 5) console.log(`Ignored row: Status=${status}, Almox=${almox}`);
        }
    }

    stmt.finalize();
    console.log(`Database updated. Inserted: ${count}, Ignored: ${ignored}`);
});

db.close();
