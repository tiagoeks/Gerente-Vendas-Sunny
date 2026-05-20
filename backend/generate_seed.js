const xlsx = require('xlsx');
const fs = require('fs');

const excelPath = '../Bando de Dados/Base_Padrão (Sell_In).xlsx';

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

const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

const createTableSQL = `
CREATE TABLE IF NOT EXISTS vendas (
    id SERIAL PRIMARY KEY,
    num_docto VARCHAR(50),
    emissao VARCHAR(50),
    uf VARCHAR(5),
    status VARCHAR(10),
    cliente_id VARCHAR(50),
    loja VARCHAR(20),
    nome_cliente VARCHAR(255),
    cnpj VARCHAR(20),
    pedido_cliente VARCHAR(100),
    quantidade NUMERIC,
    valor_unitario NUMERIC(10,2),
    valor_total NUMERIC(10,2),
    almox VARCHAR(20),
    produto_id VARCHAR(50),
    descricao_produto TEXT,
    vendedor_id VARCHAR(50),
    nome_vendedor VARCHAR(255),
    gerente_id VARCHAR(50),
    marca VARCHAR(100),
    ean VARCHAR(50)
);

-- Creating indexes for performance
CREATE INDEX IF NOT EXISTS idx_cliente ON vendas(cliente_id);
CREATE INDEX IF NOT EXISTS idx_vendedor ON vendas(vendedor_id);
CREATE INDEX IF NOT EXISTS idx_status ON vendas(status);

TRUNCATE TABLE vendas;
\n`;

let insertSQL = `INSERT INTO vendas (
    num_docto, emissao, uf, status, cliente_id, loja, nome_cliente,
    cnpj, pedido_cliente, quantidade, valor_unitario, valor_total,
    almox, produto_id, descricao_produto, vendedor_id, nome_vendedor,
    gerente_id, marca, ean
) VALUES \n`;

const values = [];

for (const row of data) {
    // Apply rules "Status IN (5,6)" and "Almox <> 20"
    const status = row['STATUS'] || '';
    const almox = row['Almox.'] || '';
    
    // Convert to string to avoid comparison issues (almox might be numeric 20 in excel but rule says Almox <> 20)
    const statusStr = String(status).trim();
    const almoxStr = String(almox).trim();

    if ((statusStr === '5' || statusStr === '6') && almoxStr !== '20') {
      const escape = (val) => val ? "'" + String(val).replace(/'/g, "''") + "'" : "NULL";
      values.push(`(
        ${escape(row['Num. Docto.'])}, ${escape(row['Emissao'])}, ${escape(row['U.F.'])}, ${escape(row['STATUS'])}, 
        ${escape(row['Cliente'])}, ${escape(row['Loja'])}, ${escape(row['Nome'])}, ${escape(row['C.N.P.J.'])}, 
        ${escape(row['Pedido Cliente'])}, ${Number(row['Quantidade']) || 0}, ${Number(row['Vlr.Unitario']) || 0}, 
        ${Number(row['Vlr.Total']) || 0}, ${escape(row['Almox.'])}, ${escape(row['Produto'])}, 
        ${escape(row['Descricao'])}, ${escape(row['Vendedor'])}, ${escape(row['Nome Vendedor'])}, 
        ${escape(row['Gerente'])}, ${escape(row['Marca'])}, ${escape(row['EAN'])}
      )`);
    }
}

// Write in chunks to avoid syntax errors if there are too many variables natively
const CHUNK_SIZE = 1000;
let finalSQL = createTableSQL;

for (let i = 0; i < values.length; i += CHUNK_SIZE) {
    const chunk = values.slice(i, i + CHUNK_SIZE);
    finalSQL += insertSQL + chunk.join(',\n') + ';\n\n';
}

fs.writeFileSync('seed.sql', finalSQL, 'utf8');
console.log("seed.sql generated successfully with " + values.length + " valid records based on rules.");
