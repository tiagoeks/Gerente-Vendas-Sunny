const sqlite3 = require('sqlite3').verbose();
const path = require('path');

const dbPath = path.resolve(__dirname, 'sunny.sqlite');
const FIXED_WHERE = "(status = '5' OR status = '6') AND almox = '20'";
const id = '010639';

const db = new sqlite3.Database(dbPath);

const sql = `
    SELECT 
        descricao_produto as produto, 
        SUM(quantidade) as qtd,
        SUM(valor_total) as fat,
        ROUND(SUM(valor_total) / NULLIF(SUM(quantidade), 0), 2) as preco_medio,
        MAX(emissao) as ultima_compra
    FROM vendas 
    WHERE cliente_id = ? AND ${FIXED_WHERE} 
    GROUP BY produto_id, descricao_produto 
    ORDER BY fat DESC
    LIMIT 5
`;

db.all(sql, [id], (err, rows) => {
    if (err) {
        console.error(err);
    } else {
        console.log(JSON.stringify(rows, null, 2));
    }
    db.close();
});
