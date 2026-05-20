const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('sunny.sqlite');
db.all("SELECT cliente_id, nome_cliente, cnpj FROM vendas LIMIT 20", (err, rows) => {
    console.log(JSON.stringify(rows, null, 2));
    db.close();
});
