const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.resolve('c:/Users/Tiago/Desktop/Dev/Miguel/Vendas/sunny/backend/sunny.sqlite');
const db = new sqlite3.Database(dbPath);
db.all("SELECT emissao FROM vendas LIMIT 10", [], (err, rows) => {
    if (err) console.error(err);
    else console.log(JSON.stringify(rows, null, 2));
    db.close();
});
