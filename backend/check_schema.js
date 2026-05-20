const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const dbPath = path.resolve(__dirname, 'sunny.sqlite');
const db = new sqlite3.Database(dbPath);

db.all("PRAGMA table_info(vendas)", [], (err, rows) => {
    if (err) console.error(err);
    else console.log(JSON.stringify(rows, null, 2));
    db.close();
});
