const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('sunny.sqlite');
db.serialize(() => {
    db.all("SELECT name FROM sqlite_master WHERE type='table'", (e, r) => {
        if (e) return console.error(e);
        r.forEach(t => {
            db.all(`PRAGMA table_info(${t.name})`, (e, i) => {
                console.log(t.name);
                console.log(i);
            });
        });
    });
});
