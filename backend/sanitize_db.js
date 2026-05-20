const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const db = new sqlite3.Database(path.resolve(__dirname, 'sunny.sqlite'));

db.serialize(() => {
    console.log('Iniciando sanitização de base histórica...');
    
    // SQLite não suporta regex avançado nativamente, vamos fazer via JS para garantir
    db.all("SELECT id, produto_id, num_docto FROM vendas", (err, rows) => {
        if (err) return console.error(err);
        
        db.run("BEGIN TRANSACTION");
        rows.forEach(row => {
            const cleanProd = String(row.produto_id || '').trim().replace(/^0+/, '');
            const cleanDoc = String(row.num_docto || '').trim().replace(/^0+/, '');
            db.run("UPDATE vendas SET produto_id = ?, num_docto = ? WHERE id = ?", [cleanProd, cleanDoc, row.id]);
        });
        db.run("COMMIT", () => {
            console.log('Sanitização de Vendas concluída!');
            
            // Agora Estoque
            db.all("SELECT id, cod_produto FROM estoque", (err, rows) => {
                if (err) return;
                db.run("BEGIN TRANSACTION");
                rows.forEach(row => {
                    const clean = String(row.cod_produto || '').trim().replace(/^0+/, '');
                    db.run("UPDATE estoque SET cod_produto = ? WHERE id = ?", [clean, row.id]);
                });
                db.run("COMMIT", () => {
                    console.log('Sanitização de Estoque concluída!');
                    db.close();
                });
            });
        });
    });
});
