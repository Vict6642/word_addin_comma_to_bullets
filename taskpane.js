Office.onReady(() => {
  document.getElementById('convert').onclick = convertSelectionToColumns;
  document.getElementById('convert-selection').onclick = convertSelectionToColumns;
});

function setStatus(msg) {
  const s = document.getElementById('status');
  s.innerText = msg;
}

async function convertSelectionToColumns() {
  const maxChars = Math.max(50, Number(document.getElementById('maxChars').value) || 200);
  setStatus('Kører...');

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load('text');
      await context.sync();

      const raw = range.text || '';
      if (!raw.trim()) {
        setStatus('Ingen tekst fundet i markeringen. Sørg for at markere teksten inde i tekstfeltet.');
        return;
      }

      const items = raw.split(/[,;\n\r]+/).map(s => s.trim()).filter(s => s.length>0);
      if (items.length === 0) {
        setStatus('Ingen punkter fundet efter opdeling.');
        return;
      }

      const totalChars = items.reduce((a,b) => a + b.length, 0);
      let cols = Math.ceil(totalChars / maxChars);
      cols = Math.min(Math.max(cols, 1), 3); // begræns 1..3

      const rows = Math.ceil(items.length / cols);
      const table = Array.from({length: rows}, () => Array.from({length: cols}, ()=>''));

      for (let i=0; i<items.length; i++) {
        const col = Math.floor(i/rows);
        const row = i % rows;
        table[row][col] = items[i];
      }

      const insertedTable = range.insertTable(rows, cols, Word.InsertLocation.replace, table);
      insertedTable.load();
      await context.sync();

      for (let r=0; r<rows; r++){
        for (let c=0; c<cols; c++){
          const cell = insertedTable.getCell(r,c);
          cell.body.load('paragraphs');
        }
      }
      await context.sync();

      for (let r=0; r<rows; r++){
        for (let c=0; c<cols; c++){
          const text = table[r][c];
          const cell = insertedTable.getCell(r,c);
          cell.body.clear();
          if (!text) continue;
          const parts = text.split(/\n+/).map(s=>s.trim()).filter(s=>s.length>0);
          for (let i=0;i<parts.length;i++){
            const p = cell.body.insertParagraph(parts[i], Word.InsertLocation.end);
            p.font.size = 11;
          }
          cell.body.paragraphs.load('items');
        }
      }
      await context.sync();

      for (let r=0; r<rows; r++){
        for (let c=0; c<cols; c++){
          const cell = insertedTable.getCell(r,c);
          const paras = cell.body.paragraphs;
          paras.load('items');
        }
      }
      await context.sync();

      for (let r=0; r<rows; r++){
        for (let c=0; c<cols; c++){
          const cell = insertedTable.getCell(r,c);
          const paras = cell.body.paragraphs.items;
          if (paras.length === 0) continue;
          for (let i=0;i<paras.length;i++){
            paras[i].listItem = { level: 0, listType: 'Bullet' };
          }
        }
      }

      await context.sync();
      setStatus(`Konverteret ${items.length} punkter til ${cols} kolonne(r).`);
    });
  }
  catch (e) {
    setStatus('Fejl: ' + (e && e.message ? e.message : JSON.stringify(e)));
  }
}
