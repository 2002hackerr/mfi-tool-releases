import Papa from 'papaparse';
import * as ExcelJS from 'exceljs';
import { db, clearDatabase } from '../lib/db';

async function processFile(file, table) {
  self.postMessage({ type: 'PROGRESS', table, status: 'Starting parse...' });

  if (file.name.toLowerCase().endsWith('.csv')) {
    await processCSV(file, table);
  } else if (file.name.toLowerCase().match(/\.xlsx?$/)) {
    await processExcel(file, table);
  } else {
    throw new Error(`Unsupported file type for ${file.name}`);
  }
}

async function processCSV(file, table) {
  return new Promise((resolve, reject) => {
    let chunkCount = 0;
    
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      chunk: async (results, parser) => {
        parser.pause(); // Pause streaming to wait for DB write
        
        try {
          // Normalize headers slightly by trimming spaces from keys
          const normalized = results.data.map(row => {
            const newRow = {};
            for (let key in row) {
              newRow[key.trim()] = row[key];
            }
            return newRow;
          });
          
          await db[table].bulkAdd(normalized);
          chunkCount += results.data.length;
          
          self.postMessage({ type: 'PROGRESS', table, status: `Parsed ${chunkCount} rows...` });
          
          parser.resume();
        } catch (err) {
          parser.abort();
          reject(err);
        }
      },
      complete: () => {
        self.postMessage({ type: 'COMPLETE', table, count: chunkCount });
        resolve();
      },
      error: (err) => reject(err)
    });
  });
}

async function processExcel(file, table) {
  self.postMessage({ type: 'PROGRESS', table, status: 'Reading entire Excel file (Warning: Requires memory)...' });
  const arrayBuffer = await file.arrayBuffer();
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  
  const worksheet = workbook.worksheets[0]; // first sheet
  if (!worksheet) throw new Error("Excel file is empty.");
  
  let rows = [];
  let headers = [];
  
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) {
      headers = row.values.map(v => (v ? v.toString().trim() : ''));
      // row.values is 1-indexed in exceljs!
    } else {
      const rowData = {};
      row.values.forEach((val, idx) => {
        if (idx > 0 && headers[idx]) {
          // Handle rich text or object values in excel cells
          rowData[headers[idx]] = typeof val === 'object' && val !== null ? (val.text || val.result || val) : val;
        }
      });
      rows.push(rowData);
    }
  });

  self.postMessage({ type: 'PROGRESS', table, status: `Bulk inserting ${rows.length} rows into database...` });
  
  // Bulk insert in chunks to avoid blocking the DB transaction too hard
  const chunkSize = 10000;
  for (let i = 0; i < rows.length; i += chunkSize) {
    await db[table].bulkAdd(rows.slice(i, i + chunkSize));
    self.postMessage({ type: 'PROGRESS', table, status: `Inserted ${Math.min(i + chunkSize, rows.length)} rows...` });
  }
  
  self.postMessage({ type: 'COMPLETE', table, count: rows.length });
}

self.onmessage = async (e) => {
  const { action, payload } = e.data;
  
  if (action === 'INIT_DB') {
    try {
      await clearDatabase();
      self.postMessage({ type: 'DB_CLEARED' });
    } catch (err) {
      self.postMessage({ type: 'ERROR', error: err.message });
    }
  } else if (action === 'PARSE_FILE') {
    try {
      const { file, table } = payload;
      await processFile(file, table);
    } catch (err) {
      self.postMessage({ type: 'ERROR', table: payload.table, error: err.message });
    }
  }
};
