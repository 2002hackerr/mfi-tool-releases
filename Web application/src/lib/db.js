import Dexie from 'dexie';

// Define the local IndexedDB database schema
export const db = new Dexie('MFIDatabase');

db.version(1).stores({
  // ++id gives an auto-incrementing primary key
  claims: '++id, barcode, sid, po, asin, [sid+po+asin]',
  rebni: '++id, sid, po, asin, [sid+po+asin]',
  // invoice_search is massive, index heavily queried fields
  invoice_search: '++id, invoice, sid, po, asin, [sid+po+asin]'
});

export async function clearDatabase() {
  await db.claims.clear();
  await db.rebni.clear();
  await db.invoice_search.clear();
}
