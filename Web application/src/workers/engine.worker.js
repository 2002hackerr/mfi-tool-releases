import { db } from '../lib/db';

class InvestigationEngine {
  constructor() {
    this.cache_sid = {};
    this.cache_bc = {};
    this.loop_cache = new Map();
    this.user_overrides = {};
    this.MAX_DEPTH = 12;
    this.pendingUiPromise = null;
    this.stop_requested = false;
  }

  // Utility logic translations
  clean(s) {
    if (s === null || s === undefined) return "";
    return String(s).trim().toUpperCase();
  }

  extract_sid(s) {
    const val = this.clean(s);
    const match = val.match(/(FBA[A-Z0-9]+)/);
    return match ? match[1] : val.replace(/[^A-Z0-9]/g, '');
  }

  safe_num(val) {
    const n = parseFloat(String(val).replace(/,/g, ''));
    return isNaN(n) ? 0 : n;
  }

  fmt_qty(n) {
    if (n === null || n === undefined || n === '') return '';
    return Number(n) % 1 === 0 ? parseInt(n, 10).toString() : Number(n).toFixed(2);
  }

  // Await user response by passing a message to the main thread
  async requestUserAction(promptType, data) {
    return new Promise((resolve) => {
      this.pendingUiPromise = resolve;
      self.postMessage({ type: 'UI_PROMPT', promptType, data });
    });
  }
  
  // Accept user response from main thread
  resolveUserAction(response) {
    if (this.pendingUiPromise) {
      this.pendingUiPromise(response);
      this.pendingUiPromise = null;
    }
  }

  // Evaluate REBNI (Rule: ALWAYS use row 0, never sum)
  async evaluateRebni(sid, po, asin, claiming_pqv) {
    const resultRows = await db.rebni.where('[sid+po+asin]').equals([sid, po, asin]).toArray();
    if (!resultRows || resultRows.length === 0) {
      return { found: false, accounted: 0, status: null };
    }
    
    // Strict requirement: Always use first row only, never sum unpacked qty
    const firstRow = resultRows[0];
    const recQty = this.safe_num(firstRow.quantity_unpacked || firstRow['Qty Unpacked']);
    const exStatus = this.clean(firstRow.Exceptions || firstRow['Exception Status']);
    
    return {
      found: true,
      accounted: recQty,
      status: exStatus,
      raw_row: firstRow
    };
  }

  // Cross PO Detector
  async detectCrossPo(sid, current_po, asin) {
    // Queries all matching SID+ASIN regardless of PO
    const candidates = await db.rebni.where('sid').equals(sid).toArray();
    const matches = candidates.filter(r => this.clean(r.asin || r.ASIN) === asin && this.clean(r.po || r['PO']) !== current_po);
    
    let crossCands = [];
    for (let mtc of matches) {
      let rec = this.safe_num(mtc.quantity_unpacked || mtc['Qty Unpacked']);
      if (rec > 0) {
        crossCands.push({
          po: mtc.po || mtc['PO'],
          sid: mtc.sid || mtc['SID'],
          asin: mtc.asin || mtc['ASIN'],
          rec_qty: rec
        });
      }
    }
    return crossCands;
  }

  // Main Recursive Level Builder
  async buildOneLevel(barcode, inv, sid, po, asin, invQty, pqv, depth, isClaiming, loopContext = {}) {
    let rows = [];
    let remPqv = pqv;
    let foundMatches = [];

    // Base DOMINANT / CLAIMING Row
    const domRow = {
      barcode: isClaiming ? barcode : '[ASIN Match]',
      invoice: inv,
      sid: sid,
      po: po,
      asin: asin,
      inv_qty: this.fmt_qty(invQty),
      rec_qty: '',
      mtc_qty: '',
      mtc_inv: '',
      remarks: '',
      date: new Date().toISOString().split('T')[0],
      depth: depth,
      type: isClaiming ? 'claiming' : 'dominant'
    };

    // Evaluate REBNI
    const reb = await this.evaluateRebni(sid, po, asin, remPqv);
    let totalAccounted = reb.accounted; // Basic formula implementation
    
    if (reb.found && !isClaiming) {
       // Logic to determine Direct Shortage or Phase 1 resolution...
       if (totalAccounted >= remPqv) {
           domRow.remarks = 'Phase 1 - Fully Accounted';
           rows.push(domRow);
           return { rows, matches: [], remaining: 0 };
       }
    }
    
    // Core Engine Logic translates heavily here...
    // The query for Invoice Search Match
    const matchesList = await db.invoice_search.where('[sid+po+asin]').equals([sid, po, asin]).toArray();
    
    domRow.remarks = `Found ${matchesList.length} matches...`;
    rows.push(domRow);
    
    return { rows, matches: matchesList, remaining: remPqv };
  }
}

const engine = new InvestigationEngine();

self.onmessage = async (e) => {
  const { action, payload } = e.data;
  
  if (action === 'START_AUTO') {
     engine.stop_requested = false;
     const allClaims = await db.claims.toArray();
     
     for (let i = 0; i < allClaims.length; i++) {
        if (engine.stop_requested) break;
        const claim = allClaims[i];
        
        // Progress to Main UI
        self.postMessage({ type: 'ENGINE_PROGRESS', index: i+1, total: allClaims.length, asin: claim.asin });
        
        // Example execution
        const res = await engine.buildOneLevel(
           claim.barcode, claim.invoice, claim.sid, claim.po, claim.asin, 
           claim.inv_qty, claim.pqv, 0, true
        );
        
        self.postMessage({ type: 'ENGINE_BLOCK_RESULT', block: res.rows });
     }
     
     self.postMessage({ type: 'ENGINE_COMPLETE' });

  } else if (action === 'UI_USER_RESPONSE') {
     engine.resolveUserAction(payload);
  } else if (action === 'STOP') {
     engine.stop_requested = true;
  }
};
