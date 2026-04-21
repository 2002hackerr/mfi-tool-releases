import { useState, useEffect, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Upload, Play, Square, Settings, Database, FileSpreadsheet, ServerCrash, X, CheckCircle2, AlertCircle } from 'lucide-react';
import { clsx } from 'clsx';
import { twMerge } from 'tailwind-merge';

function cn(...inputs) {
  return twMerge(clsx(inputs));
}

// ── Components ─────────────────────────────────────────────

function AnimatedModal({ isOpen, onClose, title, children }) {
  return (
    <AnimatePresence>
      {isOpen && (
        <>
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.3 }}
            className="fixed inset-0 bg-black/60 backdrop-blur-sm z-40"
            onClick={onClose}
          />
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 20 }}
            transition={{ type: "spring", damping: 25, stiffness: 300 }}
            className="fixed left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 w-full max-w-lg z-50 p-1"
          >
            <div className="glass-panel p-6 rounded-2xl relative overflow-hidden">
              {/* Subtle top glow */}
              <div className="absolute top-0 left-1/4 right-1/4 h-px bg-gradient-to-r from-transparent via-accent-blue/50 to-transparent" />
              
              <div className="flex justify-between items-center mb-6">
                <h3 className="text-xl font-semibold tracking-tight text-white">{title}</h3>
                <button onClick={onClose} className="p-2 hover:bg-white/5 rounded-full transition-colors text-text-muted hover:text-white">
                  <X size={20} />
                </button>
              </div>
              <div className="space-y-4">
                {children}
              </div>
            </div>
          </motion.div>
        </>
      )}
    </AnimatePresence>
  );
}

function ProfessionalDropzone({ label, accept, icon: Icon, file, setFile }) {
  const [isDragActive, setIsDragActive] = useState(false);

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") setIsDragActive(true);
    else if (e.type === "dragleave") setIsDragActive(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragActive(false);
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFile(e.dataTransfer.files[0]);
    }
  };

  const handleChange = (e) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  return (
    <div className="flex flex-col gap-2">
      <label className="text-sm font-medium text-text-muted flex justify-between items-end">
        {label}
        {accept.includes('.csv') && accept.includes('.xlsx') && (
          <span className="text-[10px] font-bold uppercase tracking-wider text-accent-blue bg-accent-blue/10 px-2 py-0.5 rounded">.CSV & .XLSX</span>
        )}
      </label>
      
      <motion.div 
        whileHover={{ scale: 1.01 }}
        whileTap={{ scale: 0.99 }}
        className="relative group"
      >
        <input 
          type="file" 
          accept={accept} 
          onChange={handleChange} 
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
          onDragEnter={handleDrag} onDragLeave={handleDrag} onDragOver={handleDrag} onDrop={handleDrop}
        />
        
        <div className={cn(
          "relative w-full overflow-hidden rounded-xl border p-4 transition-all duration-300 flex items-center gap-4",
          isDragActive ? "border-accent-blue bg-accent-blue/5" : "border-white/10 bg-[#12131c]/50 hover:bg-[#12131c]",
          file && "border-accent-green/30 bg-accent-green/5"
        )}>
          {/* Active indicator bar */}
          {file && <motion.div layoutId={`active_${label}`} className="absolute left-0 top-0 bottom-0 w-1 bg-accent-green" />}
          
          <div className={cn(
            "p-3 rounded-lg flex-shrink-0 transition-colors",
            file ? "bg-accent-green/10 text-accent-green" : "bg-white/5 text-text-muted group-hover:text-white"
          )}>
            {file ? <CheckCircle2 size={24} /> : <Icon size={24} />}
          </div>
          
          <div className="flex flex-col min-w-0 pr-4">
            <span className={cn(
              "text-sm font-medium truncate transition-colors",
              file ? "text-white" : "text-text-muted group-hover:text-white"
            )}>
              {file ? file.name : "Click or drag file to upload"}
            </span>
            <span className="text-xs text-text-subtle truncate">
              {file ? `${(file.size / 1024 / 1024).toFixed(2)} MB` : `Supports ${accept.replace(/,/g, ', ')}`}
            </span>
          </div>
        </div>
      </motion.div>
    </div>
  );
}

// ── Application ─────────────────────────────────────────────

export default function App() {
  const [claimsFile, setClaimsFile] = useState(null);
  const [rebniFile, setRebniFile] = useState(null);
  const [invFile, setInvFile] = useState(null);
  
  // Investigation State & Worker
  const [isIngesting, setIsIngesting] = useState(false);
  const [ingestStatus, setIngestStatus] = useState({});
  const workerRef = useRef(null);

  useEffect(() => {
    // Initialize Web Worker
    workerRef.current = new Worker(new URL('./workers/ingestion.worker.js', import.meta.url), {
      type: 'module'
    });

    workerRef.current.onmessage = (e) => {
      const { type, table, status, count, error } = e.data;
      
      if (type === 'DB_CLEARED') {
        // Start parsing files once DB is cleared
        if (claimsFile) workerRef.current.postMessage({ action: 'PARSE_FILE', payload: { file: claimsFile, table: 'claims' } });
        if (rebniFile) workerRef.current.postMessage({ action: 'PARSE_FILE', payload: { file: rebniFile, table: 'rebni' } });
        if (invFile) workerRef.current.postMessage({ action: 'PARSE_FILE', payload: { file: invFile, table: 'invoice_search' } });
      } else if (type === 'PROGRESS') {
        setIngestStatus(prev => ({ ...prev, [table]: { loading: true, msg: status } }));
      } else if (type === 'COMPLETE') {
        setIngestStatus(prev => ({ ...prev, [table]: { loading: false, msg: `Loaded ${count} rows successfully.`, done: true } }));
      } else if (type === 'ERROR') {
        setIngestStatus(prev => ({ ...prev, [table || 'global']: { loading: false, msg: `Error: ${error}`, error: true } }));
      }
    };

    return () => workerRef.current?.terminate();
  }, [claimsFile, rebniFile, invFile]);

  const handleRunInvestigation = () => {
    if (!claimsFile || !rebniFile || !invFile) {
      alert("Please upload all 3 files first.");
      return;
    }
    setIsIngesting(true);
    setIngestStatus({
      claims: { loading: true, msg: 'Waiting...' },
      rebni: { loading: true, msg: 'Waiting...' },
      invoice_search: { loading: true, msg: 'Waiting...' }
    });
    // Kick off process by wiping old DB
    workerRef.current.postMessage({ action: 'INIT_DB' });
  };
  
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [mode, setMode] = useState('manual');
  const [ticketType, setTicketType] = useState('PDTT');
  const [ticketId, setTicketId] = useState('');

  return (
    <div className="min-h-screen pt-12 pb-24 px-6 md:px-12 flex flex-col gap-10 max-w-7xl mx-auto selection:bg-accent-red/30">
      
      {/* Header */}
      <motion.header 
        initial={{ y: -20, opacity: 0 }}
        animate={{ y: 0, opacity: 1 }}
        className="flex flex-col md:flex-row justify-between items-start md:items-end border-b border-white/10 pb-6 gap-4"
      >
        <div>
          <div className="flex items-center gap-3 mb-2">
            <div className="p-2.5 bg-accent-red/10 rounded-xl">
              <ServerCrash size={28} className="text-accent-red" />
            </div>
            <h1 className="text-3xl md:text-4xl font-bold tracking-tight text-white">
              MFI Investigation Framework
            </h1>
          </div>
          <p className="text-text-muted text-sm ml-[3.25rem] flex items-center gap-3">
            Local-First Browser Engine 
            <span className="hidden md:inline-block w-1.5 h-1.5 rounded-full bg-accent-blue animate-pulse"/>
            <span className="text-white/60">Ready exclusively for the browser</span>
          </p>
        </div>
        <div className="flex flex-col items-end gap-3">
          <span className="text-sm italic font-medium bg-gradient-to-r from-accent-blue to-accent-green bg-clip-text text-transparent">Developed by Mukesh</span>
          <button 
            onClick={() => setIsSettingsOpen(true)}
            className="flex items-center gap-2 px-4 py-2 rounded-lg bg-white/5 hover:bg-white/10 transition-colors border border-white/5 text-sm font-medium"
          >
            <Settings size={18} className="text-text-muted" /> Config
          </button>
        </div>
      </motion.header>

      {/* Main Grid */}
      <main className="grid grid-cols-1 lg:grid-cols-12 gap-8">
        
        {/* Left Column: Data Sources */}
        <motion.div 
          initial={{ x: -20, opacity: 0 }}
          animate={{ x: 0, opacity: 1 }}
          transition={{ delay: 0.1 }}
          className="col-span-1 lg:col-span-8 flex flex-col gap-6"
        >
          <section className="glass-panel p-8 rounded-2xl relative overflow-hidden">
            {/* Ambient background glow */}
            <div className="absolute top-0 right-0 -mr-20 -mt-20 w-64 h-64 bg-accent-blue/5 rounded-full blur-3xl pointer-events-none" />
            
            <div className="flex items-center justify-between mb-8">
              <h2 className="text-xl font-semibold tracking-tight text-white flex items-center gap-3">
                <Database size={22} className="text-accent-blue"/> Data Ingestion
              </h2>
            </div>
            
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-6">
              <ProfessionalDropzone 
                label="Claims Sheet" 
                accept=".csv,.xlsx" 
                icon={FileSpreadsheet} 
                file={claimsFile} 
                setFile={setClaimsFile} 
              />
              <ProfessionalDropzone 
                label="REBNI Result" 
                accept=".csv,.xlsx" 
                icon={FileSpreadsheet} 
                file={rebniFile} 
                setFile={setRebniFile} 
              />
            </div>
            
            <div className="border-t border-white/5 pt-6">
              <ProfessionalDropzone 
                label="Invoice Search (Primary Volume Data)" 
                accept=".csv,.xlsx" 
                icon={Database} 
                file={invFile} 
                setFile={setInvFile} 
              />
            </div>
          </section>

          {/* Ticket ID Quick Entry */}
          <section className="glass-panel p-6 rounded-xl flex items-center justify-between gap-6">
            <span className="text-text-muted font-medium whitespace-nowrap">Session Ticket ID:</span>
            <input 
              type="text" 
              value={ticketId} 
              onChange={(e)=>setTicketId(e.target.value)} 
              placeholder="e.g. TICK-90210" 
              className="w-full bg-[#0a0a0f] border border-white/10 rounded-lg px-4 py-2.5 text-white focus:outline-none focus:border-accent-blue/50 focus:ring-1 focus:ring-accent-blue/50 transition-all placeholder:text-white/20 font-mono text-sm"
            />
          </section>
        </motion.div>

        {/* Right Column: Execution */}
        <motion.div 
          initial={{ x: 20, opacity: 0 }}
          animate={{ x: 0, opacity: 1 }}
          transition={{ delay: 0.2 }}
          className="col-span-1 lg:col-span-4 flex flex-col gap-6"
        >
          {/* Action Panel */}
          <section className="glass-panel p-6 rounded-2xl flex flex-col relative overflow-hidden h-full min-h-[300px]">
            <div className="absolute bottom-0 left-0 right-0 h-32 bg-gradient-to-t from-accent-red/5 to-transparent pointer-events-none" />
            
            <div className="flex-1 flex flex-col items-center justify-center text-center gap-6 z-10 px-4">
              <motion.div 
                animate={{ y: [0, -5, 0] }}
                transition={{ duration: 4, repeat: Infinity, ease: "easeInOut" }}
                className="w-16 h-16 rounded-2xl bg-gradient-to-br from-[#1f202f] to-[#0a0a0f] border border-white/10 shadow-xl flex items-center justify-center"
              >
                <Play fill="currentColor" className="text-accent-red w-8 h-8 ml-1" />
              </motion.div>
              
              <div>
                <h3 className="text-lg font-bold text-white mb-2">Execute Engine</h3>
                <p className="text-sm text-text-muted balance">
                  File structures are parsed seamlessly using Web Workers. Memory is safely delegated to IndexedDB.
                </p>
              </div>
            </div>

            <div className="flex flex-col gap-3 mt-auto relative z-10">
              <button 
                onClick={handleRunInvestigation}
                disabled={isIngesting}
                className={cn(
                  "glow-effect w-full text-white font-bold py-3.5 rounded-xl uppercase tracking-widest text-sm shadow-[0_0_20px_rgba(233,69,96,0.2)] transition-all",
                  isIngesting ? "bg-accent-red/50 cursor-not-allowed hidden" : "bg-accent-red hover:shadow-[0_0_25px_rgba(233,69,96,0.4)]"
                )}
              >
                Run Investigation
              </button>
              
              {isIngesting && (
                <div className="w-full bg-black/40 rounded-xl p-4 border border-white/10 text-left text-xs mb-3 space-y-3">
                  <div className="font-semibold text-white mb-2 uppercase tracking-wide">Data Ingestion Engine Active</div>
                  
                  {['claims', 'rebni', 'invoice_search'].map(table => (
                    <div key={table} className="flex flex-col gap-1">
                      <div className="flex justify-between text-text-muted">
                        <span className="capitalize">{table.replace('_', ' ')}</span>
                        {ingestStatus[table]?.loading && <span className="animate-pulse text-accent-blue">Working...</span>}
                        {ingestStatus[table]?.done && <span className="text-accent-green">Complete</span>}
                        {ingestStatus[table]?.error && <span className="text-accent-red">Failed</span>}
                      </div>
                      <div className="text-white/80">{ingestStatus[table]?.msg || 'Pending'}</div>
                    </div>
                  ))}
                  
                  {/* Button to artificially hide loader for debug right now */}
                  <button onClick={() => setIsIngesting(false)} className="mt-4 w-full text-text-subtle hover:text-white underline text-center">Hide Overlay / Cancel</button>
                </div>
              )}

              <button className="w-full bg-white/5 hover:bg-white/10 text-text-muted hover:text-white border border-white/5 transition-colors font-semibold py-3 rounded-xl flex items-center justify-center gap-2 text-sm">
                <Square size={16} fill="currentColor"/> Abort Action
              </button>
            </div>
          </section>

        </motion.div>

      </main>

      {/* Settings Modal overlay */}
      <AnimatedModal isOpen={isSettingsOpen} onClose={() => setIsSettingsOpen(false)} title="Engine Configuration">
        
        <div className="space-y-6">
          <div className="p-4 rounded-xl bg-white/5 border border-white/10">
            <h4 className="text-sm font-semibold text-white mb-4">Investigation Strategy</h4>
            <div className="flex flex-col sm:flex-row gap-3">
              <label className={cn(
                "flex-1 p-3 rounded-lg border cursor-pointer transition-all",
                mode === 'auto' ? "bg-accent-blue/10 border-accent-blue text-accent-blue" : "bg-black/20 border-white/5 text-text-muted hover:bg-white/5"
              )}>
                <input type="radio" className="hidden" checked={mode === 'auto'} onChange={() => setMode('auto')} />
                <div className="font-semibold text-sm mb-1">Automatic</div>
                <div className="text-xs opacity-70">Silent execution with prompt falls.</div>
              </label>
              
              <label className={cn(
                "flex-1 p-3 rounded-lg border cursor-pointer transition-all",
                mode === 'manual' ? "bg-accent-blue/10 border-accent-blue text-accent-blue" : "bg-black/20 border-white/5 text-text-muted hover:bg-white/5"
              )}>
                <input type="radio" className="hidden" checked={mode === 'manual'} onChange={() => setMode('manual')} />
                <div className="font-semibold text-sm mb-1">Manual (Interactive)</div>
                <div className="text-xs opacity-70">Step-by-step depth alignment.</div>
              </label>
            </div>
          </div>

          <div className="p-4 rounded-xl bg-white/5 border border-white/10">
            <h4 className="text-sm font-semibold text-white mb-4">Ticket Type Scope</h4>
            <div className="flex flex-col gap-3">
              <label className="flex items-center gap-3 p-2 cursor-pointer group">
                <div className={cn(
                  "w-5 h-5 rounded-full border flex items-center justify-center bg-[#0a0a0f] transition-colors",
                  ticketType === 'PDTT' ? "border-accent-green" : "border-white/20 group-hover:border-white/40"
                )}>
                  {ticketType === 'PDTT' && <div className="w-2.5 h-2.5 rounded-full bg-accent-green" />}
                </div>
                <input type="radio" className="hidden" checked={ticketType === 'PDTT'} onChange={() => setTicketType('PDTT')} />
                <div>
                  <div className="font-semibold text-sm text-white">PDTT <span className="text-text-subtle font-normal ml-2">(Default)</span></div>
                  <div className="text-xs text-text-muted">Full-chain sub-level shortage inheritance.</div>
                </div>
              </label>

              <label className="flex items-center gap-3 p-2 cursor-pointer group">
                <div className={cn(
                  "w-5 h-5 rounded-full border flex items-center justify-center bg-[#0a0a0f] transition-colors",
                  ticketType === 'REMASH' ? "border-[#f0c060]" : "border-white/20 group-hover:border-white/40"
                )}>
                  {ticketType === 'REMASH' && <div className="w-2.5 h-2.5 rounded-full bg-[#f0c060]" />}
                </div>
                <input type="radio" className="hidden" checked={ticketType === 'REMASH'} onChange={() => setTicketType('REMASH')} />
                <div>
                  <div className="font-semibold text-sm text-white">REMASH TT</div>
                  <div className="text-xs text-[#f0c060]/70">Halt investigation at root claiming shipment depth (`depth=0`).</div>
                </div>
              </label>
            </div>
          </div>
        </div>

      </AnimatedModal>

    </div>
  );
}
