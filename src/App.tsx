import React, { useState, useEffect, useMemo, useRef, Suspense } from 'react';
import { 
  FileText, 
  Settings, 
  Download, 
  PlusCircle, 
  Trash2, 
  Info, 
  ArrowRight,
  CheckCircle2,
  Copy,
  FileType,
  Upload,
  Loader2,
  Book,
  Save,
  Search,
  ShieldCheck,
  Lock
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { Document, Packer, Paragraph, TextRun, AlignmentType, Header, ImageRun, TextWrappingType, TextWrappingSide, HorizontalPositionRelativeFrom, VerticalPositionRelativeFrom } from 'docx';
import { saveAs } from 'file-saver';
import mammoth from 'mammoth';
import DOMPurify from 'dompurify';


import 'react-quill-new/dist/quill.snow.css';

// Modular Imports
import { Variable, SavedTemplate, AppStep } from './types';
import { extractVariables, fillTemplate } from './utils/documentUtils';
import { analyzeDocument } from './services/aiService';
import { DEFAULT_TEMPLATES, INITIAL_TEMPLATE } from './constants/templates';

const ReactQuill = React.lazy(() => import('react-quill-new'));

// --- SVG Logo Component ---
const SimplesLogo = ({ className = "" }: { className?: string }) => {
  const [error, setError] = useState(false);
  if (error) {
    return <span className={`font-cinzel font-bold text-inherit block text-center ${className}`}>SIMPLES</span>;
  }
  return <img src="logo.png" alt="Simples Logo" className={`object-contain ${className}`} onError={() => setError(true)} />;
};

export default function App() {
  const [template, setTemplate] = useState<string>(INITIAL_TEMPLATE);
  const [variables, setVariables] = useState<Variable[]>([]);
  const [activeStep, setActiveStep] = useState<AppStep>('editor');
  const [copySuccess, setCopySuccess] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [savedTemplates, setSavedTemplates] = useState<SavedTemplate[]>([]);
  const [templateName, setTemplateName] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [showSettings, setShowSettings] = useState(false);
  const [apiKey, setApiKey] = useState(localStorage.getItem('gemini_api_key') || "");
  const [useBackendKey, setUseBackendKey] = useState(true);
  const [updateStatus, setUpdateStatus] = useState<'idle' | 'available' | 'downloading' | 'downloaded' | 'error'>('idle');
  const [updateMessage, setUpdateMessage] = useState("");

  
  const fileInputRef = useRef<HTMLInputElement>(null);
  const printRef = useRef<HTMLDivElement>(null);

  // Quill Modules Configuration
  const quillModules = useMemo(() => ({
    toolbar: [
      [{ header: [1, 2, 3, false] }],
      ['bold', 'italic', 'underline', 'strike'],
      [{ list: 'ordered' }, { list: 'bullet' }],
      [{ align: [] }],
      ['clean'],
    ],
  }), []);

  const filteredTemplates = useMemo(() => {
    return savedTemplates.filter(t => 
      t.name.toLowerCase().includes(searchQuery.toLowerCase())
    );
  }, [savedTemplates, searchQuery]);

  // Load templates and settings from LocalStorage
  useEffect(() => {
    const stored = localStorage.getItem('auto_termos_templates');
    if (stored) {
      setSavedTemplates(JSON.parse(stored));
    }
  }, []);

  // Save API Key
  const saveSettings = () => {
    localStorage.setItem('gemini_api_key', apiKey);
    setShowSettings(false);
  };

  // Electron Update Listeners
  useEffect(() => {
    const api = (window as any).electronAPI;
    if (!api) return;

    api.onUpdateAvailable(() => {
      setUpdateStatus('available');
      setUpdateMessage("Nova atualização disponível! Baixando...");
    });

    api.onUpdateDownloaded(() => {
      setUpdateStatus('downloaded');
      setUpdateMessage("Atualização baixada com sucesso!");
    });

    api.onUpdateError((msg: string) => {
      setUpdateStatus('error');
      setUpdateMessage(`Erro na atualização: ${msg}`);
    });
  }, []);

  const handleCheckUpdates = () => {
    const api = (window as any).electronAPI;
    const isProbablyElectron = navigator.userAgent.toLowerCase().includes('electron');

    if (api) {
      setUpdateStatus('downloading');
      setUpdateMessage("Verificando atualizações...");
      api.checkForUpdates();
    } else if (isProbablyElectron) {
      alert("Sistema desktop detectado, mas a interface de atualização não respondeu. Tente reiniciar o aplicativo.");
    } else {
      alert("Recurso disponível apenas na versão instalada (Desktop).");
    }
  };

  const handleRestartApp = () => {
    const api = (window as any).electronAPI;
    if (api) api.restartApp();
  };

  // Save templates to LocalStorage
  const saveToLibrary = () => {
    const name = templateName.trim() || `Modelo ${savedTemplates.length + 1}`;
    const newTemplate: SavedTemplate = {
      id: crypto.randomUUID(),
      name,
      content: template,
      lastUsed: Date.now()
    };
    const updated = [newTemplate, ...savedTemplates];
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
    setTemplateName("");
  };

  const deleteFromLibrary = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    const updated = savedTemplates.filter(t => t.id !== id);
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
  };

  const selectTemplate = (t: SavedTemplate) => {
    setTemplate(t.content);
    const updated = savedTemplates.map(s => s.id === t.id ? { ...s, lastUsed: Date.now() } : s);
    setSavedTemplates(updated);
    localStorage.setItem('auto_termos_templates', JSON.stringify(updated));
  };

  // Sync variables with template
  useEffect(() => {
    const keys = extractVariables(template);
    setVariables(prev => {
      return keys.map(key => {
        const existing = prev.find(p => p.key === key);
        return { key, value: existing ? existing.value : '' };
      });
    });
  }, [template]);

  const filledText = useMemo(() => fillTemplate(template, variables), [template, variables]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!useBackendKey && !apiKey) {
      alert("Por favor, configure sua Chave de API Gemini nas configurações ou utilize a chave do servidor.");
      setShowSettings(true);
      return;
    }


    setIsAnalyzing(true);
    try {
      const reader = new FileReader();
      if (file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        reader.onload = async (event) => {
          const arrayBuffer = event.target?.result as ArrayBuffer;
          const result = await mammoth.extractRawText({ arrayBuffer });
          const processed = await analyzeDocument("", "text/plain", result.value);
          if (processed) setTemplate(processed);
          setIsAnalyzing(false);
        };
        reader.readAsArrayBuffer(file);
      } else {
        alert("Por favor, envie um arquivo .docx");
        setIsAnalyzing(false);
      }
    } catch (err) {
      console.error(err);
      alert("Erro ao processar arquivo com IA. Verifique sua chave de API e conexão.");
      setIsAnalyzing(false);
    }
  };

  const handleDownloadDocx = async () => {
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = filledText;
    
    const docChildren: Paragraph[] = [];
    const elements = tempDiv.querySelectorAll('p, h1, h2, h3, h4, h5, h6, li');
    
    if (elements.length > 0) {
      elements.forEach(el => {
        const isCenter = (el as HTMLElement).style?.textAlign === 'center' || el.classList?.contains('ql-align-center');
        const isRight = (el as HTMLElement).style?.textAlign === 'right' || el.classList?.contains('ql-align-right');
        const isJustify = (el as HTMLElement).style?.textAlign === 'justify' || el.classList?.contains('ql-align-justify');
        
        const alignment = isCenter ? AlignmentType.CENTER : 
                          isRight ? AlignmentType.RIGHT : 
                          isJustify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT;
        
        const textOptions: any = {
          text: el.textContent || '',
          size: el.nodeName === 'H1' ? 32 : el.nodeName === 'H2' ? 28 : el.nodeName === 'H3' ? 24 : 22,
          font: 'Cormorant Garamond'
        };

        if (el.nodeName.startsWith('H') || el.innerHTML?.includes('<strong>') || el.innerHTML?.includes('<b>')) {
          textOptions.bold = true;
        }
        if (el.innerHTML?.includes('<em>') || el.innerHTML?.includes('<i>')) {
          textOptions.italics = true;
        }

        const paragraphOptions: any = {
          children: [new TextRun(textOptions)],
          spacing: { after: 200 },
          alignment
        };

        if (el.nodeName === 'LI') {
          paragraphOptions.bullet = { level: 0 };
        }

        docChildren.push(new Paragraph(paragraphOptions));
      });
    } else {
      docChildren.push(new Paragraph({ 
        children: [new TextRun({ text: filledText.replace(/<[^>]*>?/gm, ''), font: 'Cormorant Garamond', size: 22 })] 
      }));
    }

    let headers: any = undefined;
    let bgFileBytes: Uint8Array | null = null;
    try {
      const resp = await fetch('a.png');
      if (resp.ok) {
        const arrayBuffer = await resp.arrayBuffer();
        bgFileBytes = new Uint8Array(arrayBuffer);
      }
    } catch (e) {
      console.warn("Could not load a.png");
    }

    if (bgFileBytes) {
      headers = {
        default: new Header({
          children: [
            new Paragraph({
              children: [
                new ImageRun({
                  data: bgFileBytes,
                  transformation: { width: 794, height: 1123 },
                  floating: {
                    horizontalPosition: { relative: HorizontalPositionRelativeFrom.PAGE, offset: 0 },
                    verticalPosition: { relative: VerticalPositionRelativeFrom.PAGE, offset: 0 },
                    behindDocument: true,
                    wrap: { type: TextWrappingType.NONE, side: TextWrappingSide.BOTH_SIDES },
                  },
                }),
              ],
            }),
          ],
        })
      };
    }

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            margin: { top: 4500, bottom: 2800, left: 1400, right: 1400 },
          },
        },
        headers,
        children: docChildren,
      }],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `Termo_Preenchido_${new Date().getTime()}.docx`);
  };

  const handleCopyClipboard = () => {
    const plainText = filledText.replace(/<[^>]*>?/gm, '\n').replace(/\n\n+/g, '\n\n');
    navigator.clipboard.writeText(plainText);
    setCopySuccess(true);
    setTimeout(() => setCopySuccess(false), 2000);
  };

  return (
    <div className="min-h-screen font-sans selection:bg-bg-bordo text-text-creme">
      {/* Header */}
      <header className="border-b border-bg-bordo bg-bg-ebano/95 backdrop-blur-md sticky top-0 z-20">
        <div className="max-w-6xl mx-auto px-4 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="h-8 flex items-center justify-center text-ouro-imperial">
               <SimplesLogo className="h-full w-auto object-contain" />
            </div>
            <h1 className="text-xl font-cinzel font-bold tracking-tight text-ouro-imperial uppercase mt-1">
              <span className="text-text-terra font-normal text-sm md:text-lg hidden md:inline ml-1">- AutoTermos v1</span>
            </h1>
          </div>
          
          <nav className="hidden md:flex items-center gap-8">
            {(['editor', 'filler', 'preview'] as AppStep[]).map((step, idx) => (
              <button 
                key={step}
                onClick={() => setActiveStep(step)}
                className={`text-sm font-cinzel font-bold tracking-widest uppercase transition-colors ${activeStep === step ? 'text-ouro-imperial' : 'text-text-terra hover:text-text-creme'}`}
              >
                {idx + 1}. {step === 'editor' ? 'Modelo' : step === 'filler' ? 'Preencher' : 'Concluído'}
              </button>
            ))}
          </nav>

          <div className="flex items-center gap-2">
            {/* Settings gear removed as per user request */}
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8 md:py-12">
        <AnimatePresence mode="wait">
          {/* STEP 1: EDITOR */}
          {activeStep === 'editor' && (
            <motion.div 
              key="editor"
              initial={{ opacity: 0, y: 10 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -10 }}
              className="grid grid-cols-1 lg:grid-cols-12 gap-8"
            >
              <div className="lg:col-span-3 space-y-6">
                <div className="bg-bg-noite rounded-xl border border-bg-bordo shadow-inner overflow-hidden min-h-[500px] flex flex-col">
                  <div className="p-4 border-b border-bg-bordo space-y-3 bg-bg-bordo-profundo/30">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2 text-ouro-imperial font-cinzel font-bold text-sm uppercase tracking-wider">
                        <Book size={16} className="text-ouro-imperial" />
                        Biblioteca
                      </div>
                    </div>
                    <div className="relative">
                      <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-text-terra" size={14} />
                      <input 
                        type="text"
                        placeholder="Buscar modelo..."
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        className="w-full pl-9 pr-3 py-2 bg-bg-ebano border border-bg-bordo rounded-lg text-xs text-text-creme outline-none focus:ring-1 focus:ring-ouro-imperial transition-all font-medium shadow-sm placeholder:text-text-terra"
                      />
                    </div>
                  </div>
                  
                  <div className="flex-1 overflow-y-auto p-2 space-y-4 custom-scrollbar">
                    {!searchQuery && (
                      <div className="space-y-1">
                        <div className="px-3 py-1 text-[10px] font-cinzel font-bold text-text-areia uppercase tracking-widest">Sugeridos</div>
                        {DEFAULT_TEMPLATES.map(t => (
                          <button
                            key={t.id}
                            onClick={() => selectTemplate(t)}
                            className={`w-full text-left p-3 rounded-lg hover:bg-bg-bordo-profundo transition-colors flex flex-col border border-transparent ${template === t.content ? 'bg-bg-bordo border-ouro-escuro/30' : ''}`}
                          >
                            <div className="text-sm font-semibold text-text-creme truncate">{t.name}</div>
                            <div className="text-[10px] text-ouro-escuro font-bold uppercase tracking-widest mt-0.5">Padrão</div>
                          </button>
                        ))}
                      </div>
                    )}

                    <div className="space-y-1">
                      <div className="px-3 py-1 text-[10px] font-cinzel font-bold text-text-areia uppercase tracking-widest">
                        {searchQuery ? "Resultados" : "Seus Modelos"}
                      </div>
                      {filteredTemplates.length === 0 && !searchQuery ? (
                        <div className="p-4 text-center text-text-terra text-[10px] italic">Nenhum modelo personalizado.</div>
                      ) : (
                        filteredTemplates.map(t => (
                          <div key={t.id} className="group relative">
                            <button
                              onClick={() => selectTemplate(t)}
                              className={`w-full text-left p-3 pr-10 rounded-lg hover:bg-bg-bordo-profundo transition-colors flex flex-col border border-transparent ${template === t.content ? 'bg-bg-bordo border-ouro-escuro/30' : ''}`}
                            >
                              <div className="text-sm font-semibold text-text-creme truncate pr-2">{t.name}</div>
                              <div className="text-[10px] text-text-terra uppercase tracking-widest mt-0.5">Criado em {new Date(t.lastUsed).toLocaleDateString()}</div>
                            </button>
                            <button
                              onClick={(e) => deleteFromLibrary(t.id, e)}
                              className="absolute right-2 top-1/2 -translate-y-1/2 p-2 text-gray-300 hover:text-red-500 rounded-md transition-all active:scale-95"
                            >
                              <Trash2 size={14} />
                            </button>
                          </div>
                        ))
                      )}
                    </div>
                  </div>
                </div>

                <div className="bg-bg-bordo-profundo border border-bg-bordo p-6 rounded-xl space-y-4 relative overflow-hidden">
                  <div className="absolute top-0 right-0 w-32 h-32 bg-ouro-imperial/5 rounded-full blur-3xl" />
                  <h4 className="text-xs font-cinzel font-bold text-ouro-claro uppercase tracking-widest relative z-10">Importar com IA</h4>
                  <p className="text-xs text-text-creme/80 leading-relaxed relative z-10">Suba um .docx e nossa IA identificará as variáveis.</p>
                  <input type="file" ref={fileInputRef} onChange={handleFileUpload} accept=".docx" className="hidden" />
                  <button 
                    onClick={() => fileInputRef.current?.click()}
                    disabled={isAnalyzing}
                    className="w-full py-2.5 px-4 bg-bg-ebano border border-ouro-escuro/50 text-ouro-claro rounded-lg text-sm font-semibold hover:bg-bg-bordo transition-all flex items-center justify-center gap-2 disabled:opacity-50 relative z-10"
                  >
                    {isAnalyzing ? <Loader2 size={16} className="animate-spin" /> : <Upload size={16} />}
                    {isAnalyzing ? 'Analisando...' : 'Subir Arquivo'}
                  </button>
                </div>

                {/* System Info & Updates Section */}
                <div className="bg-bg-ebano/40 border border-bg-bordo/50 p-4 rounded-xl space-y-3">
                  <div className="flex items-center justify-between">
                    <div className="flex flex-col">
                      <span className="text-[10px] font-cinzel font-bold text-text-terra uppercase tracking-widest">Versão</span>
                      <span className="text-xs text-ouro-claro font-mono">v1.0.0</span>
                    </div>
                    {updateStatus === 'downloaded' ? (
                      <button 
                        onClick={handleRestartApp}
                        className="px-3 py-1.5 bg-green-600/20 text-green-400 border border-green-600/50 rounded-lg text-[10px] font-bold uppercase hover:bg-green-600/40 transition-all"
                      >
                        Reiniciar
                      </button>
                    ) : (
                      <button 
                        onClick={handleCheckUpdates}
                        disabled={updateStatus === 'downloading'}
                        className="px-3 py-1.5 bg-bg-bordo/50 border border-ouro-escuro/30 text-ouro-claro rounded-lg text-[10px] font-bold uppercase hover:bg-bg-bordo transition-all disabled:opacity-50"
                      >
                        {updateStatus === 'downloading' ? '...' : 'Atualizar'}
                      </button>
                    )}
                  </div>
                  {updateMessage && (
                    <p className={`text-[9px] text-center ${updateStatus === 'error' ? 'text-red-400' : 'text-ouro-imperial'} font-medium`}>
                      {updateMessage}
                    </p>
                  )}
                </div>
              </div>

              <div className="lg:col-span-6 space-y-4">
                <div className="flex items-center justify-between">
                  <input 
                    type="text" 
                    value={templateName}
                    onChange={(e) => setTemplateName(e.target.value)}
                    placeholder="Nome do Modelo (Opcional)"
                    className="flex-1 max-w-xs bg-transparent border-b border-bg-bordo focus:border-ouro-imperial outline-none px-1 py-1 text-sm font-medium text-text-creme"
                  />
                  <button 
                    onClick={saveToLibrary}
                    className="flex items-center gap-2 text-xs font-cinzel font-bold text-ouro-imperial hover:bg-bg-bordo px-3 py-1.5 rounded-lg transition-colors uppercase tracking-wider"
                  >
                    <Save size={14} /> Salvar Modelo
                  </button>
                </div>
                
                <div className="relative group shadow-sm bg-champagne rounded-xl border border-ouro-escuro/30 overflow-hidden text-black">
                  <Suspense fallback={<div className="w-full h-[550px] flex items-center justify-center bg-bg-noite"><Loader2 className="animate-spin text-ouro-imperial" size={32} /></div>}>
                    <ReactQuill theme="snow" value={template} onChange={setTemplate} modules={quillModules} placeholder="Ex: [NOME] compareceu..." className="h-[550px] mb-12" />
                  </Suspense>
                  {isAnalyzing && (
                    <div className="absolute inset-0 bg-bg-ebano/80 backdrop-blur-sm flex flex-col items-center justify-center gap-4 z-10 text-ouro-imperial text-center p-8">
                      <Loader2 size={40} className="animate-spin" />
                      <span className="font-cinzel font-bold text-sm tracking-widest uppercase mt-4">IA Analisando Documento</span>
                    </div>
                  )}
                </div>
              </div>

              <div className="lg:col-span-3 space-y-6">
                <div className="bg-bg-noite p-6 rounded-xl border border-bg-bordo shadow-inner space-y-4 sticky top-24 relative overflow-hidden">
                  <div className="flex items-center gap-2 text-ouro-claro relative z-10">
                    <Info size={18} />
                    <h3 className="font-cinzel font-bold uppercase tracking-widest text-xs">Instruções</h3>
                  </div>
                  <ul className="space-y-3 text-sm text-text-creme/80 relative z-10">
                    <li className="flex gap-2"><span className="text-ouro-imperial font-bold">•</span><span>Use <b>[colchetes]</b> para variáveis.</span></li>
                    <li className="flex gap-2"><span className="text-ouro-imperial font-bold">•</span><span>Variáveis repetidas serão preenchidas juntas.</span></li>
                  </ul>
                  <div className="pt-4 space-y-2 relative z-10">
                    <h4 className="text-[10px] font-cinzel font-bold text-text-terra uppercase tracking-widest">Resumo</h4>
                    <div className="flex items-center justify-between text-sm">
                      <span className="text-text-areia">Variáveis</span>
                      <span className="font-bold text-bg-ebano bg-ouro-claro px-2 py-0.5 rounded">{variables.length}</span>
                    </div>
                  </div>
                  <button 
                    onClick={() => setActiveStep('filler')}
                    disabled={variables.length === 0}
                    className="w-full mt-6 py-4 px-4 bg-gradient-to-r from-ouro-escuro to-ouro-imperial hover:from-ouro-imperial hover:to-ouro-claro disabled:opacity-50 text-bg-ebano rounded-xl font-cinzel font-bold text-sm tracking-wider uppercase transition-all flex items-center justify-center gap-3"
                  >
                    Começar Preenchimento <ArrowRight size={18} />
                  </button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 2: FILLER */}
          {activeStep === 'filler' && (
            <motion.div key="filler" initial={{ opacity: 0, x: 20 }} animate={{ opacity: 1, x: 0 }} exit={{ opacity: 0, x: -20 }} className="max-w-3xl mx-auto space-y-8">
              <div className="text-center space-y-2">
                <h2 className="text-3xl font-cinzel font-bold text-ouro-imperial uppercase tracking-wide">Dados do Documento</h2>
                <p className="text-text-areia font-serif italic text-lg">Preencha as informações abaixo.</p>
              </div>
              <div className="bg-bg-noite border-t-2 border-t-ouro-imperial border-x border-b border-bg-bordo rounded-2xl shadow-2xl overflow-hidden relative">
                <div className="p-8 space-y-6 max-h-[600px] overflow-y-auto custom-scrollbar">
                  {variables.map((variable, idx) => (
                    <div key={variable.key} className="space-y-1.5 group">
                      <label className="block text-xs font-cinzel font-bold text-text-terra uppercase tracking-widest group-focus-within:text-ouro-claro transition-colors">{variable.key}</label>
                      <input
                        type="text"
                        value={variable.value}
                        onChange={(e) => {
                          const newVars = [...variables];
                          newVars[idx].value = e.target.value;
                          setVariables(newVars);
                        }}
                        placeholder={`Digite o ${variable.key.toLowerCase()}...`}
                        className="w-full px-4 py-4 bg-bg-ebano border border-bg-bordo rounded-xl focus:ring-1 focus:ring-ouro-imperial outline-none transition-all text-text-creme"
                      />
                    </div>
                  ))}
                </div>
                <div className="bg-bg-bordo-profundo/50 p-8 flex flex-col sm:flex-row gap-4 border-t border-bg-bordo">
                  <button onClick={() => setActiveStep('editor')} className="flex-1 py-4 px-6 border border-ouro-escuro/50 text-text-creme rounded-xl font-cinzel uppercase font-bold hover:bg-ouro-imperial/10">Editar Modelo</button>
                  <button onClick={() => setActiveStep('preview')} className="flex-[3] py-4 px-6 bg-gradient-to-r from-ouro-escuro to-ouro-imperial text-bg-ebano rounded-xl font-cinzel uppercase font-bold hover:to-ouro-claro transition-all flex items-center justify-center gap-3">Gerar Documento <CheckCircle2 size={20} /></button>
                </div>
              </div>
            </motion.div>
          )}

          {/* STEP 3: PREVIEW */}
          {activeStep === 'preview' && (
            <motion.div key="preview" initial={{ opacity: 0, scale: 0.98 }} animate={{ opacity: 1, scale: 1 }} exit={{ opacity: 0, scale: 0.98 }} className="grid grid-cols-1 lg:grid-cols-4 gap-8">
              <div className="lg:col-span-1 space-y-4 order-2 lg:order-1">
                <h3 className="text-sm font-cinzel font-bold text-ouro-imperial uppercase tracking-widest">Exportar</h3>
                <div className="space-y-3">
                  <button onClick={handleDownloadDocx} className="w-full flex items-center gap-4 p-5 bg-bg-noite border border-bg-bordo rounded-2xl hover:border-ouro-imperial transition-all text-left group">
                    <div className="p-3 bg-bg-bordo-profundo rounded-xl group-hover:bg-bg-bordo transition-colors"><FileType className="text-ouro-imperial" size={24} /></div>
                    <div><div className="font-cinzel font-bold text-base text-text-creme">Baixar Word</div><div className="text-[10px] text-text-terra font-bold uppercase">Editável</div></div>
                  </button>
                  <button onClick={handleCopyClipboard} className="w-full flex items-center gap-4 p-5 bg-bg-noite border border-bg-bordo rounded-2xl hover:border-ouro-imperial transition-all text-left group">
                    <div className="p-3 bg-bg-bordo-profundo rounded-xl group-hover:bg-bg-bordo transition-colors">{copySuccess ? <CheckCircle2 className="text-green-500" size={24} /> : <Copy className="text-text-areia" size={24} />}</div>
                    <div><div className="font-cinzel font-bold text-base text-text-creme">{copySuccess ? 'Copiado!' : 'Copiar Texto'}</div><div className="text-[10px] text-text-terra font-bold uppercase">Clipboard</div></div>
                  </button>
                </div>
                <button onClick={() => setActiveStep('filler')} className="w-full mt-8 py-4 px-4 bg-bg-bordo-profundo text-text-creme border border-ouro-escuro/30 rounded-2xl font-cinzel font-bold uppercase transition-all">Reiniciar Campos</button>
              </div>

              <div className="lg:col-span-3 space-y-4 order-1 lg:order-2">
                <div className="bg-white border border-ouro-escuro/30 rounded-2xl shadow-2xl min-h-[1122px] relative overflow-hidden" style={{ backgroundImage: 'url(a.png)', backgroundSize: '100% 1122px', backgroundRepeat: 'repeat-y', backgroundPosition: 'top center' }}>
                  <div className="relative z-10 h-full flex flex-col" style={{ paddingTop: '280px', paddingBottom: '160px', paddingLeft: '28mm', paddingRight: '28mm' }}>
                    <div className={`ql-editor font-serif text-black leading-relaxed text-[17px]`} dangerouslySetInnerHTML={{ __html: DOMPurify.sanitize(filledText) }} />
                  </div>

                </div>
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      {/* Settings Modal */}
      <AnimatePresence>
        {showSettings && (
          <motion.div 
            initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
            className="fixed inset-0 bg-black/80 backdrop-blur-md z-50 flex items-center justify-center p-4"
          >
            <motion.div 
              initial={{ scale: 0.9, y: 20 }} animate={{ scale: 1, y: 0 }} exit={{ scale: 0.9, y: 20 }}
              className="bg-bg-noite border border-ouro-escuro/30 p-8 rounded-3xl max-w-md w-full shadow-2xl space-y-6"
            >
              <div className="flex items-center gap-3 text-ouro-imperial">
                <ShieldCheck size={32} />
                <h2 className="text-2xl font-cinzel font-bold uppercase">Segurança</h2>
              </div>
              <p className="text-sm text-text-terra leading-relaxed">
                As informações são processadas com segurança. A IA é gerenciada pelo servidor para proteger suas chaves.
              </p>

              <div className="space-y-2">
                <label className="flex items-center gap-2 text-xs font-cinzel font-bold uppercase tracking-widest text-text-areia">
                  <Lock size={12} /> Gemini API Key
                </label>
                <input 
                  type="password" 
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="Cole sua chave aqui..."
                  className="w-full bg-bg-ebano border border-bg-bordo rounded-xl px-4 py-3 outline-none focus:ring-1 focus:ring-ouro-imperial text-ouro-claro"
                />
              </div>

              <div className="border-t border-bg-bordo pt-6 space-y-4">
                <div className="flex items-center justify-between">
                  <div className="flex flex-col">
                    <span className="text-xs font-cinzel font-bold uppercase tracking-widest text-text-areia">Versão do App</span>
                    <span className="text-sm text-ouro-claro">v1.0.0</span>
                  </div>
                  {updateStatus === 'downloaded' ? (
                    <button 
                      onClick={handleRestartApp}
                      className="px-4 py-2 bg-green-600 text-white rounded-lg text-xs font-bold uppercase hover:bg-green-500 transition-all"
                    >
                      Reiniciar e Instalar
                    </button>
                  ) : (
                    <button 
                      onClick={handleCheckUpdates}
                      disabled={updateStatus === 'downloading'}
                      className="px-4 py-2 bg-bg-ebano border border-ouro-escuro/50 text-ouro-claro rounded-lg text-xs font-bold uppercase hover:bg-bg-bordo transition-all disabled:opacity-50"
                    >
                      {updateStatus === 'downloading' ? 'Verificando...' : 'Verificar Atualizações'}
                    </button>
                  )}
                </div>
                {updateMessage && (
                  <p className={`text-[10px] text-center ${updateStatus === 'error' ? 'text-red-400' : 'text-ouro-imperial'} font-medium animate-pulse`}>
                    {updateMessage}
                  </p>
                )}
              </div>

              <div className="flex gap-4 pt-4">
                <button onClick={() => setShowSettings(false)} className="flex-1 py-3 text-sm font-cinzel font-bold uppercase text-text-terra hover:text-text-creme">Cancelar</button>
                <button onClick={saveSettings} className="flex-1 py-3 bg-ouro-imperial text-bg-ebano rounded-xl font-cinzel font-bold uppercase hover:bg-ouro-claro transition-all">Salvar</button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
