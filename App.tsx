
import React, { useState, useCallback, useRef, useEffect } from 'react';
import { 
  FileSpreadsheet, Send, Sparkles, Undo2,
  Bold, Italic, AlignLeft, AlignCenter, AlignRight, Underline, Grid3X3,
  Scissors, Copy, Clipboard, X, Wand2, LayoutDashboard, BrainCircuit, 
  Table2, Baseline, Palette, AlignVerticalJustifyCenter, AlignVerticalJustifyStart, 
  AlignVerticalJustifyEnd, Combine, Eraser, CheckSquare, PanelRightClose, PanelRightOpen, MessageSquare
} from 'lucide-react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
  LineChart, Line, PieChart as RePieChart, Pie, Cell 
} from 'recharts';
import ExcelTable from './components/ExcelTable';
import { SheetData, Message, AIProcessingResult, SelectionRange, CellStyle } from './types';
import { parseExcelFile, exportExcelFile } from './utils/excelUtils';
import { processWithAI } from './services/geminiService';

const DEFAULT_ROWS = 50;
const DEFAULT_COLS = 20;
const INITIAL_COL_WIDTH = 100;
const INITIAL_ROW_HEIGHT = 24;
const MAX_HISTORY_STEPS = 50;

const createEmptySheet = (): SheetData => {
  const columns = Array.from({ length: DEFAULT_COLS }, (_, i) => {
    let label = '';
    let tempIdx = i;
    while (tempIdx >= 0) {
      label = String.fromCharCode((tempIdx % 26) + 65) + label;
      tempIdx = Math.floor(tempIdx / 26) - 1;
    }
    return label;
  });
  const colWidths: { [key: string]: number } = {};
  columns.forEach(col => colWidths[col] = INITIAL_COL_WIDTH);
  const rowHeights: { [index: number]: number } = {};
  for(let i=0; i<DEFAULT_ROWS; i++) rowHeights[i] = INITIAL_ROW_HEIGHT;
  const rows = Array.from({ length: DEFAULT_ROWS }, () => Object.fromEntries(columns.map(col => [col, null])));
  return { name: 'Sheet1', columns, rows, cellStyles: {}, colWidths, rowHeights, merges: [] };
};

const COLORS = ['#107c41', '#3b82f6', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899'];

const App: React.FC = () => {
  const [sheetData, setSheetData] = useState<SheetData>(createEmptySheet());
  const [history, setHistory] = useState<SheetData[]>([]);
  const [messages, setMessages] = useState<Message[]>([]);
  const [userInput, setUserInput] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [selection, setSelection] = useState<SelectionRange | null>(null);
  const [formulaValue, setFormulaValue] = useState('');
  const [isAssistantOpen, setIsAssistantOpen] = useState(true);
  
  const [showMergeModal, setShowMergeModal] = useState(false);
  const [mergeOption, setMergeOption] = useState<'standard' | 'content'>('standard');
  const [rememberMergeAction, setRememberMergeAction] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);
  const bgColorRef = useRef<HTMLInputElement>(null);
  const textColorRef = useRef<HTMLInputElement>(null);

  const updateSheetData = useCallback((newData: SheetData | ((prev: SheetData) => SheetData)) => {
    setHistory(prev => [sheetData, ...prev].slice(0, MAX_HISTORY_STEPS));
    if (typeof newData === 'function') setSheetData(prev => newData(prev));
    else setSheetData(newData);
  }, [sheetData]);

  const handleUndo = useCallback(() => {
    if (history.length === 0) return;
    const previousState = history[0];
    setSheetData(previousState);
    setHistory(prev => prev.slice(1));
  }, [history]);

  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const modKey = isMac ? e.metaKey : e.ctrlKey;
      if (modKey && e.key.toLowerCase() === 'z') {
        if (history.length > 0) {
          e.preventDefault();
          handleUndo();
        }
      }
    };
    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => window.removeEventListener('keydown', handleGlobalKeyDown);
  }, [history, handleUndo]);

  useEffect(() => {
    if (selection) {
      const { start } = selection;
      setFormulaValue(sheetData.rows[start.rowIndex]?.[start.colKey]?.toString() || '');
    } else setFormulaValue('');
  }, [selection, sheetData]);

  const applyStyleToSelection = useCallback((style: Partial<CellStyle>) => {
    if (!selection) return;
    const { start, end } = selection;
    const minR = Math.min(start.rowIndex, end.rowIndex);
    const maxR = Math.max(start.rowIndex, end.rowIndex);
    const minC = Math.min(start.colIndex, end.colIndex);
    const maxC = Math.max(start.colIndex, end.colIndex);
    const newStyles = { ...sheetData.cellStyles };
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const colKey = sheetData.columns[c];
        const cellId = `${r}-${colKey}`;
        newStyles[cellId] = { ...(newStyles[cellId] || {}), ...style };
      }
    }
    updateSheetData({ ...sheetData, cellStyles: newStyles });
  }, [selection, sheetData, updateSheetData]);

  const executeMerge = useCallback((type: 'standard' | 'content') => {
    if (!selection) return;
    const { start, end } = selection;
    const minR = Math.min(start.rowIndex, end.rowIndex);
    const maxR = Math.max(start.rowIndex, end.rowIndex);
    const minC = Math.min(start.colIndex, end.colIndex);
    const maxC = Math.max(start.colIndex, end.colIndex);
    const startColKey = sheetData.columns[minC];

    const newRows = [...sheetData.rows];
    let finalValue: string | null = null;

    if (type === 'standard') {
      finalValue = newRows[minR][startColKey] as string;
    } else {
      const contents: string[] = [];
      for (let r = minR; r <= maxR; r++) {
        for (let c = minC; c <= maxC; c++) {
          const val = newRows[r][sheetData.columns[c]];
          if (val !== null && val !== '') contents.push(val.toString());
        }
      }
      finalValue = contents.join(' ');
    }

    for (let r = minR; r <= maxR; r++) {
      newRows[r] = { ...newRows[r] };
      for (let c = minC; c <= maxC; c++) {
        newRows[r][sheetData.columns[c]] = (r === minR && c === minC) ? finalValue : null;
      }
    }

    const newStyles = { ...sheetData.cellStyles };
    newStyles[`${minR}-${startColKey}`] = { 
      ...(newStyles[`${minR}-${startColKey}`] || {}), 
      textAlign: 'center', 
      verticalAlign: 'middle' 
    };

    updateSheetData({ 
      ...sheetData, 
      rows: newRows, 
      cellStyles: newStyles,
      merges: [...(sheetData.merges || []), { start, end }]
    });
    setShowMergeModal(false);
  }, [selection, sheetData, updateSheetData]);

  const handleMergeToggle = useCallback(() => {
    if (!selection) return;
    const { start, end } = selection;
    const minR = Math.min(start.rowIndex, end.rowIndex);
    const maxR = Math.max(start.rowIndex, end.rowIndex);
    const minC = Math.min(start.colIndex, end.colIndex);
    const maxC = Math.max(start.colIndex, end.colIndex);

    const existingMergeIdx = (sheetData.merges || []).findIndex(m => 
      Math.min(m.start.rowIndex, m.end.rowIndex) === minR && 
      Math.max(m.start.rowIndex, m.end.rowIndex) === maxR &&
      Math.min(m.start.colIndex, m.end.colIndex) === minC &&
      Math.max(m.start.colIndex, m.end.colIndex) === maxC
    );

    if (existingMergeIdx !== -1) {
      const newMerges = [...(sheetData.merges || [])];
      newMerges.splice(existingMergeIdx, 1);
      updateSheetData({ ...sheetData, merges: newMerges });
      return;
    }

    let contentCount = 0;
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        const val = sheetData.rows[r][sheetData.columns[c]];
        if (val !== null && val !== '') contentCount++;
      }
    }

    if (contentCount > 1) {
      setShowMergeModal(true);
    } else {
      executeMerge('standard');
    }
  }, [selection, sheetData, updateSheetData, executeMerge]);

  const handleClearFormat = useCallback(() => {
    if (!selection) return;
    const { start, end } = selection;
    const minR = Math.min(start.rowIndex, end.rowIndex);
    const maxR = Math.max(start.rowIndex, end.rowIndex);
    const minC = Math.min(start.colIndex, end.colIndex);
    const maxC = Math.max(start.colIndex, end.colIndex);
    const newStyles = { ...sheetData.cellStyles };
    for (let r = minR; r <= maxR; r++) {
      for (let c = minC; c <= maxC; c++) {
        delete newStyles[`${r}-${sheetData.columns[c]}`];
      }
    }
    updateSheetData({ ...sheetData, cellStyles: newStyles });
  }, [selection, sheetData, updateSheetData]);

  const runAICommand = async (command: string) => {
    if (!isAssistantOpen) setIsAssistantOpen(true);
    setIsProcessing(true);
    try {
      const result = await processWithAI(command, sheetData, selection);
      const updatedSheet = { ...sheetData };
      let hasChanges = false;

      if (result.updatedRows && Array.isArray(result.updatedRows)) {
        const cleanedRows = result.updatedRows.filter(r => r !== null);
        if (cleanedRows.length > 0) {
          updatedSheet.rows = cleanedRows;
          if (result.newColumns) updatedSheet.columns = [...new Set([...updatedSheet.columns, ...result.newColumns])];
          hasChanges = true;
        }
      }

      if (result.updatedStyles && typeof result.updatedStyles === 'object') {
        const mergedStyles = { ...updatedSheet.cellStyles };
        Object.entries(result.updatedStyles).forEach(([cellId, style]) => {
          if (style && typeof style === 'object') {
            mergedStyles[cellId] = { ...(mergedStyles[cellId] || {}), ...style } as CellStyle;
          }
        });
        updatedSheet.cellStyles = mergedStyles;
        hasChanges = true;
      }

      if (hasChanges) updateSheetData(updatedSheet);

      const assistantMessage: Message = { 
        id: Date.now().toString(), 
        role: 'assistant', 
        content: result.insight || "已完成您的 AI 任务。", 
        timestamp: Date.now(),
        chart: (result.chartData && result.chartData.length > 0) ? { data: result.chartData, type: result.chartType || 'bar' } : undefined
      };
      setMessages(prev => [...prev, assistantMessage]);
    } catch (e) { console.error(e); } finally { setIsProcessing(false); }
  };

  const handleSendMessage = async () => {
    if (!userInput.trim() || isProcessing) return;
    const userMessage: Message = { id: Date.now().toString(), role: 'user', content: userInput, timestamp: Date.now() };
    setMessages(prev => [...prev, userMessage]);
    const input = userInput;
    setUserInput('');
    await runAICommand(input);
  };

  const renderChart = (chart: any) => {
    if (!chart || !chart.data || chart.data.length === 0) return null;
    return (
      <div className="mt-3 p-2 bg-white rounded-lg border border-slate-100 shadow-inner overflow-hidden" style={{ height: '200px', width: '100%' }}>
        <ResponsiveContainer width="100%" height="100%">
          {chart.type === 'bar' ? (
            <BarChart data={chart.data} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
              <XAxis dataKey="name" fontSize={8} />
              <YAxis fontSize={8} />
              <Tooltip contentStyle={{fontSize: '10px'}} />
              <Bar dataKey="value" fill="#107c41" radius={[2, 2, 0, 0]} />
            </BarChart>
          ) : chart.type === 'pie' ? (
            <RePieChart>
              <Pie data={chart.data} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={50}>
                {chart.data.map((_: any, index: number) => <Cell key={index} fill={COLORS[index % COLORS.length]} />)}
              </Pie>
              <Tooltip />
            </RePieChart>
          ) : (
            <LineChart data={chart.data} margin={{ top: 5, right: 5, left: -25, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
              <XAxis dataKey="name" fontSize={8} />
              <YAxis fontSize={8} />
              <Tooltip />
              <Line type="monotone" dataKey="value" stroke="#107c41" strokeWidth={2} />
            </LineChart>
          )}
        </ResponsiveContainer>
      </div>
    );
  };

  return (
    <div className="h-screen flex flex-col bg-[#f3f3f3] text-slate-900 overflow-hidden select-none font-sans relative">
      {showMergeModal && (
        <div className="fixed inset-0 z-[100] bg-black/30 backdrop-blur-sm flex items-center justify-center animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-2xl w-[460px] border border-slate-200 overflow-hidden flex flex-col animate-in zoom-in-95 duration-200">
            <div className="p-4 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
               <h3 className="text-sm font-bold text-slate-700">WPS 表格</h3>
               <button onClick={() => setShowMergeModal(false)} className="text-slate-400 hover:text-red-500 transition-colors"><X size={16} /></button>
            </div>
            <div className="p-6">
               <p className="text-[13px] text-slate-600 mb-6">选定区域包含不同的内容，请选择您期望的合并方式。</p>
               <div className="grid grid-cols-2 gap-4 mb-6">
                  <div 
                    onClick={() => setMergeOption('standard')}
                    className={`cursor-pointer rounded-lg border-2 p-3 flex flex-col items-center gap-4 transition-all hover:shadow-md ${mergeOption === 'standard' ? 'border-[#3b82f6] bg-blue-50/50' : 'border-slate-100 hover:border-slate-300'}`}
                  >
                    <div className="w-24 h-20 bg-white border border-slate-200 rounded relative overflow-hidden flex items-center justify-center">
                       <div className="grid grid-cols-1 grid-rows-3 w-10 gap-1 opacity-40">
                          <div className="h-2 bg-orange-400 rounded"></div>
                          <div className="h-2 bg-green-500 rounded"></div>
                          <div className="h-2 bg-slate-300 rounded"></div>
                       </div>
                       <div className="mx-2 text-slate-300">>>></div>
                       <div className="w-8 h-4 bg-orange-400 rounded"></div>
                    </div>
                    <span className="text-xs font-medium text-slate-700">合并居中 (常规)</span>
                  </div>
                  <div 
                    onClick={() => setMergeOption('content')}
                    className={`cursor-pointer rounded-lg border-2 p-3 flex flex-col items-center gap-4 transition-all hover:shadow-md ${mergeOption === 'content' ? 'border-[#3b82f6] bg-blue-50/50' : 'border-slate-100 hover:border-slate-300'}`}
                  >
                    <div className="w-24 h-20 bg-white border border-slate-200 rounded relative overflow-hidden flex items-center justify-center">
                       <div className="grid grid-cols-1 grid-rows-3 w-10 gap-1 opacity-40">
                          <div className="h-2 bg-orange-400 rounded"></div>
                          <div className="h-2 bg-green-500 rounded"></div>
                          <div className="h-2 bg-slate-300 rounded"></div>
                       </div>
                       <div className="mx-2 text-slate-300">>>></div>
                       <div className="flex flex-col gap-0.5">
                          <div className="w-6 h-1.5 bg-orange-400 rounded-full"></div>
                          <div className="w-6 h-1.5 bg-green-500 rounded-full"></div>
                          <div className="w-6 h-1.5 bg-slate-300 rounded-full"></div>
                       </div>
                    </div>
                    <span className="text-xs font-medium text-slate-700">合并内容</span>
                  </div>
               </div>
               <div className="flex items-center gap-2 mb-6">
                  <button onClick={() => setRememberMergeAction(!rememberMergeAction)} className={`w-4 h-4 rounded border flex items-center justify-center transition-colors ${rememberMergeAction ? 'bg-[#3b82f6] border-[#3b82f6]' : 'bg-white border-slate-300'}`}>
                    {rememberMergeAction && <CheckSquare size={12} className="text-white" />}
                  </button>
                  <span className="text-xs text-slate-500">记住此次操作</span>
               </div>
               <div className="flex justify-end gap-3">
                  <button onClick={() => setShowMergeModal(false)} className="px-6 py-1.5 border border-slate-200 rounded text-xs font-medium hover:bg-slate-50 transition-colors">取消</button>
                  <button onClick={() => executeMerge(mergeOption)} className="px-8 py-1.5 bg-[#3b82f6] text-white rounded text-xs font-medium shadow-lg shadow-blue-200 hover:bg-blue-600 transition-all">确定</button>
               </div>
            </div>
          </div>
        </div>
      )}

      <header className="bg-white border-b border-slate-200 px-4 py-1 flex items-center justify-between z-30 shrink-0 shadow-sm h-10">
        <div className="flex items-center gap-6">
          <div className="flex items-center gap-2">
            <div className="bg-[#107c41] p-1 rounded text-white"><FileSpreadsheet size={16} /></div>
            <span className="font-bold text-sm tracking-tight text-slate-700">SmartSheet AI</span>
          </div>
        </div>
        <div className="flex items-center gap-3">
          <button onClick={handleUndo} disabled={history.length === 0} className={`p-1.5 rounded transition-colors ${history.length === 0 ? 'text-slate-300' : 'text-slate-600 hover:bg-slate-100'}`} title="撤回 (Ctrl+Z)"><Undo2 size={16} /></button>
          <div className="w-[1px] h-4 bg-slate-200"></div>
          <button onClick={() => exportExcelFile(sheetData, sheetData.name)} className="flex items-center gap-2 px-3 py-1 text-xs font-bold text-[#107c41] bg-white border border-[#107c41] rounded hover:bg-green-50 shadow-sm transition-all active:scale-95">导出</button>
          <button onClick={() => fileInputRef.current?.click()} className="flex items-center gap-2 px-3 py-1 text-xs font-bold text-white bg-[#107c41] rounded hover:bg-[#0a5e31] shadow-sm transition-all active:scale-95">导入</button>
          <div className="w-[1px] h-4 bg-slate-200 mx-1"></div>
          <button 
            onClick={() => setIsAssistantOpen(!isAssistantOpen)} 
            className={`flex items-center gap-2 px-3 py-1 text-xs font-bold rounded transition-all shadow-sm active:scale-95 ${isAssistantOpen ? 'bg-slate-100 text-slate-600 border border-slate-200' : 'bg-blue-500 text-white border border-blue-600 hover:bg-blue-600'}`}
          >
            {isAssistantOpen ? <PanelRightClose size={14} /> : <PanelRightOpen size={14} />}
            {isAssistantOpen ? '收起助手' : '智能助手'}
          </button>
          <input type="file" ref={fileInputRef} className="hidden" accept=".xlsx, .xls" onChange={async (e) => {
             const file = e.target.files?.[0];
             if (file) { const data = await parseExcelFile(file); updateSheetData(data); }
          }} />
        </div>
      </header>

      <div className="bg-white border-b border-slate-300 flex items-center h-[92px] shrink-0 overflow-x-auto no-scrollbar shadow-inner px-2 py-1 gap-1">
        <div className="flex flex-col items-center border-r border-slate-200 px-3 h-full justify-between pb-1">
           <div className="flex items-center gap-1 mt-1">
             <button onClick={() => runAICommand("粘贴")} className="flex flex-col items-center hover:bg-slate-100 p-2 rounded transition-colors group">
               <Clipboard size={20} className="text-slate-600 group-hover:text-blue-600" />
               <span className="text-[10px] mt-0.5">粘贴</span>
             </button>
             <div className="flex flex-col gap-0.5">
               <button onClick={() => runAICommand("剪切")} className="flex items-center gap-2 hover:bg-slate-100 px-2 py-1 rounded text-[10px] transition-colors"><Scissors size={14}/> 剪切</button>
               <button onClick={() => runAICommand("复制")} className="flex items-center gap-2 hover:bg-slate-100 px-2 py-1 rounded text-[10px] transition-colors"><Copy size={14}/> 复制</button>
             </div>
           </div>
           <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">编辑</span>
        </div>

        <div className="flex flex-col items-center border-r border-slate-200 px-3 h-full justify-between pb-1">
           <div className="flex flex-col gap-1.5 mt-1">
             <div className="flex gap-1 items-center">
                <select className="text-[11px] border border-slate-200 px-1 py-0.5 h-7 w-28 rounded focus:ring-1 focus:ring-blue-400 outline-none" onChange={(e) => applyStyleToSelection({fontFamily: e.target.value})}>
                   {['等线', '微软雅黑', 'Arial', 'Times New Roman'].map(f => <option key={f} value={f}>{f}</option>)}
                </select>
                <select className="text-[11px] border border-slate-200 px-1 py-0.5 h-7 w-14 rounded focus:ring-1 focus:ring-blue-400 outline-none" onChange={(e) => applyStyleToSelection({fontSize: parseInt(e.target.value)})}>
                   {[8, 9, 10, 11, 12, 14, 16, 18, 20].map(s => <option key={s} value={s}>{s}</option>)}
                </select>
             </div>
             <div className="flex gap-0.5 items-center">
                <button onClick={() => applyStyleToSelection({fontWeight: 'bold'})} className="p-1.5 hover:bg-slate-100 rounded transition-colors" title="加粗"><Bold size={15} /></button>
                <button onClick={() => applyStyleToSelection({fontStyle: 'italic'})} className="p-1.5 hover:bg-slate-100 rounded transition-colors" title="倾斜"><Italic size={15} /></button>
                <button onClick={() => applyStyleToSelection({textDecoration: 'underline'})} className="p-1.5 hover:bg-slate-100 rounded transition-colors" title="下划线"><Underline size={15} /></button>
                <div className="w-[1px] h-4 bg-slate-200 mx-1"></div>
                <button onClick={() => applyStyleToSelection({border: 'thin', borderColor: '#cbd5e1'})} className="p-1.5 hover:bg-slate-100 rounded transition-colors" title="边框"><Grid3X3 size={15}/></button>
                <button onClick={() => bgColorRef.current?.click()} className="p-1.5 hover:bg-slate-100 rounded transition-colors relative group" title="填充颜色">
                   <Palette size={15}/>
                   <div className="absolute bottom-1 left-1 right-1 h-1 bg-yellow-400 rounded-full"></div>
                   <input type="color" ref={bgColorRef} className="hidden" onChange={(e) => applyStyleToSelection({backgroundColor: e.target.value})} />
                </button>
                <button onClick={() => textColorRef.current?.click()} className="p-1.5 hover:bg-slate-100 rounded transition-colors relative group" title="字体颜色">
                   <Baseline size={15}/>
                   <div className="absolute bottom-1 left-1 right-1 h-1 bg-red-500 rounded-full"></div>
                   <input type="color" ref={textColorRef} className="hidden" onChange={(e) => applyStyleToSelection({color: e.target.value})} />
                </button>
             </div>
           </div>
           <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">字体</span>
        </div>

        <div className="flex flex-col items-center border-r border-slate-200 px-3 h-full justify-between pb-1">
           <div className="flex flex-col gap-1.5 mt-1">
             <div className="flex gap-1 justify-center">
                <button onClick={() => applyStyleToSelection({verticalAlign: 'top'})} className="p-1.5 hover:bg-slate-100 rounded" title="顶端对齐"><AlignVerticalJustifyStart size={15} /></button>
                <button onClick={() => applyStyleToSelection({verticalAlign: 'middle'})} className="p-1.5 hover:bg-slate-100 rounded" title="垂直居中"><AlignVerticalJustifyCenter size={15} /></button>
                <button onClick={() => applyStyleToSelection({verticalAlign: 'bottom'})} className="p-1.5 hover:bg-slate-100 rounded" title="底端对齐"><AlignVerticalJustifyEnd size={15} /></button>
             </div>
             <div className="flex gap-1 justify-center">
                <button onClick={() => applyStyleToSelection({textAlign: 'left'})} className="p-1.5 hover:bg-slate-100 rounded" title="左对齐"><AlignLeft size={15} /></button>
                <button onClick={() => applyStyleToSelection({textAlign: 'center'})} className="p-1.5 hover:bg-slate-100 rounded" title="居中"><AlignCenter size={15} /></button>
                <button onClick={() => applyStyleToSelection({textAlign: 'right'})} className="p-1.5 hover:bg-slate-100 rounded" title="右对齐"><AlignRight size={15} /></button>
             </div>
           </div>
           <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">对齐方式</span>
        </div>

        <div className="flex flex-col items-center border-r border-slate-200 px-3 h-full justify-between pb-1">
           <div className="flex flex-col gap-1 mt-1">
             <button onClick={handleMergeToggle} className="flex items-center gap-2 hover:bg-slate-100 px-3 py-1.5 rounded text-[11px] font-medium border border-transparent hover:border-slate-200 transition-all">
               <Combine size={16} className="text-blue-600" /> 合并后居中
             </button>
             <button onClick={handleClearFormat} className="flex items-center gap-2 hover:bg-slate-100 px-3 py-1.5 rounded text-[11px] font-medium border border-transparent hover:border-slate-200 transition-all">
               <Eraser size={16} className="text-orange-500" /> 清除格式
             </button>
           </div>
           <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">单元格</span>
        </div>

        <div className="flex flex-col items-center border-r border-slate-200 px-4 h-full justify-between pb-1 bg-gradient-to-b from-blue-50/50 to-transparent">
           <div className="flex gap-4 items-center py-2 mt-1">
             <button onClick={() => runAICommand("美化表格样式")} className="flex flex-col items-center group">
               <div className="p-2 bg-blue-100 text-blue-600 rounded-xl group-hover:bg-blue-600 group-hover:text-white transition-all shadow-sm">
                 <Wand2 size={18} />
               </div>
               <span className="text-[10px] mt-1 font-bold text-blue-700">一键美化</span>
             </button>
             <button onClick={() => runAICommand("智能填充")} className="flex flex-col items-center group">
               <div className="p-2 bg-purple-100 text-purple-600 rounded-xl group-hover:bg-purple-600 group-hover:text-white transition-all shadow-sm">
                 <BrainCircuit size={18} />
               </div>
               <span className="text-[10px] mt-1 font-bold text-purple-700">智能填充</span>
             </button>
             <button onClick={() => runAICommand("数据分析")} className="flex flex-col items-center group">
               <div className="p-2 bg-orange-100 text-orange-600 rounded-xl group-hover:bg-orange-600 group-hover:text-white transition-all shadow-sm">
                 <LayoutDashboard size={18} />
               </div>
               <span className="text-[10px] mt-1 font-bold text-orange-700">深度分析</span>
             </button>
           </div>
           <span className="text-[10px] text-blue-500 font-black uppercase tracking-widest flex items-center gap-1"><Sparkles size={8} className="animate-pulse" /> AI 智能助手</span>
        </div>
      </div>

      <div className="flex-1 flex overflow-hidden relative">
        <div className="flex-1 flex flex-col overflow-hidden bg-[#e6e6e6]">
          <div className="bg-white border-b border-slate-300 flex items-center h-8 px-2 shrink-0 shadow-sm relative z-20">
            <div className="flex items-center justify-center min-w-[70px] h-full border-r border-slate-200 mr-2 text-[11px] font-bold text-[#107c41] bg-[#f9f9f9] italic">
              {selection ? `${selection.start.colKey}${selection.start.rowIndex + 1}` : 'fx'}
            </div>
            <input 
              type="text" 
              value={formulaValue}
              onChange={(e) => setFormulaValue(e.target.value)}
              placeholder="输入内容、公式或 AI 指令..."
              className="flex-1 px-2 h-full text-xs outline-none bg-transparent font-medium text-slate-700"
              disabled={!selection}
              onKeyDown={e => {
                if(e.key === 'Enter' && formulaValue.startsWith('/ai')) {
                  runAICommand(formulaValue.replace('/ai', ''));
                  setFormulaValue('');
                }
              }}
            />
          </div>
          <div className="flex-1 flex flex-col overflow-hidden relative">
            <ExcelTable data={sheetData} onDataChange={updateSheetData} selection={selection} setSelection={setSelection} />
          </div>
          <footer className="bg-white border-t border-slate-200 px-4 h-6 flex items-center justify-between shrink-0 text-[9px] text-slate-500 font-bold uppercase">
             <div className="flex items-center gap-4">
               <span>就绪</span>
               {selection && <span className="text-blue-600">已选中: {selection.start.colKey}{selection.start.rowIndex+1}:{selection.end.colKey}{selection.end.rowIndex+1}</span>}
             </div>
             <div className="flex items-center gap-2">
               <Table2 size={10} />
               <span>100% (Gemini 3 Pro)</span>
             </div>
          </footer>
        </div>

        {/* 智能助手侧边栏 */}
        <aside 
          className={`bg-white border-l border-slate-200 flex flex-col shadow-2xl z-40 transition-all duration-300 ease-in-out ${isAssistantOpen ? 'w-[380px] opacity-100' : 'w-0 opacity-0 overflow-hidden'}`}
        >
          <div className="p-4 border-b border-slate-100 bg-slate-50/80 flex justify-between items-center shrink-0">
            <h2 className="font-black text-[11px] uppercase tracking-[0.15em] text-slate-500 flex items-center gap-2">
              <Sparkles size={16} className="text-blue-500 animate-pulse" /> 智能助手
            </h2>
            <div className="flex items-center gap-2">
              <button onClick={() => setMessages([])} className="text-[10px] text-slate-400 hover:text-red-500 font-bold uppercase transition-colors px-2">清除对话</button>
              <button onClick={() => setIsAssistantOpen(false)} className="text-slate-400 hover:text-slate-600 transition-colors"><X size={16} /></button>
            </div>
          </div>
          <div className="flex-1 overflow-y-auto p-4 space-y-4 no-scrollbar bg-slate-50/30">
            {messages.length === 0 && (
              <div className="h-full flex flex-col items-center justify-center text-slate-400 gap-4 opacity-50 text-center px-10">
                <BrainCircuit size={48} className="text-slate-200" />
                <p className="text-xs font-bold leading-relaxed">选中数据并提问，我可以帮您分析趋势、美化样式或自动生成内容。</p>
              </div>
            )}
            {messages.map(msg => (
              <div key={msg.id} className={`flex ${msg.role === 'user' ? 'justify-end' : 'justify-start'} animate-in fade-in slide-in-from-bottom-2`}>
                <div className={`max-w-[90%] p-4 rounded-2xl text-[12px] leading-relaxed shadow-sm ${msg.role === 'user' ? 'bg-[#107c41] text-white' : 'bg-white border border-slate-100 text-slate-700'}`}>
                  {msg.content}
                  {msg.chart && renderChart(msg.chart)}
                  <div className="text-[9px] mt-2 opacity-50 text-right">
                    {new Date(msg.timestamp).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}
                  </div>
                </div>
              </div>
            ))}
            {isProcessing && (
              <div className="flex gap-2 items-center p-3 text-xs text-blue-500 font-bold animate-pulse">
                <BrainCircuit size={16} className="animate-spin duration-[3000ms]" />
                AI 正在生成深度洞察...
              </div>
            )}
          </div>
          <div className="p-4 bg-white border-t border-slate-100 shrink-0">
            <div className="relative group">
              <textarea 
                rows={3} 
                value={userInput} 
                onChange={e => setUserInput(e.target.value)} 
                onKeyDown={e => { if(e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); handleSendMessage(); } }}
                placeholder="询问 AI 或 输入指令..." 
                className="w-full p-4 pr-12 text-xs border border-slate-200 rounded-2xl focus:ring-2 focus:ring-[#107c41] focus:border-transparent outline-none resize-none transition-all shadow-sm group-hover:shadow-md" 
                disabled={isProcessing} 
              />
              <button 
                onClick={handleSendMessage} 
                disabled={isProcessing || !userInput.trim()} 
                className={`absolute right-3 bottom-3 p-2 rounded-xl transition-all ${userInput.trim() ? 'bg-[#107c41] text-white shadow-lg' : 'bg-slate-100 text-slate-400'}`}
              >
                <Send size={16} />
              </button>
            </div>
          </div>
        </aside>

        {/* 悬浮唤醒助手按钮 (仅在侧边栏关闭时显示) */}
        {!isAssistantOpen && (
          <button 
            onClick={() => setIsAssistantOpen(true)}
            className="fixed bottom-10 right-10 z-[60] bg-[#107c41] text-white p-4 rounded-full shadow-2xl hover:scale-110 active:scale-95 transition-all animate-in zoom-in-0 duration-300 group"
            title="开启助手"
          >
            <MessageSquare size={24} />
            <span className="absolute right-full mr-3 bg-slate-800 text-white text-[10px] px-2 py-1 rounded whitespace-nowrap opacity-0 group-hover:opacity-100 transition-opacity">点击开启智能助手</span>
          </button>
        )}
      </div>
    </div>
  );
};

export default App;
