
import React, { useState, useEffect, useRef, useMemo } from 'react';
import { SheetData, SelectionRange, CellPosition, CellStyle } from '../types';
import FormulaSuggestions from './FormulaSuggestions';

interface ExcelTableProps {
  data: SheetData;
  onDataChange: (newData: SheetData) => void;
  selection: SelectionRange | null;
  setSelection: (selection: SelectionRange | null) => void;
}

const DEFAULT_COL_WIDTH = 100;
const DEFAULT_ROW_HEIGHT = 24;
const WPS_GREEN = '#107c41';

type SelectionMode = 'cell' | 'row' | 'col' | null;

const ExcelTable: React.FC<ExcelTableProps> = ({ data, onDataChange, selection, setSelection }) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editingPos, setEditingPos] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState('');
  const [isMouseDown, setIsMouseDown] = useState(false);
  const [selectionMode, setSelectionMode] = useState<SelectionMode>(null);
  
  const [suggestionFilter, setSuggestionFilter] = useState<string | null>(null);
  const [suggestionRect, setSuggestionRect] = useState<DOMRect | null>(null);

  const [resizingCol, setResizingCol] = useState<string | null>(null);
  const [resizingRow, setResizingRow] = useState<number | null>(null);
  const [startX, setStartX] = useState(0);
  const [startY, setStartY] = useState(0);
  const [startSize, setStartSize] = useState(0);

  const inputRef = useRef<HTMLInputElement>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  const tableRef = useRef<HTMLTableElement>(null);
  const [overlayStyle, setOverlayStyle] = useState<React.CSSProperties>({ display: 'none' });

  useEffect(() => {
    if (isEditing && inputRef.current) {
      inputRef.current.focus();
      inputRef.current.select();
    }
  }, [isEditing]);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const val = e.target.value;
    setEditValue(val);
    if (val.startsWith('=')) {
      const match = val.match(/=([A-Z]*)$/i);
      if (match) {
        setSuggestionFilter(match[1]);
        if (inputRef.current) setSuggestionRect(inputRef.current.getBoundingClientRect());
      } else {
        setSuggestionFilter(null);
      }
    } else {
      setSuggestionFilter(null);
    }
  };

  const onSelectSuggestion = (fnName: string) => {
    setEditValue(`=${fnName}(`);
    setSuggestionFilter(null);
    if (inputRef.current) inputRef.current.focus();
  };

  const handleSaveEdit = (targetPos?: CellPosition, value?: string) => {
    const pos = targetPos || editingPos;
    const val = value !== undefined ? value : editValue;
    if (!pos) {
      setIsEditing(false);
      setEditingPos(null);
      return;
    }
    const currentVal = data.rows[pos.rowIndex]?.[pos.colKey]?.toString() || '';
    if (val !== currentVal) {
      const newRows = [...data.rows];
      newRows[pos.rowIndex] = { ...newRows[pos.rowIndex], [pos.colKey]: val };
      onDataChange({ ...data, rows: newRows });
    }
    setIsEditing(false);
    setEditingPos(null);
    setSuggestionFilter(null);
    containerRef.current?.focus();
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (isEditing) {
       if (e.key === 'Escape') { e.preventDefault(); setIsEditing(false); containerRef.current?.focus(); }
       if (e.key === 'Enter') { e.preventDefault(); handleSaveEdit(); moveSelection(1, 0, false); }
       if (e.key === 'Tab') { e.preventDefault(); handleSaveEdit(); moveSelection(0, 1, false); }
       return;
    }

    const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
    const modKey = isMac ? e.metaKey : e.ctrlKey;

    if (!modKey) {
      switch(e.key) {
        case 'ArrowUp': e.preventDefault(); moveSelection(-1, 0, e.shiftKey); break;
        case 'ArrowDown': e.preventDefault(); moveSelection(1, 0, e.shiftKey); break;
        case 'ArrowLeft': e.preventDefault(); moveSelection(0, -1, e.shiftKey); break;
        case 'ArrowRight': e.preventDefault(); moveSelection(0, 1, e.shiftKey); break;
        case 'Enter': e.preventDefault(); if (selection) startEditing(selection.start, data.rows[selection.start.rowIndex]?.[selection.start.colKey]?.toString() || ''); break;
        case 'Tab': e.preventDefault(); moveSelection(0, 1, false); break;
        case 'Delete': 
        case 'Backspace': 
          e.preventDefault(); 
          clearSelectedContent(); 
          break;
        case 'Escape':
          e.preventDefault();
          setSelection(null);
          break;
      }
    }

    if (!modKey && e.key.length === 1 && !['ArrowUp','ArrowDown','ArrowLeft','ArrowRight','Tab','Enter','Escape'].includes(e.key)) {
      if (selection) startEditing(selection.start, e.key);
    }
  };

  const moveSelection = (rowDelta: number, colDelta: number, expand: boolean) => {
    if (!selection) return;
    const { start, end } = selection;
    const targetEndR = Math.min(data.rows.length - 1, Math.max(0, end.rowIndex + rowDelta));
    const targetEndC = Math.min(data.columns.length - 1, Math.max(0, end.colIndex + colDelta));
    const targetColKey = data.columns[targetEndC];

    const nextEnd = { rowIndex: targetEndR, colIndex: targetEndC, colKey: targetColKey };
    if (expand) {
      setSelection({ start, end: nextEnd });
    } else {
      setSelection({ start: nextEnd, end: nextEnd });
    }
  };

  const clearSelectedContent = () => {
    if (!selection) return;
    const { start, end } = selection;
    const minR = Math.min(start.rowIndex, end.rowIndex);
    const maxR = Math.max(start.rowIndex, end.rowIndex);
    const minC = Math.min(start.colIndex, end.colIndex);
    const maxC = Math.max(start.colIndex, end.colIndex);

    const newRows = [...data.rows];
    for (let r = minR; r <= maxR; r++) {
      const newRow = { ...newRows[r] };
      for (let c = minC; c <= maxC; c++) {
        newRow[data.columns[c]] = null;
      }
      newRows[r] = newRow;
    }
    onDataChange({ ...data, rows: newRows });
  };

  const startEditing = (pos: CellPosition, initialValue: string) => {
    setEditingPos(pos);
    setEditValue(initialValue);
    setIsEditing(true);
  };

  useEffect(() => {
    if (!selection || !tableRef.current) { setOverlayStyle({ display: 'none' }); return; }
    const minR = Math.min(selection.start.rowIndex, selection.end.rowIndex);
    const maxR = Math.max(selection.start.rowIndex, selection.end.rowIndex);
    const minC = Math.min(selection.start.colIndex, selection.end.colIndex);
    const maxC = Math.max(selection.start.colIndex, selection.end.colIndex);
    const startCellId = `cell-${minR}-${minC}`;
    const endCellId = `cell-${maxR}-${maxC}`;
    const startCell = tableRef.current.querySelector(`#${startCellId}`) as HTMLElement;
    const endCell = tableRef.current.querySelector(`#${endCellId}`) as HTMLElement;
    if (startCell && endCell) {
      setOverlayStyle({
        position: 'absolute',
        top: startCell.offsetTop - 1,
        left: startCell.offsetLeft - 1,
        width: (endCell.offsetLeft + endCell.offsetWidth) - startCell.offsetLeft + 1,
        height: (endCell.offsetTop + endCell.offsetHeight) - startCell.offsetTop + 1,
        border: `2px solid ${WPS_GREEN}`,
        backgroundColor: 'rgba(16, 124, 65, 0.05)',
        pointerEvents: 'none',
        zIndex: 30,
        display: 'block',
        boxSizing: 'border-box'
      });
    }
  }, [selection, data]);

  const onColResizeStart = (e: React.MouseEvent, col: string) => {
    e.stopPropagation(); setResizingCol(col); setStartX(e.pageX); setStartSize(data.colWidths?.[col] || DEFAULT_COL_WIDTH);
  };

  const onRowResizeStart = (e: React.MouseEvent, rowIndex: number) => {
    e.stopPropagation(); setResizingRow(rowIndex); setStartY(e.pageY); setStartSize(data.rowHeights?.[rowIndex] || DEFAULT_ROW_HEIGHT);
  };

  useEffect(() => {
    const handleMouseMove = (e: MouseEvent) => {
      if (resizingCol) {
        onDataChange({ ...data, colWidths: { ...(data.colWidths || {}), [resizingCol]: Math.max(40, startSize + e.pageX - startX) } });
      } else if (resizingRow !== null) {
        onDataChange({ ...data, rowHeights: { ...(data.rowHeights || {}), [resizingRow]: Math.max(20, startSize + e.pageY - startY) } });
      }
    };
    const handleMouseUp = () => { setResizingCol(null); setResizingRow(null); };
    if (resizingCol !== null || resizingRow !== null) { window.addEventListener('mousemove', handleMouseMove); window.addEventListener('mouseup', handleMouseUp); }
    return () => { window.removeEventListener('mousemove', handleMouseMove); window.removeEventListener('mouseup', handleMouseUp); };
  }, [resizingCol, resizingRow, startX, startY, startSize, data, onDataChange]);

  const handleMouseDown = (rowIndex: number, colIndex: number, colKey: string) => {
    if (isEditing && editingPos && (editingPos.rowIndex !== rowIndex || editingPos.colKey !== colKey)) handleSaveEdit();
    const pos = { rowIndex, colIndex, colKey };
    setSelection({ start: pos, end: pos });
    setIsMouseDown(true);
    setSelectionMode('cell');
    if (!(editingPos && editingPos.rowIndex === rowIndex && editingPos.colKey === colKey)) { setIsEditing(false); setEditingPos(null); }
    containerRef.current?.focus();
  };

  const handleMouseEnter = (rowIndex: number, colIndex: number, colKey: string) => {
    if (!isMouseDown || !selection) return;

    if (selectionMode === 'cell') {
      setSelection({ ...selection, end: { rowIndex, colIndex, colKey } });
    } else if (selectionMode === 'col') {
      setSelection({ ...selection, end: { rowIndex: data.rows.length - 1, colIndex, colKey } });
    } else if (selectionMode === 'row') {
      setSelection({ ...selection, end: { rowIndex, colIndex: data.columns.length - 1, colKey: data.columns[data.columns.length - 1] } });
    }
  };

  const handleColHeaderMouseDown = (e: React.MouseEvent, colKey: string, colIndex: number) => {
    if ((e.target as HTMLElement).classList.contains('resizer')) return;
    if (isEditing) handleSaveEdit();
    
    const startPos = { rowIndex: 0, colIndex, colKey };
    const endPos = { rowIndex: data.rows.length - 1, colIndex, colKey };
    
    setSelection({ start: startPos, end: endPos });
    setIsMouseDown(true);
    setSelectionMode('col');
    containerRef.current?.focus();
  };

  const handleColHeaderMouseEnter = (colKey: string, colIndex: number) => {
    if (!isMouseDown || selectionMode !== 'col' || !selection) return;
    setSelection({
      ...selection,
      end: { rowIndex: data.rows.length - 1, colIndex, colKey }
    });
  };

  const handleRowHeaderMouseDown = (e: React.MouseEvent, rowIndex: number) => {
    if ((e.target as HTMLElement).classList.contains('resizer')) return;
    if (isEditing) handleSaveEdit();
    
    const startPos = { rowIndex, colIndex: 0, colKey: data.columns[0] };
    const endPos = { rowIndex, colIndex: data.columns.length - 1, colKey: data.columns[data.columns.length - 1] };
    
    setSelection({ start: startPos, end: endPos });
    setIsMouseDown(true);
    setSelectionMode('row');
    containerRef.current?.focus();
  };

  const handleRowHeaderMouseEnter = (rowIndex: number) => {
    if (!isMouseDown || selectionMode !== 'row' || !selection) return;
    setSelection({
      ...selection,
      end: { rowIndex, colIndex: data.columns.length - 1, colKey: data.columns[data.columns.length - 1] }
    });
  };

  const handleSelectAll = () => {
    if (isEditing) handleSaveEdit();
    setSelection({
      start: { rowIndex: 0, colIndex: 0, colKey: data.columns[0] },
      end: { 
        rowIndex: data.rows.length - 1, 
        colIndex: data.columns.length - 1, 
        colKey: data.columns[data.columns.length - 1] 
      }
    });
    setSelectionMode(null);
    containerRef.current?.focus();
  };

  const handleMouseUpGlobal = () => {
    setIsMouseDown(false);
    setSelectionMode(null);
  };

  const selectionBounds = useMemo(() => {
    if (!selection) return null;
    return {
      minR: Math.min(selection.start.rowIndex, selection.end.rowIndex),
      maxR: Math.max(selection.start.rowIndex, selection.end.rowIndex),
      minC: Math.min(selection.start.colIndex, selection.end.colIndex),
      maxC: Math.max(selection.start.colIndex, selection.end.colIndex),
    };
  }, [selection]);

  const getMergeInfo = (rIdx: number, cIdx: number) => {
    if (!data.merges) return null;
    for (const merge of data.merges) {
      const minR = Math.min(merge.start.rowIndex, merge.end.rowIndex);
      const maxR = Math.max(merge.start.rowIndex, merge.end.rowIndex);
      const minC = Math.min(merge.start.colIndex, merge.end.colIndex);
      const maxC = Math.max(merge.start.colIndex, merge.end.colIndex);
      if (rIdx === minR && cIdx === minC) return { isRoot: true, rowSpan: maxR - minR + 1, colSpan: maxC - minC + 1 };
      if (rIdx >= minR && rIdx <= maxR && cIdx >= minC && cIdx <= maxC) return { isRoot: false };
    }
    return null;
  };

  const formatValue = (val: any, style?: CellStyle): string => {
    if (val === null || val === undefined) return '';
    const num = Number(val);
    if (isNaN(num)) return val.toString();
    if (style?.numberFormat === 'percent') return (num * 100).toFixed(style.precision ?? 0) + '%';
    if (style?.numberFormat === 'currency') return '¥' + num.toLocaleString(undefined, { minimumFractionDigits: style.precision ?? 2, maximumFractionDigits: style.precision ?? 2 });
    if (style?.numberFormat === 'number' || style?.precision !== undefined) return num.toLocaleString(undefined, { minimumFractionDigits: style?.precision ?? 0, maximumFractionDigits: style?.precision ?? 2 });
    return val.toString();
  };

  return (
    <div 
      ref={containerRef}
      className="flex-1 overflow-auto bg-slate-100 relative scrollbar-thin scrollbar-thumb-slate-300 select-none focus:outline-none" 
      onMouseUp={handleMouseUpGlobal} 
      onKeyDown={handleKeyDown}
      tabIndex={0}
    >
      <div className="relative inline-block align-top">
        <table ref={tableRef} className="border-separate border-spacing-0 bg-white shadow-sm ring-1 ring-slate-200">
          <thead className="sticky top-0 z-40">
            <tr className="bg-[#f2f2f2]">
              <th 
                className="w-10 border-r border-b border-slate-300 bg-[#f2f2f2] sticky left-0 z-50 shadow-[1px_0_0_#cbd5e1] hover:bg-slate-200 cursor-pointer active:bg-slate-300"
                onClick={handleSelectAll}
              >
              </th>
              {data.columns.map((col, idx) => (
                <th 
                  key={col} 
                  style={{ width: data.colWidths?.[col] || DEFAULT_COL_WIDTH }} 
                  className={`h-8 border-r border-b border-slate-300 px-2 text-center text-[11px] font-bold relative transition-colors cursor-pointer group ${selectionBounds && idx >= selectionBounds.minC && idx <= selectionBounds.maxC ? 'bg-[#dae8fc] text-blue-800 shadow-[inset_0_-2px_0_#107c41]' : 'text-slate-600 hover:bg-slate-200'}`}
                  onMouseDown={(e) => handleColHeaderMouseDown(e, col, idx)}
                  onMouseEnter={() => handleColHeaderMouseEnter(col, idx)}
                >
                  {col}
                  <div onMouseDown={(e) => onColResizeStart(e, col)} className="resizer absolute right-0 top-0 w-1.5 h-full cursor-col-resize z-10 hover:bg-blue-400/50" />
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.rows.map((row, rIdx) => {
              if (!row) return null; // 增加防御性检查
              return (
                <tr key={rIdx} style={{ height: data.rowHeights?.[rIdx] || DEFAULT_ROW_HEIGHT }}>
                  <td 
                    className={`w-10 border-r border-b border-slate-300 text-[10px] text-center font-bold sticky left-0 z-20 shadow-[1px_0_0_#cbd5e1] relative cursor-pointer group ${selectionBounds && rIdx >= selectionBounds.minR && rIdx <= selectionBounds.maxR ? 'bg-[#dae8fc] text-blue-800 shadow-[inset_-2px_0_0_#107c41,1px_0_0_#cbd5e1]' : 'bg-[#f2f2f2] text-slate-500 hover:bg-slate-200'}`}
                    onMouseDown={(e) => handleRowHeaderMouseDown(e, rIdx)}
                    onMouseEnter={() => handleRowHeaderMouseEnter(rIdx)}
                  >
                    {rIdx + 1}
                    <div onMouseDown={(e) => onRowResizeStart(e, rIdx)} className="resizer absolute bottom-0 left-0 w-full h-1.5 cursor-row-resize z-10 hover:bg-blue-400/50" />
                  </td>
                  {data.columns.map((col, cIdx) => {
                    const mergeInfo = getMergeInfo(rIdx, cIdx);
                    if (mergeInfo && !mergeInfo.isRoot) return null;
                    const cellId = `${rIdx}-${col}`;
                    const style = data.cellStyles?.[cellId];
                    const editing = isEditing && editingPos?.rowIndex === rIdx && editingPos?.colKey === col;
                    const active = selection?.start.rowIndex === rIdx && selection?.start.colIndex === cIdx;
                    let width = data.colWidths?.[col] || DEFAULT_COL_WIDTH;
                    if (mergeInfo?.isRoot) for (let i = 1; i < (mergeInfo.colSpan || 1); i++) width += (data.colWidths?.[data.columns[cIdx + i]] || DEFAULT_COL_WIDTH);
                    
                    const cellStyles: React.CSSProperties = {
                      backgroundColor: style?.backgroundColor,
                      color: style?.color,
                      fontWeight: style?.fontWeight,
                      fontStyle: style?.fontStyle,
                      textDecoration: style?.textDecoration,
                      fontSize: style?.fontSize ? `${style.fontSize}px` : undefined,
                      fontFamily: style?.fontFamily,
                      textAlign: style?.textAlign as any,
                      verticalAlign: style?.verticalAlign as any,
                      whiteSpace: style?.wrapText ? 'normal' : 'nowrap',
                      border: style?.border ? `1px solid ${style.borderColor || '#cbd5e1'}` : undefined,
                      width, minWidth: width, maxWidth: width
                    };

                    return (
                      <td 
                          id={`cell-${rIdx}-${cIdx}`} 
                          key={cellId} 
                          style={cellStyles} 
                          rowSpan={mergeInfo?.rowSpan} 
                          colSpan={mergeInfo?.colSpan}
                          onMouseDown={() => handleMouseDown(rIdx, cIdx, col)} 
                          onMouseEnter={() => handleMouseEnter(rIdx, cIdx, col)} 
                          onDoubleClick={() => startEditing({rowIndex:rIdx, colIndex:cIdx, colKey:col}, row[col]?.toString()||'')} 
                          className={`border-r border-b border-slate-200 text-xs px-1.5 py-0 truncate relative ${active ? 'z-10' : ''}`}
                      >
                        {editing ? (
                          <input 
                            ref={inputRef} 
                            className="absolute inset-0 w-full h-full px-1.5 py-0 border-none outline-none ring-2 ring-blue-500 bg-white z-[60]" 
                            value={editValue} 
                            onChange={handleInputChange} 
                            onBlur={() => !suggestionFilter && handleSaveEdit()}
                            onMouseDown={e => e.stopPropagation()} 
                          />
                        ) : (
                          <span className="block truncate">{formatValue(row[col], style)}</span>
                        )}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
        <div style={overlayStyle}>
          <div className="absolute -bottom-1 -right-1 w-2 h-2 bg-[#107c41] border border-white cursor-crosshair z-30"></div>
        </div>
      </div>
      {suggestionFilter !== null && <FormulaSuggestions filter={suggestionFilter} onSelect={onSelectSuggestion} onAIRequest={() => setSuggestionFilter(null)} anchorRect={suggestionRect} />}
    </div>
  );
};

export default ExcelTable;
