
import React, { useEffect, useState } from 'react';
import { FunctionSquare, Sparkles } from 'lucide-react';

export interface FunctionDoc {
  name: string;
  params: string;
  desc: string;
}

export const SUPPORTED_FUNCTIONS: FunctionDoc[] = [
  { name: 'ABS', params: 'number', desc: '返回数字的绝对值' },
  { name: 'AVERAGE', params: 'number1, [number2], ...', desc: '返回参数的平均值' },
  { name: 'CONCAT', params: 'text1, [text2], ...', desc: '将多个文本项合并为一个' },
  { name: 'COUNT', params: 'value1, [value2], ...', desc: '计算包含数字的单元格个数' },
  { name: 'DATE', params: 'year, month, day', desc: '返回特定日期的序列号' },
  { name: 'IF', params: 'logical_test, value_if_true, value_if_false', desc: '根据条件判断返回对应值' },
  { name: 'LEN', params: 'text', desc: '返回文本字符串中的字符个数' },
  { name: 'LOWER', params: 'text', desc: '将文本转换为全小写' },
  { name: 'MAX', params: 'number1, [number2], ...', desc: '返回一组值中的最大值' },
  { name: 'MIN', params: 'number1, [number2], ...', desc: '返回一组值中的最小值' },
  { name: 'NOW', params: '', desc: '返回当前日期和时间的序列号' },
  { name: 'ROUND', params: 'number, num_digits', desc: '按指定位数对数字进行四舍五入' },
  { name: 'SUM', params: 'number1, [number2], ...', desc: '计算单元格区域中所有数值的和' },
  { name: 'TRIM', params: 'text', desc: '移除文本中多余的空格' },
  { name: 'UPPER', params: 'text', desc: '将文本转换为全大写' },
];

interface FormulaSuggestionsProps {
  filter: string;
  onSelect: (fnName: string) => void;
  onAIRequest: () => void;
  anchorRect: DOMRect | null;
}

const FormulaSuggestions: React.FC<FormulaSuggestionsProps> = ({ filter, onSelect, onAIRequest, anchorRect }) => {
  const [selectedIndex, setSelectedIndex] = useState(0);

  const filtered = SUPPORTED_FUNCTIONS.filter(f => 
    f.name.startsWith(filter.toUpperCase())
  ).slice(0, 10);

  useEffect(() => {
    setSelectedIndex(0);
  }, [filter]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setSelectedIndex(prev => (prev + 1) % (filtered.length + 1));
      } else if (e.key === 'ArrowUp') {
        e.preventDefault();
        setSelectedIndex(prev => (prev - 1 + (filtered.length + 1)) % (filtered.length + 1));
      } else if (e.key === 'Enter' || e.key === 'Tab') {
        e.preventDefault();
        if (selectedIndex === filtered.length) {
          onAIRequest();
        } else if (filtered[selectedIndex]) {
          onSelect(filtered[selectedIndex].name);
        }
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [filtered, selectedIndex, onSelect, onAIRequest]);

  if (!anchorRect || (filtered.length === 0 && !filter)) return null;

  return (
    <div 
      className="fixed z-[100] bg-white border border-slate-200 rounded-xl shadow-2xl w-64 overflow-hidden animate-in fade-in zoom-in-95 duration-100"
      style={{ 
        top: anchorRect.bottom + 4, 
        left: anchorRect.left,
        maxHeight: '320px' 
      }}
    >
      <div className="p-1 space-y-0.5 overflow-y-auto max-h-[260px]">
        {filtered.map((fn, idx) => (
          <button
            key={fn.name}
            className={`w-full flex items-center gap-3 px-3 py-2 text-left rounded-lg transition-colors ${idx === selectedIndex ? 'bg-blue-50 text-blue-700' : 'hover:bg-slate-50 text-slate-700'}`}
            onClick={() => onSelect(fn.name)}
          >
            <div className="shrink-0 text-slate-400 italic font-serif font-bold text-sm">fx</div>
            <div className="flex-1 min-w-0">
              <div className="text-xs font-bold tracking-tight">{fn.name}</div>
              <div className="text-[10px] text-slate-400 truncate">{fn.params}</div>
            </div>
          </button>
        ))}
      </div>
      
      <div className="border-t border-slate-100 bg-slate-50/50">
        <button
          className={`w-full flex items-center gap-3 px-4 py-3 text-left transition-colors ${selectedIndex === filtered.length ? 'bg-blue-50 text-blue-700' : 'hover:bg-slate-50 text-slate-600'}`}
          onClick={onAIRequest}
        >
          <Sparkles size={14} className="text-purple-500" />
          <span className="text-xs font-bold">AI 写公式</span>
        </button>
      </div>
    </div>
  );
};

export default FormulaSuggestions;
