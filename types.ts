
import { CSSProperties } from 'react';

export interface CellStyle {
  backgroundColor?: string;
  color?: string;
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  textDecoration?: 'none' | 'underline';
  fontSize?: number;
  fontFamily?: string;
  border?: 'none' | 'thin' | 'medium' | 'thick';
  borderColor?: string;
  textAlign?: 'left' | 'center' | 'right';
  verticalAlign?: 'top' | 'middle' | 'bottom';
  wrapText?: boolean;
  numberFormat?: 'general' | 'number' | 'currency' | 'percent' | 'date';
  precision?: number;
}

export interface SheetRow {
  [key: string]: string | number | boolean | null;
}

export interface SheetData {
  name: string;
  columns: string[];
  rows: SheetRow[];
  rowHeights?: { [index: number]: number };
  colWidths?: { [key: string]: number };
  cellStyles?: { [cellId: string]: CellStyle }; // Format: "row-colKey"
  merges?: SelectionRange[];
}

export interface CellPosition {
  rowIndex: number;
  colIndex: number;
  colKey: string;
}

export interface SelectionRange {
  start: CellPosition;
  end: CellPosition;
}

export interface AIStyleUpdate {
  cellId: string;
  style: Partial<CellStyle>;
}

export interface AIProcessingResult {
  updatedRows?: SheetRow[];
  /**
   * The style updates returned by the AI, normalized to a map for easier merging.
   * Format: { "rowIndex-colKey": { ...styleProperties } }
   */
  updatedStyles?: { [cellId: string]: Partial<CellStyle> }; 
  insight?: string;
  newColumns?: string[];
  chartData?: {
    name: string;
    value: number;
  }[];
  chartType?: 'bar' | 'line' | 'pie';
}

export interface Message {
  id: string;
  role: 'user' | 'assistant' | 'system';
  content: string;
  timestamp: number;
  chart?: {
    data: any[];
    type: 'bar' | 'line' | 'pie';
  };
}
