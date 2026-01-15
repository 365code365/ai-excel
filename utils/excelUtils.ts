
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { SheetData, SheetRow } from '../types';

export const parseExcelFile = async (file: File): Promise<SheetData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        if (jsonData.length === 0) {
           resolve({ name: firstSheetName, columns: ['A'], rows: [{ 'A': null }] });
           return;
        }

        const maxCols = Math.max(...jsonData.map(r => r.length));
        const columns = Array.from({ length: maxCols }, (_, i) => {
          let label = '';
          let tempIdx = i;
          while (tempIdx >= 0) {
            label = String.fromCharCode((tempIdx % 26) + 65) + label;
            tempIdx = Math.floor(tempIdx / 26) - 1;
          }
          return label;
        });

        const rows = jsonData.map(rowArr => {
          const rowObj: SheetRow = {};
          columns.forEach((col, idx) => {
            rowObj[col] = rowArr[idx] !== undefined ? rowArr[idx] : null;
          });
          return rowObj;
        });
        
        resolve({
          name: firstSheetName,
          columns,
          rows
        });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsBinaryString(file);
  });
};

const hexToARGB = (hex?: string) => {
  if (!hex) return undefined;
  const cleanHex = hex.replace('#', '');
  return `FF${cleanHex.toUpperCase()}`;
};

export const exportExcelFile = async (data: SheetData, fileName: string) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet(data.name || 'Sheet1');

  // 设置列宽
  if (data.colWidths) {
    worksheet.columns = data.columns.map(col => ({
      header: col,
      key: col,
      width: (data.colWidths![col] || 100) / 7.2
    }));
  }

  // 填充数据并应用样式
  data.rows.forEach((row, rIdx) => {
    const newRow = worksheet.addRow(data.columns.map(col => {
      const val = row[col];
      if (typeof val === 'string' && val.startsWith('=')) return { formula: val.substring(1) };
      return val;
    }));

    if (data.rowHeights && data.rowHeights[rIdx] !== undefined) {
      newRow.height = data.rowHeights[rIdx] * 0.75;
    }

    data.columns.forEach((col, cIdx) => {
      const cell = newRow.getCell(cIdx + 1);
      const cellId = `${rIdx}-${col}`;
      const style = data.cellStyles?.[cellId];

      if (style) {
        if (style.backgroundColor) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: hexToARGB(style.backgroundColor) }
          };
        }
        cell.font = {
          bold: style.fontWeight === 'bold',
          italic: style.fontStyle === 'italic',
          underline: style.textDecoration === 'underline',
          size: style.fontSize || 10,
          name: style.fontFamily || '等线',
          color: { argb: hexToARGB(style.color) || 'FF000000' }
        };
        cell.alignment = {
          vertical: (style.verticalAlign as any) || 'middle',
          horizontal: (style.textAlign as any) || 'left',
          wrapText: style.wrapText || false
        };
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
          right: { style: 'thin', color: { argb: 'FFD1D5DB' } }
        };
      } else {
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        };
      }
    });
  });

  // 处理合并单元格
  if (data.merges && data.merges.length > 0) {
    data.merges.forEach(merge => {
      const minR = Math.min(merge.start.rowIndex, merge.end.rowIndex) + 1;
      const maxR = Math.max(merge.start.rowIndex, merge.end.rowIndex) + 1;
      const minC = Math.min(merge.start.colIndex, merge.end.colIndex) + 1;
      const maxC = Math.max(merge.start.colIndex, merge.end.colIndex) + 1;
      
      try {
        worksheet.mergeCells(minR, minC, maxR, maxC);
      } catch (e) {
        console.warn('Merge failed during export', e);
      }
    });
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `${fileName}.xlsx`;
  anchor.click();
};
