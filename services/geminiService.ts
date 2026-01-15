
import { GoogleGenAI, Type } from "@google/genai";
import { SheetData, AIProcessingResult, SelectionRange } from "../types";

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

export const processWithAI = async (
  prompt: string, 
  data: SheetData, 
  selection: SelectionRange | null
): Promise<AIProcessingResult> => {
  const model = "gemini-3-pro-preview";
  
  // 构建选区上下文
  let selectionContext = "";
  if (selection) {
    const minR = Math.min(selection.start.rowIndex, selection.end.rowIndex);
    const maxR = Math.max(selection.start.rowIndex, selection.end.rowIndex);
    const minC = Math.min(selection.start.colIndex, selection.end.colIndex);
    const maxC = Math.max(selection.start.colIndex, selection.end.colIndex);
    selectionContext = `Current User Selection: Rows ${minR+1} to ${maxR+1}, Columns ${data.columns[minC]} to ${data.columns[maxC]}.`;
  }

  const dataContext = JSON.stringify(data.rows.slice(0, 100));
  
  const response = await ai.models.generateContent({
    model: model,
    contents: {
      parts: [
        {
          text: `You are an advanced AI Excel assistant.
Sheet Name: ${data.name}
Existing Columns: ${data.columns.join(', ')}
${selectionContext}
Data Snapshot (JSON): ${dataContext}

Rules:
1. If the user asks for analysis, provide a clear 'insight'.
2. If the user asks to modify data, return the updated dataset in 'updatedRows'. The objects in the array should have keys corresponding to column labels (e.g., "A", "B", "C").
3. If the user asks for styling, return an array of updates in 'updatedStyles'. Each update has a 'cellId' (format: "rowIndex-colKey") and a 'style' object.
4. For charts, provide 'chartData' as array of {name, value}.
5. Ensure 'updatedRows' contains valid row objects matching the sheet column keys.

Request: ${prompt}`
        }
      ]
    },
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.OBJECT,
        properties: {
          updatedRows: {
            type: Type.ARRAY,
            items: { 
              type: Type.OBJECT,
              properties: {
                //满足非空对象要求，但指示AI可以使用任何列名
                "temp": { type: Type.STRING, description: "Placeholder for any column key like A, B, C..." }
              }
            }
          },
          updatedStyles: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                cellId: { type: Type.STRING },
                style: {
                  type: Type.OBJECT,
                  properties: {
                    backgroundColor: { type: Type.STRING },
                    color: { type: Type.STRING },
                    fontWeight: { type: Type.STRING }
                  }
                }
              },
              required: ["cellId", "style"]
            }
          },
          insight: { type: Type.STRING },
          newColumns: { type: Type.ARRAY, items: { type: Type.STRING } },
          chartData: {
            type: Type.ARRAY,
            items: {
              type: Type.OBJECT,
              properties: {
                name: { type: Type.STRING },
                value: { type: Type.NUMBER }
              },
              required: ["name", "value"]
            }
          },
          chartType: { type: Type.STRING }
        }
      }
    }
  });

  try {
    const text = response.text || '{}';
    const parsed = JSON.parse(text);
    
    // 过滤掉可能的 null 行，防止渲染崩溃
    if (parsed.updatedRows) {
      parsed.updatedRows = parsed.updatedRows.filter((row: any) => row !== null && typeof row === 'object');
    }

    if (Array.isArray(parsed.updatedStyles)) {
      const styleMap: any = {};
      parsed.updatedStyles.forEach((item: any) => {
        if (item.cellId && item.style) {
          styleMap[item.cellId] = item.style;
        }
      });
      parsed.updatedStyles = styleMap;
    }
    
    return parsed as AIProcessingResult;
  } catch (e) {
    console.error("AI Result Parsing Error", e);
    return { insight: "AI 响应解析失败，请稍后重试。" };
  }
};
