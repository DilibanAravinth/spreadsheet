import React, {
  useState,
  useCallback,
  useMemo,
  useRef,
  useEffect,
} from "react";
import { cn } from "../lib/utils";
import useWindowDimensions from "../hooks/useWindowDimensions";

interface CellData {
  value: string;
  formula?: string;
  computed?: number | string;
}

interface CellPosition {
  row: number;
  col: number;
}

const CELL_WIDTH = 100;
const CELL_HEIGHT = 30;
const HEADER_HEIGHT = 30;
const ROW_HEADER_WIDTH = 60;
const TOTAL_ROWS = 10000;
const TOTAL_COLS = 10000;

const getColumnLabel = (col: number): string => {
  let label = "";
  while (col >= 0) {
    label = String.fromCharCode(65 + (col % 26)) + label;
    col = Math.floor(col / 26) - 1;
  }
  return label;
};

const parseReference = (ref: string): CellPosition | null => {
  const match = ref.match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;

  const [, colStr, rowStr] = match;
  let col = 0;
  for (let i = 0; i < colStr.length; i++) {
    col = col * 26 + (colStr.charCodeAt(i) - 64);
  }
  col -= 1;

  const row = parseInt(rowStr) - 1;
  return { row, col };
};

const evaluateFormula = (
  formula: string,
  cells: Map<string, CellData>
): number | string => {
  try {
    const expression = formula.substring(1);
    const processedExpression = expression.replace(/[A-Z]+\d+/g, (match) => {
      const pos = parseReference(match);
      if (!pos) return "0";

      const cellKey = `${pos.row}-${pos.col}`;
      const cell = cells.get(cellKey);

      if (!cell) return "0";
      if (cell.computed !== undefined) return cell.computed.toString();

      const numValue = parseFloat(cell.value);
      return isNaN(numValue) ? "0" : numValue.toString();
    });

    const result = Function(`"use strict"; return (${processedExpression})`)();
    return typeof result === "number" ? result : result.toString();
  } catch {
    return "#ERROR";
  }
};

const VirtualizedExcelSheet: React.FC = () => {
  const [cells, setCells] = useState<Map<string, CellData>>(new Map());
  const [selectedCell, setSelectedCell] = useState<CellPosition>({
    row: 0,
    col: 0,
  });
  const [editingCell, setEditingCell] = useState<CellPosition | null>(null);
  const [editValue, setEditValue] = useState("");
  const [scrollTop, setScrollTop] = useState(0);
  const [scrollLeft, setScrollLeft] = useState(0);

  const containerRef = useRef<HTMLDivElement>(null);
  const inputRef = useRef<HTMLInputElement>(null);
  const { height, width } = useWindowDimensions();

  const visibleRange = useMemo(() => {
    const startRow = Math.floor(scrollTop / CELL_HEIGHT);
    const endRow = Math.min(
      startRow + Math.floor(height / CELL_HEIGHT) + 2,
      TOTAL_ROWS
    );
    const startCol = Math.floor(scrollLeft / CELL_WIDTH);
    const endCol = Math.min(
      startCol + Math.floor(width / CELL_WIDTH) + 2,
      TOTAL_COLS
    );
    const rowLength = endRow - startRow;
    return { rowLength, startRow, endRow, startCol, endCol };
  }, [scrollTop, scrollLeft]);

  const getCellData = useCallback(
    (row: number, col: number): CellData | null => {
      const key = `${row}-${col}`;
      return cells.get(key) || null;
    },
    [cells]
  );

  const updateCell = useCallback((row: number, col: number, value: string) => {
    setCells((prev) => {
      const newCells = new Map(prev);
      const key = `${row}-${col}`;

      if (value === "") {
        newCells.delete(key);
      } else {
        const isFormula = value.startsWith("=");
        const cellData: CellData = {
          value,
          formula: isFormula ? value : undefined,
          computed: isFormula ? evaluateFormula(value, newCells) : undefined,
        };
        newCells.set(key, cellData);
      }

      let changed = true;
      while (changed) {
        changed = false;
        newCells.forEach((cell) => {
          if (cell.formula) {
            const prevComputed = cell.computed;

            const computed = evaluateFormula(cell.formula, newCells);
            if (computed !== prevComputed) {
              cell.computed = computed;
              changed = true;
            }
          }
        });
      }

      return newCells;
    });
  }, []);

  const handleCellClick = useCallback((row: number, col: number) => {
    setSelectedCell({ row, col });
    setEditingCell(null);
  }, []);

  const handleCellDoubleClick = useCallback(
    (row: number, col: number) => {
      const cellData = getCellData(row, col);
      setEditingCell({ row, col });
      setEditValue(cellData?.value || "");
      setTimeout(() => inputRef.current?.focus(), 0);
    },
    [getCellData]
  );

  const handleEditConfirm = useCallback(() => {
    if (editingCell) {
      updateCell(editingCell.row, editingCell.col, editValue);
      setEditingCell(null);
      setEditValue("");
    }
  }, [editingCell, editValue, updateCell]);

  const handleScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
    const target = e.target as HTMLDivElement;
    setScrollTop(target.scrollTop);
    setScrollLeft(target.scrollLeft);
  }, []);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (editingCell) {
        if (e.key === "Enter") {
          e.preventDefault();
          handleEditConfirm();
        } else if (e.key === "Escape") {
          setEditingCell(null);
          setEditValue("");
        }
        return;
      }
    };

    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [editingCell, selectedCell, handleEditConfirm, handleCellDoubleClick]);

  const renderCell = useCallback(
    (row: number, col: number) => {
      const cellData = getCellData(row, col);
      const isSelected = selectedCell.row === row && selectedCell.col === col;
      const isEditing = editingCell?.row === row && editingCell?.col === col;

      const displayValue =
        cellData?.computed !== undefined
          ? cellData.computed.toString()
          : cellData?.value || "";

      return (
        <div
          key={`${row}-${col}`}
          className={cn(
            "absolute border border-gray-300 bg-white flex items-center px-2 text-sm",
            "hover:bg-blue-50 cursor-pointer",
            isSelected && "ring-2 ring-blue-500 bg-blue-50",
            isEditing && "ring-2 ring-green-500"
          )}
          style={{
            left: col * CELL_WIDTH,
            top: row * CELL_HEIGHT,
            width: CELL_WIDTH,
            height: CELL_HEIGHT,
          }}
          onClick={() => handleCellClick(row, col)}
          onDoubleClick={() => handleCellDoubleClick(row, col)}
        >
          {isEditing ? (
            <input
              ref={inputRef}
              value={editValue}
              onChange={(e) => setEditValue(e.target.value)}
              onBlur={handleEditConfirm}
              className="w-full h-full border-none outline-none bg-transparent"
            />
          ) : (
            <span className="truncate w-full">{displayValue}</span>
          )}
        </div>
      );
    },
    [
      selectedCell,
      editingCell,
      editValue,
      getCellData,
      handleCellClick,
      handleCellDoubleClick,
      handleEditConfirm,
    ]
  );

  return (
    <div className="h-screen flex flex-col bg-gray-100">
      <div className="h-12 bg-white border-b flex items-center px-4 gap-4">
        <div className="font-mono text-sm">
          {getColumnLabel(selectedCell.col)}
          {selectedCell.row + 1}
        </div>
        <input
          className="flex-1 px-2 py-1 border rounded text-sm"
          value={getCellData(selectedCell.row, selectedCell.col)?.value || ""}
          onChange={(e) =>
            updateCell(selectedCell.row, selectedCell.col, e.target.value)
          }
          placeholder="Enter value or formula (=A1+B1)"
        />
      </div>

      <div className="flex-1 relative overflow-hidden">
        <div
          className="absolute top-0 left-0 z-30 bg-gray-300 border-r border-b border-gray-400"
          style={{ width: ROW_HEADER_WIDTH, height: HEADER_HEIGHT }}
        ></div>

        <div
          className="absolute top-0 z-20 bg-gray-200 border-b border-gray-400 overflow-hidden"
          style={{
            left: ROW_HEADER_WIDTH,
            right: 0,
            height: HEADER_HEIGHT,
          }}
        >
          <div
            className="flex"
            style={{
              transform: `translateX(-${scrollLeft}px)`,
              width: TOTAL_COLS * CELL_WIDTH,
            }}
          >
            {Array.from({ length: TOTAL_COLS }, (_, col) => (
              <div
                key={col}
                className="border-r border-gray-300 bg-gray-200 flex items-center justify-center text-xs font-medium flex-shrink-0"
                style={{ width: CELL_WIDTH, height: HEADER_HEIGHT }}
              >
                {getColumnLabel(col)}
              </div>
            ))}
          </div>
        </div>

        <div
          className="absolute left-0 z-10 bg-gray-200 border-r border-gray-400 overflow-hidden"
          style={{
            top: HEADER_HEIGHT,
            bottom: 0,
            width: ROW_HEADER_WIDTH,
          }}
        >
          <div
            style={{
              transform: `translateY(-${scrollTop}px)`,
              height: TOTAL_ROWS * CELL_HEIGHT,
            }}
          >
            {Array.from({ length: TOTAL_ROWS }, (_, row) => (
              <div
                key={row}
                className="border-b border-gray-300 bg-gray-200 flex items-center justify-center text-xs font-medium"
                style={{ width: ROW_HEADER_WIDTH, height: CELL_HEIGHT }}
              >
                {row + 1}
              </div>
            ))}
          </div>
        </div>

        <div
          ref={containerRef}
          className="absolute inset-0 overflow-auto"
          style={{ paddingTop: HEADER_HEIGHT, paddingLeft: ROW_HEADER_WIDTH }}
          onScroll={handleScroll}
        >
          <div
            className="relative"
            style={{
              width: TOTAL_COLS * CELL_WIDTH,
              height: TOTAL_ROWS * CELL_HEIGHT,
            }}
          >
            {Array.from(
              { length: visibleRange.endRow - visibleRange.startRow },
              (_, i) => {
                const row = visibleRange.startRow + i;
                return Array.from(
                  { length: visibleRange.endCol - visibleRange.startCol },
                  (_, j) => {
                    const col = visibleRange.startCol + j;
                    return renderCell(row, col);
                  }
                );
              }
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default VirtualizedExcelSheet;
