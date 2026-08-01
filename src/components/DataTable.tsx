import React, { useState, useCallback } from 'react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { Trash2, Plus } from 'lucide-react';

// 列定义接口支持字符串或高级配置
export type ColumnDef = string | {
  name: string;
  computed?: boolean;
  cellRenderer?: string;
  options?: string[];
  width?: number;
};

export interface DataTableProps {
  columns: ColumnDef[];
  items: Record<string, string>[];
  height?: string; // 控件高度容器管理，例如 "calc(100vh - 270px)"，默认 "100%"
  onEdit?: (index: number, field: string, value: string) => void;
  onAddRow?: (index?: number) => void;
  onDeleteRow?: (index: number) => void;

  // 高级自定义插槽与展示属性
  showRowNumber?: boolean; // 是否在最左侧呈现 # 序号列
  rowNumberStart?: number; // 序号起始数值，默认 1
  renderPrefixHeader?: () => React.ReactNode; // 最左侧表头自定义插槽（如 "匹配" 列表头）
  renderPrefixCell?: (item: Record<string, string>, index: number) => React.ReactNode; // 最左侧单元格自定义插槽（如 "匹配" Popover 按钮）
  initialColumnWidths?: Record<string, number>; // 初始列宽配置
  onColumnWidthsChange?: (widths: Record<string, number>) => void;
  emptyText?: string;
  addRowText?: string;
  addRowPosition?: 'top' | 'bottom' | 'both'; // 新增行按钮位置，默认 top 让横向滚动条紧贴控件最底层边框
}

// 数字类常用标准字段（自动应用右对齐与 monospace 字体）
const NUMERIC_COLUMNS = ['数量', '不含税单价', '不含税总价', '税率', '含税单价', '含税总价'];

// 默认列宽表预设
const DEFAULT_WIDTHS: Record<string, number> = {
  '匹配': 46,
  '#': 40,
  '项目组': 110,
  '项目名称': 240,
  '单位': 80,
  '数量': 90,
  '不含税单价': 110,
  '不含税总价': 120,
  '税率': 80,
  '含税单价': 110,
  '含税总价': 120,
  '说明': 220,
};

const renderCellContent = (
  item: Record<string, string>,
  col: ColumnDef,
  index: number,
  onEdit?: DataTableProps['onEdit']
) => {
  const colName = typeof col === 'string' ? col : col.name;
  const value = item[colName] ?? '';

  if (typeof col !== 'string' && col.computed) {
    return <span className="text-sm font-medium text-gray-700 px-2 py-1 block">{value || '-'}</span>;
  }

  if (typeof col !== 'string' && col.cellRenderer === 'select' && col.options) {
    return onEdit ? (
      <select
        value={value}
        onChange={(e) => onEdit(index, colName, e.target.value)}
        className="w-full h-9 p-1 border rounded text-sm bg-white"
      >
        <option value="">-</option>
        {col.options.map((opt) => (
          <option key={opt} value={opt}>{opt}</option>
        ))}
      </select>
    ) : (
      <span className="text-sm px-2 py-1 block">{value || '-'}</span>
    );
  }

  const isNumeric = NUMERIC_COLUMNS.includes(colName);

  return onEdit ? (
    <input
      type="text"
      value={value}
      onChange={(e) => onEdit(index, colName, e.target.value)}
      className={`w-full h-9 px-2 py-1 text-sm bg-transparent border-2 border-transparent hover:bg-blue-50/50 focus:bg-white focus:border-blue-500 focus:ring-2 focus:ring-blue-100 rounded-none text-gray-800 transition-all outline-none ${
        isNumeric ? 'text-right font-mono' : 'text-left font-sans'
      }`}
    />
  ) : (
    <span className={`text-sm px-2 py-1 truncate block ${isNumeric ? 'text-right font-mono' : 'text-left font-sans'}`}>
      {value}
    </span>
  );
};

export const DataTable: React.FC<DataTableProps> = ({
  columns,
  items,
  height,
  onEdit,
  onAddRow,
  onDeleteRow,
  showRowNumber = false,
  rowNumberStart = 1,
  renderPrefixHeader,
  renderPrefixCell,
  initialColumnWidths,
  onColumnWidthsChange,
  emptyText = '暂无数据',
  addRowText = '新增数据行',
  addRowPosition = 'top',
}) => {
  const hasRowActions = Boolean(onAddRow || onDeleteRow);

  // 初始化列宽 Map
  const [columnWidths, setColumnWidths] = useState<Record<string, number>>(() => {
    const initial: Record<string, number> = { ...DEFAULT_WIDTHS, ...(initialColumnWidths || {}) };
    columns.forEach((c) => {
      const colName = typeof c === 'string' ? c : c.name;
      if (!initial[colName]) {
        initial[colName] = typeof c !== 'string' && c.width ? c.width : 120;
      }
    });
    return initial;
  });

  const updateColumnWidth = useCallback((colName: string, newWidth: number) => {
    setColumnWidths((prev) => {
      const next = { ...prev, [colName]: newWidth };
      onColumnWidthsChange?.(next);
      return next;
    });
  }, [onColumnWidthsChange]);

  // 0ms 原生 拖拽 Resize Handle
  const handleResizeStart = useCallback((colName: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    const startX = e.clientX;
    const startWidth = columnWidths[colName] || 120;

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const diff = moveEvent.clientX - startX;
      const newW = Math.max(45, startWidth + diff);
      updateColumnWidth(colName, newW);
    };

    const handleMouseUp = () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
  }, [columnWidths, updateColumnWidth]);

  // 双击分隔线 Auto-Fit 计算最大字符宽
  const handleDoubleClickAutoFit = useCallback((colName: string, e: React.MouseEvent) => {
    e.preventDefault();
    e.stopPropagation();
    let maxCharLen = colName.length;

    items.forEach((item) => {
      const val = item[colName] ?? '';
      if (val) {
        // 计算字符权重（中文字符占 2 宽）
        let len = 0;
        for (let i = 0; i < val.length; i++) {
          len += val.charCodeAt(i) > 255 ? 2 : 1;
        }
        if (len > maxCharLen) {
          maxCharLen = len;
        }
      }
    });

    // 计算宽度：字符长度 * 8.5px + 28px padding
    const autoWidth = Math.min(500, Math.max(60, Math.ceil(maxCharLen * 8.5) + 28));
    updateColumnWidth(colName, autoWidth);
  }, [items, updateColumnWidth]);

  const showTopAddBtn = onAddRow && (addRowPosition === 'top' || addRowPosition === 'both');
  const showBottomAddBtn = onAddRow && (addRowPosition === 'bottom' || addRowPosition === 'both');

  return (
    <div
      className="relative flex flex-col w-full border border-slate-200/90 rounded-xl overflow-hidden shadow-xs bg-white"
      style={{ height: height || '100%' }}
    >
      {/* 顶部控件 Header 工具栏（支持顶部新增行与行数统统计） */}
      {(showTopAddBtn || items.length > 0) && (
        <div className="p-2 px-3 border-b border-slate-200/80 bg-slate-50/70 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2">
            {showTopAddBtn && (
              <Button
                variant="outline"
                size="sm"
                onClick={() => onAddRow()}
                className="h-7 text-xs text-slate-700 bg-white border-slate-300 shadow-2xs hover:bg-slate-50 hover:text-blue-600"
              >
                <Plus className="h-3.5 w-3.5 mr-1 text-blue-600" />
                {addRowText}
              </Button>
            )}
          </div>
          <div className="text-xs text-slate-400 font-mono">
            共 <strong className="text-slate-700 font-sans">{items.length}</strong> 条记录
          </div>
        </div>
      )}

      {/* 内部 XY 双向独立滚动容器：最下端与控件底层边框完美接合 */}
      <div className="flex-1 min-h-0 min-w-0 overflow-auto">
        <Table className="border-collapse border-spacing-0 w-max min-w-full">
          <TableHeader className="sticky top-0 z-10 bg-slate-100/95 backdrop-blur-xs shadow-xs border-b border-slate-200">
            <TableRow className="hover:bg-transparent">
              {/* 可选最左侧前缀表头（如 匹配 列） */}
              {renderPrefixHeader && (
                <TableHead
                  data-col="匹配"
                  style={{ width: columnWidths['匹配'] || 46, minWidth: columnWidths['匹配'] || 46 }}
                  className="text-center font-medium text-slate-700 select-none py-2 px-1 border-r border-slate-200/80 bg-slate-100"
                >
                  {renderPrefixHeader()}
                </TableHead>
              )}

              {/* 可选序号列 # */}
              {showRowNumber && (
                <TableHead
                  data-col="#"
                  style={{ width: columnWidths['#'] || 40, minWidth: columnWidths['#'] || 40 }}
                  className="font-medium text-slate-700 select-none py-2 px-1 border-r border-slate-200/80 bg-slate-100 text-center"
                >
                  #
                </TableHead>
              )}

              {/* 主列 Header */}
              {columns.map((col) => {
                const colName = typeof col === 'string' ? col : col.name;
                const width = columnWidths[colName] || 120;
                return (
                  <TableHead
                    key={colName}
                    data-col={colName}
                    style={{ width, minWidth: width }}
                    className="font-medium text-slate-700 whitespace-nowrap select-none py-2 px-2.5 relative group border-r border-slate-200/80 bg-slate-100"
                  >
                    <span className="truncate block">{colName}</span>
                    {/* 0ms 原生拖拽 Resize 分隔线与双击 Auto-Fit */}
                    <div
                      className="absolute right-0 top-0 bottom-0 w-2 cursor-col-resize hover:bg-blue-500 active:bg-blue-600 transition-colors z-20 opacity-0 group-hover:opacity-100"
                      onMouseDown={(e) => handleResizeStart(colName, e)}
                      onDoubleClick={(e) => handleDoubleClickAutoFit(colName, e)}
                      title="拖拽调节列宽，双击自适应最佳列宽"
                    />
                  </TableHead>
                );
              })}

              {/* 行操作列 Header */}
              {hasRowActions && (
                <TableHead className="w-14 bg-slate-100 text-center font-medium text-slate-700 py-2 px-1 border-slate-200/80" />
              )}
            </TableRow>
          </TableHeader>

          <TableBody>
            {items.map((item, i) => (
              <TableRow key={i} className="hover:bg-slate-50/70 border-b border-slate-200/60">
                {/* 最左侧单元格插槽 */}
                {renderPrefixCell && (
                  <TableCell
                    data-col="匹配"
                    style={{ width: columnWidths['匹配'] || 46, minWidth: columnWidths['匹配'] || 46 }}
                    className="text-center p-1 border-r border-slate-200/50"
                  >
                    {renderPrefixCell(item, i)}
                  </TableCell>
                )}

                {/* 序号列 # */}
                {showRowNumber && (
                  <TableCell
                    data-col="#"
                    style={{ width: columnWidths['#'] || 40, minWidth: columnWidths['#'] || 40 }}
                    className="text-xs text-slate-400 font-mono p-1 text-center border-r border-slate-200/50"
                  >
                    {i + rowNumberStart}
                  </TableCell>
                )}

                {/* 各主列数据单元格 */}
                {columns.map((col) => {
                  const colName = typeof col === 'string' ? col : col.name;
                  const width = columnWidths[colName] || 120;
                  return (
                    <TableCell
                      key={colName}
                      data-col={colName}
                      style={{ width, minWidth: width }}
                      className="p-0 border-r border-slate-200/50 last:border-r-0"
                    >
                      {renderCellContent(item, col, i, onEdit)}
                    </TableCell>
                  );
                })}

                {/* 行快捷操作按钮（添加/删除） */}
                {hasRowActions && (
                  <TableCell className="whitespace-nowrap p-1.5 text-center">
                    <div className="flex items-center justify-center gap-0.5">
                      {onAddRow && (
                        <Button
                          variant="ghost"
                          size="icon"
                          onClick={() => onAddRow(i)}
                          className="h-7 w-7 text-gray-400 hover:text-green-600"
                          title="在下方插入行"
                        >
                          <Plus className="h-3.5 w-3.5" />
                        </Button>
                      )}
                      {onDeleteRow && (
                        <Button
                          variant="ghost"
                          size="icon"
                          onClick={() => onDeleteRow(i)}
                          className="h-7 w-7 text-gray-400 hover:text-red-500"
                          title="删除此行"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </Button>
                      )}
                    </div>
                  </TableCell>
                )}
              </TableRow>
            ))}
          </TableBody>
        </Table>

        {items.length === 0 && (
          <div className="text-center py-10 text-gray-400 text-xs">
            {emptyText}
          </div>
        )}
      </div>

      {/* 底部操作工具条（仅当显式设置 addRowPosition 为 bottom 或 both 时展示） */}
      {showBottomAddBtn && (
        <div className="p-2.5 px-3 border-t border-slate-200/80 bg-slate-50/50 flex items-center justify-between shrink-0">
          <Button
            variant="outline"
            size="sm"
            onClick={() => onAddRow()}
            className="text-xs text-slate-600 border-dashed hover:bg-white"
          >
            <Plus className="h-3.5 w-3.5 mr-1" />
            {addRowText}
          </Button>
        </div>
      )}
    </div>
  );
};
