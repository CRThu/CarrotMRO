import { RateCardColumn, TableItem } from '@/types';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';
import { Button } from '@/components/ui/button';
import { Trash2, Plus } from 'lucide-react';

interface DataTableProps {
  columns: RateCardColumn[];
  items: TableItem[];
  onEdit?: (index: number, field: string, value: string) => void;
  onAddRow?: (index?: number) => void;
  onDeleteRow?: (index: number) => void;
}

const renderCell = (item: TableItem, col: RateCardColumn, index: number, onEdit?: DataTableProps['onEdit']) => {
  const key = col.alias || col.name;
  const value = item[key] ?? '';

  if (col.cellRenderer === 'select' && col.options) {
    return onEdit ? (
      <select
        value={value}
        onChange={(e) => onEdit(index, key, e.target.value)}
        className="w-full p-1 border rounded text-sm"
      >
        <option value="">-</option>
        {col.options.map((opt) => (
          <option key={opt} value={opt}>{opt}</option>
        ))}
      </select>
    ) : (
      <span className="text-sm">{value || '-'}</span>
    );
  }

  return onEdit ? (
    <Input
      value={value}
      onChange={(e) => onEdit(index, key, e.target.value)}
    />
  ) : (
    <span>{value}</span>
  );
};

export const DataTable = ({ columns, items, onEdit, onAddRow, onDeleteRow }: DataTableProps) => {
  const hasRowActions = onAddRow || onDeleteRow;

  return (
    <div>
      <Table>
        <TableHeader>
          <TableRow>
            {columns.map((col) => (
              <TableHead key={col.name}>{col.name}</TableHead>
            ))}
            {hasRowActions && <TableHead className="w-20"></TableHead>}
          </TableRow>
        </TableHeader>
        <TableBody>
          {items.map((item, i) => (
            <TableRow key={i}>
              {columns.map((col) => (
                <TableCell key={col.name}>
                  {renderCell(item, col, i, onEdit)}
                </TableCell>
              ))}
              {hasRowActions && (
                <TableCell className="whitespace-nowrap">
                  <div className="flex items-center gap-0.5">
                    {onAddRow && (
                      <Button
                        variant="ghost"
                        size="icon"
                        onClick={() => onAddRow(i)}
                        className="h-7 w-7 text-gray-400 hover:text-green-600"
                        title="在下方插入行"
                      >
                        <Plus className="h-4 w-4" />
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
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    )}
                  </div>
                </TableCell>
              )}
            </TableRow>
          ))}
        </TableBody>
      </Table>
      {onAddRow && (
        <Button
          variant="outline"
          size="sm"
          onClick={() => onAddRow()}
          className="mt-3 text-gray-500"
        >
          <Plus className="h-4 w-4 mr-1" />
          末尾新增行
        </Button>
      )}
    </div>
  );
};
