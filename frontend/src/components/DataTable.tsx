import { RateCardColumn, TableItem } from '@/types';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';

interface DataTableProps {
  columns: RateCardColumn[];
  items: TableItem[];
  onEdit: (index: number, field: string, value: string) => void;
}

export const DataTable = ({ columns, items, onEdit }: DataTableProps) => {
  return (
    <Table>
      <TableHeader>
        <TableRow>
          {columns.map((col) => (
            <TableHead key={col.name}>{col.name}</TableHead>
          ))}
        </TableRow>
      </TableHeader>
      <TableBody>
        {items.map((item, i) => (
          <TableRow key={i}>
            {columns.map((col) => (
              <TableCell key={col.name}>
                <Input
                  value={item[col.name] ?? ''}
                  onChange={(e) => onEdit(i, col.name, e.target.value)}
                />
              </TableCell>
            ))}
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
};
