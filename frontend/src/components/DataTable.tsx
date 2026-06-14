import { TableItem } from '@/types';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';

interface DataTableProps {
  items: TableItem[];
  onEdit: (index: number, field: keyof TableItem, value: string) => void;
}

export const DataTable = ({ items, onEdit }: DataTableProps) => {
  return (
    <Table>
      <TableHeader>
        <TableRow>
          <TableHead>项目</TableHead>
          <TableHead>数量</TableHead>
          <TableHead>单位</TableHead>
          <TableHead>单价</TableHead>
        </TableRow>
      </TableHeader>
      <TableBody>
        {items.map((item, i) => (
          <TableRow key={i}>
            <TableCell>
              <Input value={item.name} onChange={(e) => onEdit(i, 'name', e.target.value)} />
            </TableCell>
            <TableCell>
              <Input value={item.quantity} onChange={(e) => onEdit(i, 'quantity', e.target.value)} />
            </TableCell>
            <TableCell>
              <Input value={item.unit} onChange={(e) => onEdit(i, 'unit', e.target.value)} />
            </TableCell>
            <TableCell>
              <Input value={item.unit_price || ''} onChange={(e) => onEdit(i, 'unit_price', e.target.value)} />
            </TableCell>
          </TableRow>
        ))}
      </TableBody>
    </Table>
  );
};
