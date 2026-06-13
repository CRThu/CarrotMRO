import { TableItem } from '../types';

interface DataTableProps {
  items: TableItem[];
  onEdit: (index: number, field: keyof TableItem, value: string) => void;
}

export const DataTable = ({ items, onEdit }: DataTableProps) => {
  return (
    <table className="w-full border-collapse mt-4">
      <thead>
        <tr className="border-b-2 border-gray-200">
          <th className="p-3 text-left">项目</th>
          <th className="p-3 text-left">数量</th>
          <th className="p-3 text-left">单位</th>
          <th className="p-3 text-left">单价</th>
        </tr>
      </thead>
      <tbody>
        {items.map((item, i) => (
          <tr key={i} className="border-b border-gray-100">
            <td className="p-2">
              <input value={item.name} onChange={(e) => onEdit(i, 'name', e.target.value)} className="w-full p-2 border rounded" />
            </td>
            <td className="p-2">
              <input value={item.quantity} onChange={(e) => onEdit(i, 'quantity', e.target.value)} className="w-full p-2 border rounded" />
            </td>
            <td className="p-2">
              <input value={item.unit} onChange={(e) => onEdit(i, 'unit', e.target.value)} className="w-full p-2 border rounded" />
            </td>
            <td className="p-2">
              <input value={item.unit_price || ''} onChange={(e) => onEdit(i, 'unit_price', e.target.value)} className="w-full p-2 border rounded" />
            </td>
          </tr>
        ))}
      </tbody>
    </table>
  );
};
