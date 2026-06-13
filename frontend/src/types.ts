export type TableItem = {
  name: string;
  quantity: string;
  unit: string;
  unit_price: string;
};

export type TableData = {
  items: TableItem[];
  remarks: string;
};
