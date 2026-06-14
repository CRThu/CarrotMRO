export type CellRenderer = 'input' | 'select';

export type RateCardColumn = {
  name: string;
  strict: boolean;
  alias: string | null;
  cellRenderer?: CellRenderer;
  options?: string[];
  computed?: boolean;
};

export type TableItem = Record<string, string>;

export type OcrTableData = {
  columns: RateCardColumn[];
  items: TableItem[];
  remarks: string;
};

export type RateCardTableData = {
  columns: RateCardColumn[];
  items: TableItem[];
};
