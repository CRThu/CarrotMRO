export type CellRenderer = 'input' | 'select';

export type RateCardColumn = {
  name: string;
  strict: boolean;
  alias: string | null;
  cellRenderer?: CellRenderer;
  options?: string[];
  computed?: boolean;
};

export type PresetColumn = {
  key: string;
  label: string;
  required: boolean;
  type: string;
  computed?: boolean;
};

export type ColumnMapping = Record<string, string>; // template_col → preset_label

export type MappingScope = 'ocr' | 'ratecard' | 'quotation';

export type ColumnMappings = Record<MappingScope, ColumnMapping>;

export type TableItem = Record<string, string>;

export type QuotationItem = TableItem & {
  _matchStatus?: 'pending' | 'matched' | 'custom';
  '清单名称'?: string;
};

export type OcrTableData = {
  columns: RateCardColumn[];
  items: TableItem[];
  remarks: string;
};

export type RateCardTableData = {
  columns: RateCardColumn[];
  items: TableItem[];
};
