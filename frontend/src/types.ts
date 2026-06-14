export type RateCardColumn = {
  name: string;
  strict: boolean;
  alias: string | null;
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
