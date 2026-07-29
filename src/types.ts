// 系统全局 10 项标准预制列
export const PRESET_COLUMNS = [
  '项目组',
  '项目名称',
  '单位',
  '数量',
  '不含税单价',
  '不含税总价',
  '税率',
  '含税单价',
  '含税总价',
  '说明',
] as const;

export type PresetColumnName = (typeof PRESET_COLUMNS)[number];

export interface ProjectSettings {
  name: string;
  created_at: string;
  ratecard_name: string | null;
  template_name: string | null;
  ocr_columns: string[];
  quotation_columns: string[];
}

export type TableItem = Record<string, string>;

export type QuotationItem = TableItem & {
  _matchStatus?: 'pending' | 'matched' | 'custom';
  '清单名称'?: string;
};

export interface QuotationData {
  created_at?: string;
  last_edit_time?: string;
  items: QuotationItem[];
  remarks?: string[];
}

export type RateCardTableData = {
  columns: string[];
  items: TableItem[];
};

export interface ProviderConfig {
  apiKey: string;
  model: string;
  baseUrl: string;
  proxy?: string;
}

export interface LlmConfig {
  activeProvider: string;
  providers: Record<string, ProviderConfig>;
}

export interface AppSettings {
  llm: LlmConfig;
}
