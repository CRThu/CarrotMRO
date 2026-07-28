import { useState, useRef } from 'react';
import { RateCardTableData, PRESET_COLUMNS } from '@/types';
import { DataTable } from '@/components/DataTable';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { Upload, ArrowRight, Check } from 'lucide-react';
import * as api from '@/api';

interface RateCardWorkspaceProps {
  currentRateCard: string;
  ratecardTableData: RateCardTableData;
  onRefreshData: () => Promise<void>;
}

export function RateCardWorkspace({
  currentRateCard,
  ratecardTableData,
  onRefreshData,
}: RateCardWorkspaceProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [loading, setLoading] = useState(false);

  // 导入预览对话框状态
  const [previewData, setPreviewData] = useState<{
    headers: string[];
    sampleRows: Record<string, string>[];
    allRows: Record<string, string>[];
  } | null>(null);

  // 表头映射 state: { [原始表头]: 目标标准列名 }
  const [columnMapping, setColumnMapping] = useState<Record<string, string>>({});

  const handleFileSelected = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    setLoading(true);
    try {
      const res = await api.previewRateCardImport(currentRateCard, formData);
      const { headers, sampleRows, allRows } = res.data;

      // 智能推理初始映射关系
      const initialMapping: Record<string, string> = {};
      headers.forEach(h => {
        const cleanHeader = h.trim().toLowerCase();
        const matchedPreset = PRESET_COLUMNS.find(p => {
          const cleanPreset = p.toLowerCase();
          return cleanHeader.includes(cleanPreset) || cleanPreset.includes(cleanHeader);
        });
        if (matchedPreset) {
          initialMapping[h] = matchedPreset;
        } else if (cleanHeader.includes('名') || cleanHeader.includes('品')) {
          initialMapping[h] = '项目名称';
        } else if (cleanHeader.includes('单价') || cleanHeader.includes('价格')) {
          initialMapping[h] = '不含税单价';
        } else if (cleanHeader.includes('规') || cleanHeader.includes('型') || cleanHeader.includes('备')) {
          initialMapping[h] = '说明';
        } else {
          initialMapping[h] = '';
        }
      });

      setPreviewData({ headers, sampleRows, allRows });
      setColumnMapping(initialMapping);
    } catch (err: any) {
      alert('解析 Excel/CSV 文件失败: ' + (err.response?.data?.detail || err.message));
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  };

  const handleConfirmImport = async () => {
    if (!previewData) return;
    setLoading(true);
    try {
      await api.importRateCardFile(currentRateCard, {
        headers: previewData.headers,
        items: previewData.allRows,
        mapping: columnMapping,
      });
      setPreviewData(null);
      await onRefreshData();
      alert('协议定价表导入成功！');
    } catch (err: any) {
      alert('导入定价表失败: ' + (err.response?.data?.detail || err.message));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-light text-gray-800">协议定价表: {currentRateCard}</h1>
          <p className="text-sm text-gray-500 mt-1">协议基准单价库，供报价单比对填充使用</p>
        </div>
        <div className="flex items-center gap-3">
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls,.csv"
            className="hidden"
            onChange={handleFileSelected}
          />
          <Button
            onClick={() => fileInputRef.current?.click()}
            disabled={loading}
            className="bg-blue-600 hover:bg-blue-700 text-white shadow-sm"
          >
            <Upload className="h-4 w-4 mr-1.5" />
            {loading ? '正在解析...' : '导入 Excel / CSV'}
          </Button>
        </div>
      </div>

      <Card className="shadow-sm">
        <CardContent className="pt-6">
          <DataTable
            columns={ratecardTableData?.columns ?? []}
            items={ratecardTableData?.items ?? []}
          />
        </CardContent>
      </Card>

      {/* 导入模态框：映射 Excel 表头到 10 项标准列 */}
      {previewData && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-3xl w-full p-6 max-h-[90vh] flex flex-col">
            <div className="flex items-center justify-between border-b pb-4">
              <div>
                <h2 className="text-xl font-medium text-gray-800">导入定价表 - 表头映射对齐</h2>
                <p className="text-xs text-gray-500 mt-0.5">请将 Excel 中的原始列名映射到系统 10 项标准预制列</p>
              </div>
              <span className="text-xs font-mono bg-blue-50 text-blue-700 px-2.5 py-1 rounded-full border border-blue-200">
                共 {previewData.allRows.length} 条数据
              </span>
            </div>

            <div className="flex-1 overflow-y-auto py-4 space-y-3">
              {previewData.headers.map(h => (
                <div key={h} className="flex items-center gap-3 p-3 bg-gray-50 rounded-lg border border-gray-200">
                  <div className="w-1/3 text-sm font-medium text-gray-700 truncate" title={h}>
                    {h}
                  </div>
                  <ArrowRight className="h-4 w-4 text-gray-400 shrink-0" />
                  <select
                    value={columnMapping[h] || ''}
                    onChange={(e) => setColumnMapping({ ...columnMapping, [h]: e.target.value })}
                    className="flex-1 p-2 border border-gray-300 rounded-lg text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  >
                    <option value="">(忽略此列 / 不导入)</option>
                    {PRESET_COLUMNS.map(col => (
                      <option key={col} value={col}>{col}</option>
                    ))}
                  </select>
                </div>
              ))}
            </div>

            <div className="border-t pt-4 flex items-center justify-between">
              <span className="text-xs text-gray-500">
                建议确保映射了「项目名称」与「不含税单价」或「含税单价」
              </span>
              <div className="flex items-center gap-3">
                <Button variant="outline" onClick={() => setPreviewData(null)} disabled={loading}>
                  取消
                </Button>
                <Button onClick={handleConfirmImport} disabled={loading} className="bg-blue-600 hover:bg-blue-700 text-white">
                  <Check className="h-4 w-4 mr-1.5" />
                  确认导入
                </Button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
