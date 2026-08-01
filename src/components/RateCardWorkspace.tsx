import { useState, useRef } from 'react';
import { RateCardTableData, PRESET_COLUMNS } from '@/types';
import { DataTable } from '@/components/DataTable';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';
import { TaskNotification } from '@/components/TaskNotification';
import { Upload, ArrowRight, Check, AlertTriangle, CheckCircle2, Sparkles, HelpCircle } from 'lucide-react';
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

  // 通知状态
  const [notification, setNotification] = useState<{
    status: 'processing' | 'done' | 'error';
    message?: string;
    progress?: string;
  } | null>(null);

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
    setNotification({ status: 'processing', progress: `正在解析文件 ${file.name}...` });
    try {
      const res = await api.previewRateCardImport(currentRateCard, formData);
      const headers = res.data?.headers || [];
      const sampleRows = res.data?.sampleRows || [];
      const allRows = res.data?.allRows || sampleRows || [];

      // 仅当 Excel 表头与 10 项内置列名 100% 精确全等相同时才自动勾选，无则留空供手动匹配
      const initialMapping: Record<string, string> = {};
      headers.forEach(h => {
        const cleanHeader = h.trim();
        const matched = PRESET_COLUMNS.find(p => p === cleanHeader);
        initialMapping[h] = matched || '';
      });

      setPreviewData({ headers, sampleRows, allRows });
      setColumnMapping(initialMapping);
      setNotification(null);
    } catch (err: any) {
      const errMsg = err.response?.data?.detail || err.message || '解析 Excel/CSV 文件失败';
      setNotification({ status: 'error', message: errMsg });
    } finally {
      setLoading(false);
      e.target.value = '';
    }
  };

  const handleConfirmImport = async () => {
    if (!previewData) return;

    // 校验是否选择了有效的系统内置列映射
    const hasValidMapping = Object.values(columnMapping).some(v => Boolean(v) && PRESET_COLUMNS.includes(v as any));
    if (!hasValidMapping) {
      setNotification({
        status: 'error',
        message: '请至少将一个 Excel 列映射到系统内置标准列（如「项目名称」或「不含税单价」）！',
      });
      return;
    }

    setLoading(true);
    setNotification({ status: 'processing', progress: '正在清洗并保存协议定价表...' });
    try {
      const safeItems = previewData.allRows || previewData.sampleRows || [];
      const res = await api.importRateCardFile(currentRateCard, {
        headers: previewData.headers || [],
        items: safeItems,
        mapping: columnMapping,
      });
      const count = res.data?.count ?? safeItems.length;
      setPreviewData(null);
      await onRefreshData();
      setNotification({
        status: 'done',
        progress: `协议定价表「${currentRateCard}」导入成功，共清洗导入 ${count} 条规范数据！`,
      });
    } catch (err: any) {
      const errMsg = err.response?.data?.detail || err.message || '导入定价表失败';
      setNotification({ status: 'error', message: errMsg });
    } finally {
      setLoading(false);
    }
  };

  // 校验关键字段映射状态
  const mappedCols = Object.values(columnMapping).filter(v => Boolean(v) && PRESET_COLUMNS.includes(v as any));
  const hasItemNameMapped = mappedCols.includes('项目名称');
  const hasPriceExclMapped = mappedCols.includes('不含税单价');

  return (
    <div className="space-y-6 relative">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-light text-gray-800">协议定价表: {currentRateCard}</h1>
          <p className="text-sm text-gray-500 mt-1">协议基准单价库，供报价单比对填充使用</p>
        </div>
        <div className="flex items-center gap-3">
          <input
            ref={fileInputRef}
            type="file"
            data-testid="file-input"
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

      {/* 顶部元数据与已导入列状态 Banner */}
      <Card className="shadow-sm border-gray-200 bg-gray-50/50">
        <CardContent className="py-3 px-4 flex items-center justify-between">
          <div className="flex items-center gap-4 text-xs text-gray-600">
            <span>数据容量: <strong className="text-gray-900 font-mono">{ratecardTableData?.items?.length ?? 0}</strong> 条记录</span>
            <span className="text-gray-300">|</span>
            <span>包含标准列: <strong className="text-gray-900">{(ratecardTableData?.columns ?? []).join(', ') || '暂无数据'}</strong></span>
          </div>
          {ratecardTableData?.columns && !ratecardTableData.columns.includes('项目名称') && ratecardTableData.items.length > 0 && (
            <div className="flex items-center gap-1.5 text-xs text-amber-700 bg-amber-50 border border-amber-200 px-2.5 py-1 rounded-md">
              <AlertTriangle className="h-3.5 w-3.5 text-amber-500 shrink-0" />
              <span>当前定价表缺少「项目名称」列，可能影响报价单比对效率</span>
            </div>
          )}
        </CardContent>
      </Card>

      <DataTable
        height="calc(100vh - 260px)"
        columns={ratecardTableData?.columns ?? []}
        items={ratecardTableData?.items ?? []}
        showRowNumber={true}
        emptyText="协议定价表无数据，请导入 Excel / CSV"
      />

      {/* 导入模态框：精确全等匹配 + 丰富示例数据 + 实时手动调整与校验 Banner */}
      {previewData && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-4xl w-full p-6 max-h-[90vh] flex flex-col">
            {/* Modal Header */}
            <div className="flex items-center justify-between border-b pb-4">
              <div>
                <h2 className="text-xl font-medium text-gray-800 flex items-center gap-2">
                  <span>导入定价表 - 列映射配置</span>
                  <span className="text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded-full font-normal">
                    手动确认模式
                  </span>
                </h2>
                <p className="text-xs text-gray-500 mt-1">
                  只自动预选 100% 精确同名列。请核对下方原始列内容并手动选择对应标准列
                </p>
              </div>
              <div className="flex flex-col items-end gap-1">
                <span className="text-xs font-mono bg-gray-100 text-gray-700 px-2.5 py-1 rounded-md border">
                  Excel 共 {previewData.allRows.length} 行数据
                </span>
                <span className="text-xs text-gray-500">
                  已映射: <strong className="text-blue-600">{mappedCols.length}</strong> / {previewData.headers.length} 列
                </span>
              </div>
            </div>

            {/* Modal Content - List of Columns with Sample Data */}
            <div className="flex-1 overflow-y-auto py-4 space-y-3">
              {previewData.headers.map(h => {
                const sampleVal = previewData.sampleRows[0]?.[h] ?? '';
                const currentMapped = columnMapping[h] || '';
                const isExactMatch = PRESET_COLUMNS.includes(h.trim() as any);

                return (
                  <div key={h} className="p-3.5 bg-gray-50/80 rounded-xl border border-gray-200/90 hover:border-blue-200 transition-colors">
                    <div className="flex items-center gap-4">
                      {/* 原始列名称与匹配 Status Badge */}
                      <div className="w-5/12 space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="text-sm font-semibold text-gray-800 truncate" title={h}>
                            {h}
                          </span>
                          {isExactMatch ? (
                            <span className="inline-flex items-center gap-1 text-[11px] font-medium bg-emerald-100 text-emerald-800 px-2 py-0.5 rounded-full border border-emerald-200 shrink-0">
                              <CheckCircle2 className="h-3 w-3" />
                              精准匹配
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 text-[11px] font-medium bg-gray-200/70 text-gray-600 px-2 py-0.5 rounded-full shrink-0">
                              未匹配
                            </span>
                          )}
                        </div>
                        {/* 第一条示例数据 Showcase */}
                        <div className="text-xs text-gray-500 bg-white/90 px-2 py-1 rounded border border-gray-200/70 truncate font-mono" title={sampleVal}>
                          <span className="text-gray-400 font-sans mr-1">示例:</span>
                          {sampleVal ? `"${sampleVal}"` : <span className="italic text-gray-400">(空值)</span>}
                        </div>
                      </div>

                      <ArrowRight className="h-4 w-4 text-gray-400 shrink-0" />

                      {/* 标准列选择下拉框 */}
                      <div className="flex-1">
                        <select
                          value={currentMapped}
                          onChange={(e) => {
                            const val = e.target.value;
                            setColumnMapping(prev => ({ ...prev, [h]: val }));
                          }}
                          className={`w-full p-2 border rounded-lg text-sm bg-white focus:ring-2 focus:ring-blue-500 focus:outline-none transition-all ${
                            currentMapped
                              ? 'border-blue-300 font-medium text-blue-900 bg-blue-50/30'
                              : 'border-gray-300 text-gray-500'
                          }`}
                        >
                          <option value="">(不匹配 / 忽略此列)</option>
                          {PRESET_COLUMNS.map(col => (
                            <option key={col} value={col}>
                              系统标准列 ➔ {col}
                            </option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>

            {/* 智能警示与指引 Footer Banner */}
            {(!hasItemNameMapped || !hasPriceExclMapped) && (
              <div className="mb-3 p-3 bg-amber-50 border border-amber-200 rounded-lg flex items-center gap-2 text-xs text-amber-800">
                <AlertTriangle className="h-4 w-4 text-amber-600 shrink-0" />
                <div>
                  {!hasItemNameMapped && <span>⚠️ 尚未选择映射到 <strong>「项目名称」</strong> ；</span>}
                  {!hasPriceExclMapped && <span>⚠️ 尚未选择映射到 <strong>「不含税单价」</strong> ；</span>}
                  <span className="text-amber-700 ml-1">建议手动指定对应的 Excel 列，以便报价单能成功比对单价。</span>
                </div>
              </div>
            )}

            {/* Footer Actions */}
            <div className="border-t pt-4 flex items-center justify-between">
              <span className="text-xs text-gray-500 flex items-center gap-1">
                <Sparkles className="h-3.5 w-3.5 text-blue-500" />
                合并单元格组名已自动继承解包填充
              </span>
              <div className="flex items-center gap-3">
                <Button variant="outline" onClick={() => setPreviewData(null)} disabled={loading}>
                  取消
                </Button>
                <Button onClick={handleConfirmImport} disabled={loading} className="bg-blue-600 hover:bg-blue-700 text-white shadow-sm">
                  <Check className="h-4 w-4 mr-1.5" />
                  确认映射导入
                </Button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* 底部 Task Notification 消息提醒 */}
      <TaskNotification
        status={notification}
        autoDismissMs={5000}
        onDismiss={() => setNotification(null)}
      />
    </div>
  );
}

