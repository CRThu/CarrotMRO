import { useState, useRef } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { MatchPopover } from '@/components/MatchPopover';
import { OcrProgressModal } from '@/components/OcrProgressModal';
import { QuotationItem } from '@/types';
import * as api from '@/api';
import { Save, Download, Plus, Trash2, Image, Loader2, Maximize2 } from 'lucide-react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';

interface MatchCandidate {
  name: string;
  score: number;
  columns: string[];
  values: string[];
  itemData?: Record<string, string>;
}

interface QuotationWorkspaceProps {
  currentProject: string;
  activeQuotationFilename: string;
  quotationItems: QuotationItem[];
  projectRateCard: string | null;
  projectTemplate: string | null;
  quotationColumns: string[];
  onEdit: (index: number, field: string, value: string) => void;
  onAddRow: (index?: number) => void;
  onDeleteRow: (index: number) => void;
  onSave: () => void;
  onQuotationDataChange: (items: QuotationItem[]) => void;
}

// 帮助函数：自动计算联动价格
export function calculateRowFormulas(item: Record<string, string>, changedField?: string): Record<string, string> {
  const updated = { ...item };
  const qty = parseFloat(updated['数量'] ?? '');
  const priceExcl = parseFloat(updated['不含税单价'] ?? '');
  let taxRateStr = updated['税率'] ?? '';
  let taxRate = parseFloat(taxRateStr);
  if (taxRateStr.includes('%')) {
    taxRate = parseFloat(taxRateStr.replace('%', '')) / 100;
  } else if (taxRate > 1) {
    taxRate = taxRate / 100;
  }

  // 1. 不含税总价
  if (!isNaN(qty) && !isNaN(priceExcl)) {
    updated['不含税总价'] = (qty * priceExcl).toFixed(2);
  }

  // 2. 税率存在时联动计算含税单价与含税总价
  if (!isNaN(priceExcl) && !isNaN(taxRate)) {
    const priceIncl = priceExcl * (1 + taxRate);
    updated['含税单价'] = priceIncl.toFixed(2);
    if (!isNaN(qty)) {
      updated['含税总价'] = (qty * priceIncl).toFixed(2);
    }
  }

  return updated;
}

export function QuotationWorkspace({
  currentProject,
  activeQuotationFilename,
  quotationItems,
  projectRateCard,
  projectTemplate,
  quotationColumns,
  onEdit,
  onAddRow,
  onDeleteRow,
  onSave,
  onQuotationDataChange,
}: QuotationWorkspaceProps) {
  const ocrFileInputRef = useRef<HTMLInputElement>(null);

  // 独立 OCR 多图识别状态日志弹窗 state
  const [activeTaskId, setActiveTaskId] = useState<string | null>(null);
  const [ocrModalOpen, setOcrModalOpen] = useState(false);
  const [ocrStatus, setOcrStatus] = useState<'processing' | 'done' | 'error'>('processing');
  const [ocrImageCount, setOcrImageCount] = useState(0);
  const [ocrCurrentStep, setOcrCurrentStep] = useState('');
  const [ocrLogs, setOcrLogs] = useState<string[]>([]);
  const [ocrErrorMessage, setOcrErrorMessage] = useState<string | undefined>(undefined);
  const [ocrExtractedCount, setOcrExtractedCount] = useState(0);

  const [matchCandidatesMap, setMatchCandidatesMap] = useState<Record<number, MatchCandidate[]>>({});
  const [matchLoadingMap, setMatchLoadingMap] = useState<Record<number, boolean>>({});

  // 自动保存状态: 'saved' | 'saving' | 'error'
  const [saveStatus, setSaveStatus] = useState<'saved' | 'saving' | 'error'>('saved');
  const autoSaveTimerRef = useRef<NodeJS.Timeout | null>(null);

  // 触发保存并更新状态
  const triggerAutoSave = (itemsToSave: QuotationItem[]) => {
    setSaveStatus('saving');
    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = setTimeout(async () => {
      try {
        await api.saveQuotationData(currentProject, activeQuotationFilename, { items: itemsToSave });
        setSaveStatus('saved');
      } catch {
        setSaveStatus('error');
      }
    }, 1200);
  };

  // 接收上层 onQuotationDataChange 包装，自动同步触发防抖保存
  const handleItemsChange = (newItems: QuotationItem[], immediate = false) => {
    onQuotationDataChange(newItems);
    if (immediate) {
      setSaveStatus('saving');
      api.saveQuotationData(currentProject, activeQuotationFilename, { items: newItems })
        .then(() => setSaveStatus('saved'))
        .catch(() => setSaveStatus('error'));
    } else {
      triggerAutoSave(newItems);
    }
  };

  const columnsToShow = quotationColumns && quotationColumns.length > 0
    ? quotationColumns
    : ['项目组', '项目名称', '单位', '数量', '不含税单价', '不含税总价', '税率', '含税单价', '含税总价', '说明'];

  const [ocrStreamText, setOcrStreamText] = useState('');

  // 手动取消/终止任务
  const handleCancelTask = async () => {
    if (!activeTaskId) return;
    try {
      await api.cancelTask(activeTaskId);
    } catch {}
    setOcrStatus('error');
    setOcrErrorMessage('已中止识别任务。');
    setOcrCurrentStep('任务已取消');
    setActiveTaskId(null);
  };

  // 图片 OCR 多图识别导入与弹出式状态窗口
  const handleOcrFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    const count = files.length;
    setOcrImageCount(count);
    setOcrStatus('processing');
    setOcrLogs([`[${new Date().toLocaleTimeString('zh-CN', { hour12: false })}] 准备上传 ${count} 张图片...`]);
    setOcrStreamText('');
    setOcrErrorMessage(undefined);
    setOcrExtractedCount(0);
    setOcrModalOpen(true);

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append('files', files[i]);
    }

    try {
      const res = await api.uploadOcrFiles(currentProject, formData);
      const taskId = res.data.task_id;
      setActiveTaskId(taskId);

      // 轮询任务状态
      const timer = setInterval(async () => {
        try {
          const statusRes = await api.checkTaskStatus(taskId);
          const task = statusRes.data;

          if (task.progress) {
            setOcrCurrentStep(task.progress);
          }

          if (Array.isArray(task.logs)) {
            setOcrLogs(task.logs);
          }

          if (typeof task.streamText === 'string') {
            setOcrStreamText(task.streamText);
          }

          if (task.status === 'done') {
            clearInterval(timer);
            setOcrStatus('done');
            setActiveTaskId(null);

            const ocrItems: Record<string, string>[] = task.result?.items || task.result?.data || [];
            setOcrExtractedCount(ocrItems.length);

            if (ocrItems.length > 0) {
              const newQuotationItems: QuotationItem[] = ocrItems.map(raw => {
                const item: QuotationItem = { _matchStatus: 'pending' };
                columnsToShow.forEach(col => {
                  item[col] = raw[col] || '';
                });
                return calculateRowFormulas(item);
              });
              const combined = [...quotationItems, ...newQuotationItems];
              handleItemsChange(combined, true); // OCR 提取完成后立即写盘持久化
            }
          } else if (task.status === 'error') {
            clearInterval(timer);
            setOcrStatus('error');
            setActiveTaskId(null);
            setOcrErrorMessage(task.message || '大模型识别过程产生错误。');
          }
        } catch (pollErr: any) {
          if (pollErr.response?.status === 404) {
            clearInterval(timer);
            setOcrStatus('error');
            setActiveTaskId(null);
            setOcrErrorMessage('任务已取消或已被移除');
          }
        }
      }, 1000);

    } catch (err: any) {
      setOcrStatus('error');
      setActiveTaskId(null);
      const errStr = err.response?.data?.detail || err.message || String(err);
      setOcrErrorMessage(errStr);
    } finally {
      e.target.value = '';
    }
  };

  // Excel 导出
  const handleExport = async () => {
    try {
      const res = await api.exportQuotation(currentProject, activeQuotationFilename);
      const url = window.URL.createObjectURL(new Blob([res.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${activeQuotationFilename.replace('.json', '')}.xlsx`);
      document.body.appendChild(link);
      link.click();
      link.remove();
    } catch {
      alert('导出失败，请先关联有效的 Excel 模板');
    }
  };

  // 定价表比对 candidate 获取
  const fetchCandidatesForItem = async (itemIndex: number, itemName: string) => {
    if (!projectRateCard || !itemName) return;
    setMatchLoadingMap(prev => ({ ...prev, [itemIndex]: true }));
    try {
      const searchRes = await api.matchRateCard(projectRateCard, [itemName], 5);
      const matches: [string, number, Record<string, string>][] = searchRes.data[itemName] || [];

      const candidates: MatchCandidate[] = matches.map(([matchedName, score, rcItem]) => ({
        name: matchedName,
        score,
        columns: Object.keys(rcItem || {}),
        values: Object.values(rcItem || {}).map(v => String(v ?? '')),
        itemData: rcItem,
      }));

      setMatchCandidatesMap(prev => ({ ...prev, [itemIndex]: candidates }));
    } catch {
      setMatchCandidatesMap(prev => ({ ...prev, [itemIndex]: [] }));
    }
    setMatchLoadingMap(prev => ({ ...prev, [itemIndex]: false }));
  };

  const handleSelectCandidate = (itemIndex: number, candidate: MatchCandidate) => {
    const item = quotationItems[itemIndex];
    const rcData = candidate.itemData || {};

    const updatedItem: QuotationItem = {
      ...item,
      _matchStatus: 'matched' as const,
      '清单名称': candidate.name,
    };

    if (rcData['不含税单价']) updatedItem['不含税单价'] = rcData['不含税单价'];
    if (rcData['含税单价']) updatedItem['含税单价'] = rcData['含税单价'];
    if (rcData['单位']) updatedItem['单位'] = rcData['单位'];
    if (rcData['税率']) updatedItem['税率'] = rcData['税率'];

    const finalItem = calculateRowFormulas(updatedItem);
    const newItems = [...quotationItems];
    newItems[itemIndex] = finalItem;
    handleItemsChange(newItems);
  };

  const handleCellChange = (index: number, col: string, val: string) => {
    const current = quotationItems[index] || {};
    const updated = { ...current, [col]: val };
    const calculated = calculateRowFormulas(updated, col);
    
    const nextItems = [...quotationItems];
    nextItems[index] = calculated;
    handleItemsChange(nextItems);
  };

  return (
    <div className="space-y-4 relative">
      {/* 顶部标题与操作栏 */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-light text-gray-800">
            {currentProject} / <span className="font-normal text-blue-600">{activeQuotationFilename}</span>
          </h1>
          <div className="flex items-center gap-4 text-xs text-gray-500 mt-1">
            <span>定价单: <strong className="text-gray-700">{projectRateCard || '未关联'}</strong></span>
            <span>模板: <strong className="text-gray-700">{projectTemplate || '未关联'}</strong></span>
          </div>
        </div>

        <div className="flex items-center gap-2.5">
          {/* 实时自动保存状态栏 */}
          <div className="text-xs flex items-center gap-1.5 px-2.5 py-1 rounded border bg-gray-50 border-gray-200">
            {saveStatus === 'saved' && (
              <>
                <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                <span className="text-gray-500">已自动保存</span>
              </>
            )}
            {saveStatus === 'saving' && (
              <>
                <Loader2 className="h-3 w-3 animate-spin text-blue-500" />
                <span className="text-blue-600 font-medium">正在保存...</span>
              </>
            )}
            {saveStatus === 'error' && (
              <span className="text-red-500 cursor-pointer" onClick={() => handleItemsChange(quotationItems, true)}>
                ⚠️ 保存失败(点击重试)
              </span>
            )}
          </div>

          <input
            ref={ocrFileInputRef}
            type="file"
            accept="image/*"
            multiple
            className="hidden"
            onChange={handleOcrFileSelect}
          />
          <Button
            variant="outline"
            size="sm"
            onClick={() => ocrFileInputRef.current?.click()}
            className="border-purple-300 text-purple-700 hover:bg-purple-50 text-xs"
          >
            <Image className="h-3.5 w-3.5 mr-1" />
            OCR 导入
          </Button>

          {projectTemplate && (
            <Button variant="outline" size="sm" onClick={handleExport} className="text-gray-700 text-xs">
              <Download className="h-3.5 w-3.5 mr-1" />
              导出
            </Button>
          )}
        </div>
      </div>

      {/* 当用户点击“后台运行”收起窗口时，在右下角提供浮动挂起 Task 卡片，点击可随时重新唤醒 Modal 终端 */}
      {!ocrModalOpen && ocrStatus === 'processing' && activeTaskId && (
        <div className="fixed bottom-6 right-6 z-40 bg-slate-900 text-white p-3.5 px-4 rounded-xl shadow-2xl flex items-center gap-3 border border-slate-700 animate-bounce">
          <Loader2 className="h-4 w-4 animate-spin text-blue-400 shrink-0" />
          <div className="text-xs">
            <p className="font-semibold text-slate-200">AI 多图识别任务运行中 ({ocrImageCount} 张图)</p>
            <p className="text-slate-400 truncate max-w-[200px] mt-0.5">{ocrCurrentStep}</p>
          </div>
          <Button
            size="xs"
            variant="secondary"
            onClick={() => setOcrModalOpen(true)}
            className="bg-slate-700 hover:bg-slate-600 text-white text-xs ml-2"
          >
            <Maximize2 className="h-3 w-3 mr-1" /> 唤醒终端
          </Button>
        </div>
      )}

      {/* 报价单数据表格 */}
      <Card className="shadow-sm">
        <CardContent className="pt-4 overflow-x-auto">
          <Table>
            <TableHeader>
              <TableRow className="bg-gray-50/80">
                <TableHead className="w-12 text-center">匹配</TableHead>
                <TableHead className="w-10">#</TableHead>
                {columnsToShow.map(col => (
                  <TableHead key={col} className="font-medium text-gray-700 whitespace-nowrap">
                    {col}
                  </TableHead>
                ))}
                <TableHead className="w-16"></TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {quotationItems.map((item, i) => {
                const matchStatus = item?._matchStatus || 'pending';
                return (
                  <TableRow key={i} className="hover:bg-gray-50/50">
                    <TableCell className="text-center">
                      {projectRateCard ? (
                        <MatchPopover
                          status={matchStatus}
                          itemName={item?.['项目名称'] || ''}
                          baseName={item?.['清单名称'] || ''}
                          candidates={matchCandidatesMap[i] || []}
                          loading={matchLoadingMap[i]}
                          onOpen={() => fetchCandidatesForItem(i, item?.['项目名称'] || '')}
                          onClose={() => setMatchCandidatesMap(prev => { const next = { ...prev }; delete next[i]; return next; })}
                          onSelect={(c) => handleSelectCandidate(i, c)}
                          onMarkCustom={() => {
                            const newItems = [...quotationItems];
                            newItems[i] = { ...item, _matchStatus: 'custom' };
                            handleItemsChange(newItems);
                          }}
                        />
                      ) : (
                        <span className="text-gray-300">-</span>
                      )}
                    </TableCell>
                    <TableCell className="text-xs text-gray-400 font-mono">{i + 1}</TableCell>
                    {columnsToShow.map(col => (
                      <TableCell key={col} className="p-1.5 min-w-[100px]">
                        <Input
                          value={item[col] ?? ''}
                          onChange={(e) => handleCellChange(i, col, e.target.value)}
                          className="h-8 text-sm focus-visible:ring-1"
                        />
                      </TableCell>
                    ))}
                    <TableCell className="whitespace-nowrap p-1.5">
                      <div className="flex items-center gap-1">
                        <Button
                          variant="ghost"
                          size="icon"
                          onClick={() => {
                            onAddRow(i);
                            const emptyRow: QuotationItem = { _matchStatus: 'pending' };
                            const next = [...quotationItems];
                            next.splice(i + 1, 0, emptyRow);
                            handleItemsChange(next);
                          }}
                          className="h-7 w-7 text-gray-400 hover:text-green-600"
                          title="插入下一行"
                        >
                          <Plus className="h-3.5 w-3.5" />
                        </Button>
                        <Button
                          variant="ghost"
                          size="icon"
                          onClick={() => {
                            onDeleteRow(i);
                            const next = quotationItems.filter((_, idx) => idx !== i);
                            handleItemsChange(next);
                          }}
                          className="h-7 w-7 text-gray-400 hover:text-red-500"
                          title="删除当前行"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </Button>
                      </div>
                    </TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>

          {quotationItems.length === 0 && (
            <div className="text-center py-10 text-gray-400 text-xs">
              暂无数据，点击上方“OCR 导入”或下方“新增行”开始填写
            </div>
          )}

          <Button
            variant="outline"
            size="sm"
            onClick={() => {
              onAddRow();
              const emptyRow: QuotationItem = { _matchStatus: 'pending' };
              const next = [...quotationItems, emptyRow];
              handleItemsChange(next);
            }}
            className="mt-4 text-xs text-gray-600 border-dashed"
          >
            <Plus className="h-3.5 w-3.5 mr-1" />
            新增数据行
          </Button>
        </CardContent>
      </Card>

      {/* 独立 OCR 流式日志与步骤进度大尺寸终端窗口 */}
      <OcrProgressModal
        isOpen={ocrModalOpen}
        status={ocrStatus}
        imageCount={ocrImageCount}
        currentStep={ocrCurrentStep}
        logs={ocrLogs}
        streamText={ocrStreamText}
        errorMessage={ocrErrorMessage}
        itemCount={ocrExtractedCount}
        onClose={() => setOcrModalOpen(false)}
        onCancel={handleCancelTask}
      />
    </div>
  );
}
