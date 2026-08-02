import { useState, useRef, useEffect } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { MatchPopover } from '@/components/MatchPopover';
import { OcrProgressModal } from '@/components/OcrProgressModal';
import { DataTable } from '@/components/DataTable';
import { MatchValidationRules, QuotationItem, PRESET_COLUMNS } from '@/types';
import * as api from '@/api';
import { Save, Download, Plus, Trash2, Image, Loader2, Maximize2, HelpCircle, AlertCircle, MessageSquare, ShieldCheck } from 'lucide-react';
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
  quotationRemarks?: string[];
  projectRateCard: string | null;
  projectTemplate: string | null;
  quotationColumns: string[];
  matchValidationRules?: MatchValidationRules;
  onEdit: (index: number, field: string, value: string) => void;
  onAddRow: (index?: number) => void;
  onDeleteRow: (index: number) => void;
  onSave: () => void;
  onQuotationDataChange: (items: QuotationItem[]) => void;
  onQuotationRemarksChange?: (remarks: string[]) => void;
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
  quotationItems = [],
  quotationRemarks = [],
  projectRateCard,
  projectTemplate,
  quotationColumns = [],
  matchValidationRules,
  onEdit,
  onAddRow,
  onDeleteRow,
  onSave,
  onQuotationDataChange,
  onQuotationRemarksChange,
}: QuotationWorkspaceProps) {
  const ocrFileInputRef = useRef<HTMLInputElement>(null);
  const columnsToShow = quotationColumns.length > 0 ? quotationColumns : (PRESET_COLUMNS as unknown as string[]);

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

  // 触发保存并更新状态 (含 items 与 remarks)
  const triggerAutoSave = (itemsToSave: QuotationItem[], remarksToSave: string[] = quotationRemarks) => {
    setSaveStatus('saving');
    if (autoSaveTimerRef.current) clearTimeout(autoSaveTimerRef.current);
    autoSaveTimerRef.current = setTimeout(async () => {
      try {
        await api.saveQuotationData(currentProject, activeQuotationFilename, {
          items: itemsToSave,
          remarks: remarksToSave,
        });
        setSaveStatus('saved');
      } catch {
        setSaveStatus('error');
      }
    }, 1200);
  };

  // 接收上层 onQuotationDataChange 包装，自动同步触发防抖保存
  const handleItemsChange = (newItems: QuotationItem[], immediate = false, remarksToSave = quotationRemarks) => {
    onQuotationDataChange(newItems);
    if (immediate) {
      setSaveStatus('saving');
      api.saveQuotationData(currentProject, activeQuotationFilename, { items: newItems, remarks: remarksToSave })
        .then(() => setSaveStatus('saved'))
        .catch(() => setSaveStatus('error'));
    } else {
      triggerAutoSave(newItems, remarksToSave);
    }
  };

  const handleRemarksChange = (newRemarks: string[]) => {
    onQuotationRemarksChange?.(newRemarks);
    triggerAutoSave(quotationItems, newRemarks);
  };

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
            const newRemarks: string[] = Array.isArray(task.result?.remarks)
              ? task.result.remarks
              : (task.result?.remarks ? [String(task.result.remarks)] : []);

            setOcrExtractedCount(ocrItems.length);

            const combinedRemarks = Array.from(new Set([...quotationRemarks, ...newRemarks]));
            onQuotationRemarksChange?.(combinedRemarks);

            if (ocrItems.length > 0 || newRemarks.length > 0) {
              const newQuotationItems: QuotationItem[] = ocrItems.map(raw => {
                const item: QuotationItem = { _matchStatus: 'pending' };
                columnsToShow.forEach(col => {
                  item[col] = raw[col] || '';
                });
                return calculateRowFormulas(item);
              });
              const combined = [...quotationItems, ...newQuotationItems];
              handleItemsChange(combined, true, combinedRemarks); // OCR 提取完成后立即写盘持久化
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
    } catch (err: any) {
      const msg = err.response?.data?.detail || err.message || '导出失败';
      alert(`报价单导出失败: ${msg}`);
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
    const fillCols = matchValidationRules?.fill_columns ?? ['单位', '不含税单价', '含税单价', '税率', '说明'];

    const updatedItem: QuotationItem = {
      ...item,
      _matchStatus: 'matched' as const,
      '清单名称': candidate.name,
      _matchedRateCardItem: candidate.itemData,
    };

    fillCols.forEach(col => {
      if (col === '项目名称') {
        if (rcData['项目名称']) {
          updatedItem['项目名称'] = rcData['项目名称'];
        }
      } else if (rcData[col] !== undefined && rcData[col] !== '') {
        updatedItem[col] = rcData[col];
      }
    });

    const finalItem = calculateRowFormulas(updatedItem);
    const newItems = [...quotationItems];
    newItems[itemIndex] = finalItem;
    handleItemsChange(newItems);
  };

  // 输出前【一键校验】：校验整张报价单的全量条目（未匹配状态警示 + 已匹配条目严格校验列比对）
  const handleValidateQuotation = () => {
    const checkCols =
      matchValidationRules?.check_columns ??
      (matchValidationRules?.strict_name_match !== false ? ['项目名称', '单位'] : ['单位']);
    const warnings: string[] = [];
    let unmatchedCount = 0;
    let customCount = 0;
    let modifiedCount = 0;

    quotationItems.forEach((item, idx) => {
      const rowNum = idx + 1;
      const itemName = item['项目名称'] || item['清单名称'] || `第${rowNum}行`;

      // 忽略空行
      const hasContent = Object.keys(item).some(k => !k.startsWith('_') && String(item[k] || '').trim() !== '');
      if (!hasContent) return;

      const matchStatus = item._matchStatus || 'pending';

      if (matchStatus === 'pending') {
        unmatchedCount++;
        warnings.push(`第 ${rowNum} 行 [${itemName}] ⚠️ 未匹配：尚未关联/匹配协议定价单物料`);
      } else if (matchStatus === 'custom') {
        customCount++;
        warnings.push(`第 ${rowNum} 行 [${itemName}] ⚠️ 自定义非标件：已被手动标记为自定义不匹配项目`);
      } else if (matchStatus === 'matched') {
        const rc = item._matchedRateCardItem;
        if (!rc) {
          unmatchedCount++;
          warnings.push(`第 ${rowNum} 行 [${itemName}] ⚠️ 缺失匹配快照：建议重新搜索选择定价物料`);
          return;
        }

        checkCols.forEach(col => {
          const curVal = String(item[col] || '').trim();
          const origVal = String(rc[col] || '').trim();

          if (col === '项目名称') {
            const listName = String(item['清单名称'] || '').trim();
            if (curVal && origVal && curVal !== origVal) {
              modifiedCount++;
              warnings.push(`第 ${rowNum} 行 [${itemName}] ❌ 项目名称误修改：当前为 "${curVal}"，原定价单为 "${origVal}"`);
            } else if (curVal && listName && curVal !== listName) {
              modifiedCount++;
              warnings.push(`第 ${rowNum} 行 [${itemName}] ❌ 名称比对不符：报价单名称 ("${curVal}") 与协议清单名称 ("${listName}") 不一致`);
            }
          } else if (curVal && origVal && curVal !== origVal) {
            modifiedCount++;
            warnings.push(`第 ${rowNum} 行 [${itemName}] ❌ ${col}误修改：当前为 "${curVal}"，原定价单为 "${origVal}"`);
          }
        });
      }
    });

    const timestampStr = new Date().toLocaleTimeString('zh-CN', { hour12: false });

    if (warnings.length === 0) {
      const successMsg = `[${timestampStr} 一键校验结果] ✓ 全表校验通过：所有 ${quotationItems.length} 条数据均已完成匹配，且严格校验列 (${checkCols.join('、')}) 与协议定价单完全一致。`;
      const combined = Array.from(new Set([...quotationRemarks, successMsg]));
      handleRemarksChange(combined);
    } else {
      const details: string[] = [];
      if (unmatchedCount > 0) details.push(`${unmatchedCount} 条未匹配`);
      if (customCount > 0) details.push(`${customCount} 条自定义非标件`);
      if (modifiedCount > 0) details.push(`${modifiedCount} 处字段误修改`);

      const summaryMsg = `[${timestampStr} 一键校验结果] 发现 ${warnings.length} 处校验提示 (${details.join('，')})：`;
      const combined = Array.from(new Set([...quotationRemarks, summaryMsg, ...warnings]));
      handleRemarksChange(combined);
    }
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

          {projectRateCard && (
            <Button
              variant="outline"
              size="sm"
              onClick={handleValidateQuotation}
              className="border-indigo-300 text-indigo-700 hover:bg-indigo-50 text-xs"
              title="一键校验全表已匹配物料的严格校验列是否与协议定价表保持一致"
            >
              <ShieldCheck className="h-3.5 w-3.5 mr-1 text-indigo-600" />
              一键校验
            </Button>
          )}

          <Button variant="outline" size="sm" onClick={handleExport} className="text-gray-700 text-xs">
            <Download className="h-3.5 w-3.5 mr-1" />
            导出
          </Button>
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

      {/* 报价单数据表格组件（单控件大视野独立双向滚动） */}
      <DataTable
        height="calc(100vh - 260px)"
        columns={columnsToShow}
        items={quotationItems}
        onEdit={handleCellChange}
        onAddRow={(i) => {
          onAddRow(i);
          const emptyRow: QuotationItem = { _matchStatus: 'pending' };
          const next = [...quotationItems];
          if (typeof i === 'number') {
            next.splice(i + 1, 0, emptyRow);
          } else {
            next.push(emptyRow);
          }
          handleItemsChange(next);
        }}
        onDeleteRow={(i) => {
          onDeleteRow(i);
          const next = quotationItems.filter((_, idx) => idx !== i);
          handleItemsChange(next);
        }}
        showRowNumber={true}
        renderPrefixHeader={() => '匹配'}
        renderPrefixCell={(item, i) => (
          projectRateCard ? (
            <MatchPopover
              status={item?._matchStatus || 'pending'}
              itemName={item?.['项目名称'] || ''}
              baseName={item?.['清单名称'] || ''}
              candidates={matchCandidatesMap[i] || []}
              loading={matchLoadingMap[i]}
              onOpen={() => fetchCandidatesForItem(i, item?.['项目名称'] || '')}
              onSearch={(query) => fetchCandidatesForItem(i, query)}
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
          )
        )}
        emptyText="暂无数据，点击上方“OCR 导入”或下方“新增行”开始填写"
        addRowText="新增数据行"
      />

      {/* 识别提示与复核备注栏 (跟随报价单 JSON 数据保存) */}
      <div className="border border-amber-200/80 bg-amber-50/50 rounded-xl p-4 space-y-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2 text-xs font-semibold text-amber-900">
            <AlertCircle className="h-4 w-4 text-amber-600 shrink-0" />
            <span>识别提示与复核备注</span>
            <span className="text-[11px] font-normal text-amber-700 bg-amber-100/80 px-2 py-0.5 rounded-full font-mono">
              {quotationRemarks.length} 条
            </span>
          </div>
          <Button
            variant="ghost"
            size="sm"
            onClick={() => {
              const next = [...quotationRemarks, ''];
              handleRemarksChange(next);
            }}
            className="h-7 text-xs text-amber-800 hover:text-amber-900 hover:bg-amber-100/60"
          >
            <Plus className="h-3.5 w-3.5 mr-1" />
            添加复核备注
          </Button>
        </div>

        {quotationRemarks.length === 0 ? (
          <div className="text-xs text-amber-600/70 italic py-1">
            暂无识别疑义或手动备注（OCR 识别时若发现字迹模糊或争议字段将在此处自动汇总）
          </div>
        ) : (
          <div className="space-y-2">
            {quotationRemarks.map((remark, idx) => (
              <div key={idx} className="flex items-center gap-2 bg-white/90 p-2 rounded-lg border border-amber-200/60 text-xs shadow-xs">
                <span className="text-[11px] font-mono font-medium text-amber-600 bg-amber-100/70 px-1.5 py-0.5 rounded shrink-0 select-none">
                  #{idx + 1}
                </span>
                <Input
                  value={remark}
                  onChange={(e) => {
                    const next = [...quotationRemarks];
                    next[idx] = e.target.value;
                    handleRemarksChange(next);
                  }}
                  placeholder="编辑备注内容或疑义说明..."
                  className="h-7 text-xs border-transparent hover:border-amber-200 focus:border-amber-400 focus:bg-white bg-transparent"
                />
                <Button
                  variant="ghost"
                  size="icon"
                  onClick={() => {
                    const next = quotationRemarks.filter((_, i) => i !== idx);
                    handleRemarksChange(next);
                  }}
                  className="h-7 w-7 text-amber-400 hover:text-red-500 hover:bg-red-50 shrink-0"
                  title="删除此项备注"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                </Button>
              </div>
            ))}
          </div>
        )}
      </div>

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
