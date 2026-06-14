import { useState, useEffect } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { MatchPopover } from '@/components/MatchPopover';
import { RateCardColumn, RateCardTableData, TableItem } from '@/types';
import * as api from '@/api';
import { Save, Upload } from 'lucide-react';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { Input } from '@/components/ui/input';
import { Plus, Trash2 } from 'lucide-react';

interface MatchCandidate {
  name: string;
  score: number;
  columns: string[];
  values: string[];
}

interface QuotationItem extends TableItem {
  _matchStatus?: 'pending' | 'matched' | 'custom';
  '清单名称'?: string;
}

interface QuotationWorkspaceProps {
  currentProject: string;
  activeQuotationFilename: string;
  quotationItems: QuotationItem[];
  ocrFiles: string[];
  projectRateCard: string | null;
  ratecardTableData: RateCardTableData;
  onEdit: (index: number, field: string, value: string) => void;
  onAddRow: (index?: number) => void;
  onDeleteRow: (index: number) => void;
  onSave: () => void;
  onQuotationDataChange: (items: QuotationItem[]) => void;
}

const QUOTATION_DISPLAY_COLUMNS: RateCardColumn[] = [
  { name: '序号', strict: false, alias: null, computed: true },
  { name: '项目名称', strict: true, alias: 'name' },
  { name: '清单名称', strict: false, alias: null, computed: true },
  { name: '单位', strict: false, alias: null },
  { name: '数量', strict: true, alias: 'quantity' },
  { name: '单价', strict: true, alias: 'unit_price' },
  { name: '合计', strict: false, alias: null, computed: true },
  { name: '备注', strict: false, alias: null },
];

const computeTotal = (item: TableItem): string => {
  const qty = parseFloat(item['数量'] ?? '');
  const price = parseFloat(item['单价'] ?? '');
  if (isNaN(qty) || isNaN(price)) return '';
  return (qty * price).toFixed(2);
};

export function QuotationWorkspace({
  currentProject,
  activeQuotationFilename,
  quotationItems,
  ocrFiles,
  projectRateCard,
  ratecardTableData,
  onEdit,
  onAddRow,
  onDeleteRow,
  onSave,
  onQuotationDataChange,
}: QuotationWorkspaceProps) {
  const [selectedOcrFile, setSelectedOcrFile] = useState<string>('');
  const [importUnitPrice, setImportUnitPrice] = useState(false);
  const [importing, setImporting] = useState(false);
  const [matchCandidatesMap, setMatchCandidatesMap] = useState<Record<number, MatchCandidate[]>>({});
  const [matchLoadingMap, setMatchLoadingMap] = useState<Record<number, boolean>>({});

  useEffect(() => {
    if (ocrFiles.length > 0 && !selectedOcrFile) {
      setSelectedOcrFile(ocrFiles[0]);
    }
  }, [ocrFiles, selectedOcrFile]);

  const displayItems = quotationItems.map((item, i) => ({
    ...item,
    '序号': String(i + 1),
    '合计': computeTotal(item),
  }));

  const handleImportFromOcr = async () => {
    if (!selectedOcrFile) return;
    setImporting(true);
    try {
      const res = await api.getOcrData(currentProject, selectedOcrFile);
      const raw = res.data;
      const fileData = raw.data || raw;
      const ocrItems: TableItem[] = Array.isArray(fileData.items) ? fileData.items : [];

      const newItems: QuotationItem[] = ocrItems.map((ocrItem: TableItem) => ({
        '项目名称': ocrItem['项目'] || '',
        '单位': ocrItem['单位'] || '',
        '数量': ocrItem['数量'] || '',
        '单价': importUnitPrice ? (ocrItem['单价'] || '') : '',
        '备注': '',
        _matchStatus: 'pending',
        '清单名称': '',
      }));

      onQuotationDataChange(newItems);
      setMatchCandidatesMap({});
    } catch {
      alert('导入 OCR 数据失败');
    }
    setImporting(false);
  };

  const fetchCandidatesForItem = async (itemIndex: number, itemName: string) => {
    if (!projectRateCard || !itemName) return;
    setMatchLoadingMap(prev => ({ ...prev, [itemIndex]: true }));
    try {
      const searchRes = await api.matchRateCard(projectRateCard, [itemName], 5);
      const matches: [string, number][] = searchRes.data[itemName] || [];

      const rcColumns = ratecardTableData.columns || [];
      const rcItems = ratecardTableData.items || [];

      let nameCol: string | null = null;
      for (const col of rcColumns) {
        if (col.alias === 'name') { nameCol = col.name; break; }
      }
      if (!nameCol) { setMatchLoadingMap(prev => ({ ...prev, [itemIndex]: false })); return; }

      const displayCols = rcColumns.filter((c: any) => c.alias !== 'name');
      const colNames = displayCols.map((c: any) => c.name);

      const candidates: MatchCandidate[] = matches
        .map(([matchedName, score]) => {
          const rcItem = rcItems.find((ri: any) => ri[nameCol!] === matchedName);
          if (!rcItem) return null;
          const values = displayCols.map((c: any) => String(rcItem[c.name] ?? ''));
          return { name: matchedName, score, columns: colNames, values };
        })
        .filter(Boolean) as MatchCandidate[];

      setMatchCandidatesMap(prev => ({ ...prev, [itemIndex]: candidates }));
    } catch {
      setMatchCandidatesMap(prev => ({ ...prev, [itemIndex]: [] }));
    }
    setMatchLoadingMap(prev => ({ ...prev, [itemIndex]: false }));
  };

  const handleSelectCandidate = (itemIndex: number, candidate: MatchCandidate) => {
    const item = quotationItems[itemIndex];
    const priceIdx = candidate.columns.findIndex(c => c.includes('单价') || c.includes('价格'));
    const newItems = [...quotationItems];
    newItems[itemIndex] = {
      ...item,
      '单价': priceIdx >= 0 ? candidate.values[priceIdx] : '',
      _matchStatus: 'matched' as const,
      '清单名称': candidate.name,
    };
    onQuotationDataChange(newItems);
  };

  const handleMarkCustom = (itemIndex: number) => {
    const item = quotationItems[itemIndex];
    const newItem = {
      ...item,
      _matchStatus: 'custom' as const,
      '清单名称': '',
    };
    const newItems = [...quotationItems];
    newItems[itemIndex] = newItem;
    onQuotationDataChange(newItems);
  };

  const hasRateCard = !!projectRateCard;

  return (
    <>
      <div className="flex items-center justify-between mb-6">
        <h1 className="text-3xl font-light text-gray-700">项目: {currentProject}</h1>
        <Button onClick={onSave} variant="default" size="sm">
          <Save className="h-4 w-4 mr-1" />
          保存报价单
        </Button>
      </div>

      <Card className="mb-4">
        <CardContent className="pt-4">
          <div className="flex items-center gap-4 flex-wrap">
            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-600">报价单:</span>
              <span className="text-sm font-medium">{activeQuotationFilename}</span>
            </div>

            <div className="h-5 w-px bg-gray-200" />

            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-600">清单:</span>
              <span className="text-sm font-medium">{projectRateCard || '未关联'}</span>
            </div>

            <div className="h-5 w-px bg-gray-200" />

            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-600">从OCR导入:</span>
              <select
                value={selectedOcrFile}
                onChange={(e) => setSelectedOcrFile(e.target.value)}
                className="border rounded px-2 py-1 text-sm"
              >
                {ocrFiles.map(f => (
                  <option key={f} value={f}>{f}</option>
                ))}
              </select>
              <label className="flex items-center gap-1 text-sm text-gray-600 cursor-pointer">
                <input
                  type="checkbox"
                  checked={importUnitPrice}
                  onChange={(e) => setImportUnitPrice(e.target.checked)}
                  className="rounded"
                />
                同时导入单价
              </label>
              <Button
                variant="outline"
                size="sm"
                onClick={handleImportFromOcr}
                disabled={importing || !selectedOcrFile}
              >
                <Upload className="h-4 w-4 mr-1" />
                {importing ? '导入中...' : '导入'}
              </Button>
            </div>

          </div>
        </CardContent>
      </Card>

      <Card>
        <CardContent className="pt-4">
          <Table>
            <TableHeader>
              <TableRow>
                <TableHead className="w-10">匹配</TableHead>
                <TableHead className="w-8">序号</TableHead>
                <TableHead>项目名称</TableHead>
                <TableHead>清单名称</TableHead>
                <TableHead>单位</TableHead>
                <TableHead>数量</TableHead>
                <TableHead>单价</TableHead>
                <TableHead>合计</TableHead>
                <TableHead>备注</TableHead>
                <TableHead className="w-20"></TableHead>
              </TableRow>
            </TableHeader>
            <TableBody>
              {displayItems.map((item, i) => {
                const qItem = quotationItems[i];
                const matchStatus = qItem?._matchStatus || 'pending';
                return (
                  <TableRow key={i}>
                    <TableCell>
                      {hasRateCard ? (
                        <MatchPopover
                          status={matchStatus}
                          itemName={qItem?.['项目名称'] || ''}
                          baseName={qItem?.['清单名称'] || ''}
                          candidates={matchCandidatesMap[i] || []}
                          loading={matchLoadingMap[i]}
                          onOpen={() => fetchCandidatesForItem(i, qItem?.['项目名称'] || '')}
                          onClose={() => setMatchCandidatesMap(prev => { const next = { ...prev }; delete next[i]; return next; })}
                          onSelect={(c) => handleSelectCandidate(i, c)}
                          onMarkCustom={() => handleMarkCustom(i)}
                        />
                      ) : (
                        <span className="text-gray-300">-</span>
                      )}
                    </TableCell>
                    <TableCell className="text-sm text-gray-400">{item['序号']}</TableCell>
                    <TableCell>
                      <Input
                        value={item['项目名称'] ?? ''}
                        onChange={(e) => onEdit(i, '项目名称', e.target.value)}
                      />
                    </TableCell>
                    <TableCell>
                      <span className="text-sm text-gray-600 truncate block max-w-[120px]" title={item['清单名称'] || ''}>
                        {item['清单名称'] || <span className="text-gray-300">-</span>}
                      </span>
                    </TableCell>
                    <TableCell>
                      <Input
                        value={item['单位'] ?? ''}
                        onChange={(e) => onEdit(i, '单位', e.target.value)}
                      />
                    </TableCell>
                    <TableCell>
                      <Input
                        value={item['数量'] ?? ''}
                        onChange={(e) => onEdit(i, '数量', e.target.value)}
                      />
                    </TableCell>
                    <TableCell>
                      <Input
                        value={item['单价'] ?? ''}
                        onChange={(e) => onEdit(i, '单价', e.target.value)}
                      />
                    </TableCell>
                    <TableCell className="text-sm font-medium text-gray-700">{item['合计'] || '-'}</TableCell>
                    <TableCell>
                      <Input
                        value={item['备注'] ?? ''}
                        onChange={(e) => onEdit(i, '备注', e.target.value)}
                      />
                    </TableCell>
                    <TableCell className="whitespace-nowrap">
                      <div className="flex items-center gap-0.5">
                        {onAddRow && (
                          <Button
                            variant="ghost"
                            size="icon"
                            onClick={() => onAddRow(i)}
                            className="h-7 w-7 text-gray-400 hover:text-green-600"
                            title="在下方插入行"
                          >
                            <Plus className="h-4 w-4" />
                          </Button>
                        )}
                        {onDeleteRow && (
                          <Button
                            variant="ghost"
                            size="icon"
                            onClick={() => onDeleteRow(i)}
                            className="h-7 w-7 text-gray-400 hover:text-red-500"
                            title="删除此行"
                          >
                            <Trash2 className="h-4 w-4" />
                          </Button>
                        )}
                      </div>
                    </TableCell>
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
          {onAddRow && (
            <Button
              variant="outline"
              size="sm"
              onClick={() => onAddRow()}
              className="mt-3 text-gray-500"
            >
              <Plus className="h-4 w-4 mr-1" />
              末尾新增行
            </Button>
          )}
        </CardContent>
      </Card>
    </>
  );
}
