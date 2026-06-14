import { useState, useEffect } from 'react';
import { Card, CardContent } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { DataTable } from '@/components/DataTable';
import { RateCardColumn, TableItem } from '@/types';
import * as api from '@/api';
import { Save, Upload, DollarSign, Check, X } from 'lucide-react';

interface MatchCandidate {
  name: string;
  price: string;
  score: number;
}

interface QuotationWorkspaceProps {
  currentProject: string;
  activeQuotationFilename: string;
  quotationItems: TableItem[];
  ocrFiles: string[];
  projectRateCard: string | null;
  onEdit: (index: number, field: string, value: string) => void;
  onAddRow: (index?: number) => void;
  onDeleteRow: (index: number) => void;
  onSave: () => void;
  onQuotationDataChange: (items: TableItem[]) => void;
}

const QUOTATION_DISPLAY_COLUMNS: RateCardColumn[] = [
  { name: '序号', strict: false, alias: null, computed: true },
  { name: '项目名称', strict: true, alias: 'name' },
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
  onEdit,
  onAddRow,
  onDeleteRow,
  onSave,
  onQuotationDataChange,
}: QuotationWorkspaceProps) {
  const [selectedOcrFile, setSelectedOcrFile] = useState<string>('');
  const [importUnitPrice, setImportUnitPrice] = useState(false);
  const [importing, setImporting] = useState(false);
  const [matchPricesLoading, setMatchPricesLoading] = useState(false);
  const [matchResults, setMatchResults] = useState<Record<number, MatchCandidate[]>>({});
  const [showMatchPanel, setShowMatchPanel] = useState(false);

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

      const newItems: TableItem[] = ocrItems.map((ocrItem: TableItem) => ({
        '项目名称': ocrItem['项目'] || '',
        '单位': ocrItem['单位'] || '',
        '数量': ocrItem['数量'] || '',
        '单价': importUnitPrice ? (ocrItem['单价'] || '') : '',
        '备注': '',
      }));

      onQuotationDataChange(newItems);
      setShowMatchPanel(false);
      setMatchResults({});
    } catch {
      alert('导入 OCR 数据失败');
    }
    setImporting(false);
  };

  const handleMatchPrices = async () => {
    if (!projectRateCard || quotationItems.length === 0) return;
    setMatchPricesLoading(true);
    try {
      const names = quotationItems.map(item => item['项目名称'] || '').filter(Boolean);
      if (names.length === 0) {
        alert('没有可匹配的项目名称');
        setMatchPricesLoading(false);
        return;
      }

      const searchRes = await api.matchRateCard(projectRateCard, names, 5);
      const searchResults: Record<string, [string, number][]> = searchRes.data;

      const rcDataRes = await api.getRateCardData(projectRateCard);
      const rcData = rcDataRes.data;
      const rcColumns = rcData.columns || [];
      const rcItems = rcData.items || [];

      let nameCol: string | null = null;
      let priceCol: string | null = null;
      for (const col of rcColumns) {
        if (col.alias === 'name') nameCol = col.name;
        if (col.name.includes('单价') || col.name.includes('价格')) priceCol = col.name;
      }

      if (!nameCol || !priceCol) {
        alert('定价表缺少名称列或单价列');
        setMatchPricesLoading(false);
        return;
      }

      const results: Record<number, MatchCandidate[]> = {};
      quotationItems.forEach((item, idx) => {
        const itemName = item['项目名称'] || '';
        const matches = searchResults[itemName];
        if (matches && matches.length > 0) {
          const candidates: MatchCandidate[] = matches
            .map(([matchedName, score]) => {
              const rcItem = rcItems.find((ri: any) => ri[nameCol!] === matchedName);
              return {
                name: matchedName,
                price: rcItem && rcItem[priceCol!] ? String(rcItem[priceCol!]) : '',
                score,
              };
            })
            .filter(c => c.price);
          if (candidates.length > 0) {
            results[idx] = candidates;
          }
        }
      });

      setMatchResults(results);
      setShowMatchPanel(true);
    } catch {
      alert('匹配单价失败');
    }
    setMatchPricesLoading(false);
  };

  const handleSelectMatch = (itemIndex: number, price: string) => {
    onEdit(itemIndex, '单价', price);
  };

  const handleCloseMatchPanel = () => {
    setShowMatchPanel(false);
    setMatchResults({});
  };

  const matchedItemCount = Object.keys(matchResults).length;

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

            {projectRateCard && (
              <>
                <div className="h-5 w-px bg-gray-200" />
                <Button
                  variant="outline"
                  size="sm"
                  onClick={handleMatchPrices}
                  disabled={matchPricesLoading || quotationItems.length === 0}
                >
                  <DollarSign className="h-4 w-4 mr-1" />
                  {matchPricesLoading ? '匹配中...' : '从定价表匹配单价'}
                </Button>
                <span className="text-xs text-gray-400">({projectRateCard})</span>
              </>
            )}
          </div>
        </CardContent>
      </Card>

      {showMatchPanel && matchedItemCount > 0 && (
        <Card className="mb-4">
          <CardContent className="pt-4">
            <div className="flex items-center justify-between mb-3">
              <h3 className="text-sm font-medium text-gray-700">
                匹配结果（{matchedItemCount} 个项目有候选，请点击选择单价）
              </h3>
              <Button variant="ghost" size="sm" onClick={handleCloseMatchPanel}>
                <X className="h-4 w-4 mr-1" />
                关闭
              </Button>
            </div>
            <div className="max-h-80 overflow-y-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b text-left text-gray-500">
                    <th className="pb-2 pr-4 w-8">序号</th>
                    <th className="pb-2 pr-4">项目名称</th>
                    <th className="pb-2 pr-4">当前单价</th>
                    <th className="pb-2">候选匹配（点击选择）</th>
                  </tr>
                </thead>
                <tbody>
                  {quotationItems.map((item, idx) => {
                    const candidates = matchResults[idx];
                    if (!candidates) return null;
                    return (
                      <tr key={idx} className="border-b last:border-b-0">
                        <td className="py-2 pr-4 text-gray-400">{idx + 1}</td>
                        <td className="py-2 pr-4">{item['项目名称']}</td>
                        <td className="py-2 pr-4 text-gray-500">{item['单价'] || '-'}</td>
                        <td className="py-2">
                          <div className="flex flex-wrap gap-2">
                            {candidates.map((c, ci) => (
                              <button
                                key={ci}
                                onClick={() => handleSelectMatch(idx, c.price)}
                                className={`inline-flex items-center gap-1 px-2 py-1 rounded border text-xs transition-colors ${
                                  item['单价'] === c.price
                                    ? 'bg-green-50 border-green-300 text-green-700'
                                    : 'bg-white border-gray-200 hover:border-blue-300 hover:bg-blue-50'
                                }`}
                              >
                                <span className="font-medium">{c.name}</span>
                                <span className="text-gray-400">|</span>
                                <span>{c.price}</span>
                                <span className="text-gray-300">({Math.round(c.score)})</span>
                              </button>
                            ))}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
      )}

      {showMatchPanel && matchedItemCount === 0 && (
        <Card className="mb-4">
          <CardContent className="pt-4">
            <div className="flex items-center justify-between">
              <span className="text-sm text-gray-500">没有匹配到任何候选结果</span>
              <Button variant="ghost" size="sm" onClick={handleCloseMatchPanel}>
                <X className="h-4 w-4 mr-1" />
                关闭
              </Button>
            </div>
          </CardContent>
        </Card>
      )}

      <Card>
        <CardContent className="pt-4">
          <DataTable
            columns={QUOTATION_DISPLAY_COLUMNS}
            items={displayItems}
            onEdit={(index, field, value) => {
              if (field === '序号' || field === '合计') return;
              onEdit(index, field, value);
            }}
            onAddRow={onAddRow}
            onDeleteRow={onDeleteRow}
          />
        </CardContent>
      </Card>
    </>
  );
}
