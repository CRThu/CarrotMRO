import { useState } from 'react';
import { Popover, PopoverTrigger, PopoverContent } from '@/components/ui/popover';
import { Button } from '@/components/ui/button';
import { Circle, CheckCircle, AlertCircle, Search } from 'lucide-react';

interface MatchCandidate {
  name: string;
  score: number;
  columns: string[];
  values: string[];
}

interface MatchPopoverProps {
  status: 'pending' | 'matched' | 'custom';
  itemName: string;
  baseName: string;
  candidates: MatchCandidate[];
  loading?: boolean;
  onOpen: () => void;
  onClose?: () => void;
  onSelect: (candidate: MatchCandidate) => void;
  onMarkCustom: () => void;
}

export function MatchPopover({ status, itemName, baseName, candidates, loading, onOpen, onClose, onSelect, onMarkCustom }: MatchPopoverProps) {
  const [open, setOpen] = useState(false);

  const statusIcon = status === 'matched'
    ? <CheckCircle size={16} className="text-green-500" />
    : status === 'custom'
    ? <AlertCircle size={16} className="text-orange-500" />
    : <Circle size={16} className="text-gray-400" />;

  const handleOpenChange = (nextOpen: boolean) => {
    setOpen(nextOpen);
    if (nextOpen) {
      onOpen();
    } else {
      onClose?.();
    }
  };

  const handleSelect = (c: MatchCandidate) => {
    onSelect(c);
    setOpen(false);
  };

  return (
    <Popover open={open} onOpenChange={handleOpenChange}>
      <PopoverTrigger asChild>
        <button className="cursor-pointer hover:opacity-80 transition" title="匹配状态">
          {statusIcon}
        </button>
      </PopoverTrigger>
      <PopoverContent className="w-[640px] p-0" align="start">
        <div className="p-3 border-b bg-gray-50 rounded-t-md">
          <div className="flex items-center gap-2 mb-1">
            <Search size={14} className="text-gray-400" />
            <span className="text-xs font-medium text-gray-500">匹配候选</span>
          </div>
          <p className="text-sm font-medium text-gray-800 truncate">项目: {itemName || 'null'}</p>
          <p className="text-xs text-gray-500 truncate">清单: {baseName || '未匹配'}</p>
        </div>
        <div className="max-h-64 overflow-y-auto">
          {loading ? (
            <div className="p-4 text-center text-sm text-gray-400">匹配中...</div>
          ) : candidates.length === 0 ? (
            <div className="p-4 text-center text-sm text-gray-400">无候选结果</div>
          ) : (
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b text-left text-gray-400">
                  <th className="px-3 py-2 font-medium">清单名称</th>
                  {candidates[0]?.columns.map((col, i) => (
                    <th key={i} className="px-3 py-2 font-medium">{col}</th>
                  ))}
                  <th className="px-3 py-2 font-medium">相似度</th>
                </tr>
              </thead>
              <tbody>
                {candidates.map((c, i) => (
                  <tr key={i} className="border-b last:border-b-0 hover:bg-blue-50 cursor-pointer transition" onClick={() => handleSelect(c)}>
                    <td className="px-3 py-2 font-medium text-gray-800 max-w-[120px] truncate" title={c.name}>{c.name}</td>
                    {c.values.map((v, vi) => (
                      <td key={vi} className="px-3 py-2 text-gray-600 max-w-[80px] truncate" title={v}>{v || '-'}</td>
                    ))}
                    <td className="px-3 py-2 text-gray-500">{Math.round(c.score)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
        <div className="p-2 border-t bg-gray-50 rounded-b-md">
          <Button
            variant="ghost"
            size="sm"
            className="w-full text-orange-600 hover:text-orange-700 hover:bg-orange-50 text-xs"
            onClick={() => { onMarkCustom(); setOpen(false); }}
          >
            <AlertCircle size={12} className="mr-1" />
            不匹配 — 自定义
          </Button>
        </div>
      </PopoverContent>
    </Popover>
  );
}
