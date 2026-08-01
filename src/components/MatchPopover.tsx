import { useState, useRef } from 'react';
import { Popover, PopoverTrigger, PopoverContent } from '@/components/ui/popover';
import { Button } from '@/components/ui/button';
import { Circle, CheckCircle, AlertCircle, Search, Loader2, X } from 'lucide-react';

interface MatchCandidate {
  name: string;
  score: number;
  columns: string[];
  values: string[];
  itemData?: Record<string, string>;
}

interface MatchPopoverProps {
  status: 'pending' | 'matched' | 'custom';
  itemName: string;
  baseName: string;
  candidates: MatchCandidate[];
  loading?: boolean;
  onOpen: () => void;
  onSearch?: (query: string) => void;
  onClose?: () => void;
  onSelect: (candidate: MatchCandidate) => void;
  onMarkCustom: () => void;
}

export function MatchPopover({
  status,
  itemName,
  baseName,
  candidates,
  loading,
  onOpen,
  onSearch,
  onClose,
  onSelect,
  onMarkCustom,
}: MatchPopoverProps) {
  const [open, setOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState(itemName || '');
  const debounceTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const statusIcon = status === 'matched'
    ? <CheckCircle size={16} className="text-green-500" />
    : status === 'custom'
    ? <AlertCircle size={16} className="text-orange-500" />
    : <Circle size={16} className="text-gray-400" />;

  /** 立即执行搜索（清除防抖定时器后直接调用） */
  const executeSearch = (query: string) => {
    if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);
    if (onSearch) {
      onSearch(query);
    } else {
      onOpen();
    }
  };

  /** 输入变化时 250ms 防抖自动触发 */
  const handleInputChange = (val: string) => {
    setSearchQuery(val);
    if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);
    debounceTimerRef.current = setTimeout(() => {
      executeSearch(val);
    }, 250);
  };

  const handleOpenChange = (nextOpen: boolean) => {
    setOpen(nextOpen);
    if (nextOpen) {
      // 弹窗打开时用当前行的项目名称初始化搜索框并发起首次搜索
      setSearchQuery(itemName || '');
      onOpen();
      if (onSearch) {
        onSearch(itemName || '');
      }
    } else {
      if (debounceTimerRef.current) clearTimeout(debounceTimerRef.current);
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
        <button className="cursor-pointer hover:opacity-80 transition" title="点击打开物料匹配与搜索">
          {statusIcon}
        </button>
      </PopoverTrigger>
      <PopoverContent className="w-[640px] p-0 shadow-lg" align="start">
        {/* 搜索头部：输入框支持 250ms 防抖自动触发与回车键立即触发 */}
        <div className="p-3 border-b bg-gray-50 rounded-t-md space-y-2">
          <div className="relative">
            <Search size={14} className="text-gray-400 absolute left-2.5 top-2.5" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => handleInputChange(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  executeSearch(searchQuery);
                }
              }}
              placeholder="输入关键字检索协议定价库物料..."
              className="w-full pl-8 pr-7 py-1 bg-white border border-gray-300 rounded-md text-xs focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-blue-500 transition"
            />
            {searchQuery && (
              <button
                type="button"
                onClick={() => handleInputChange('')}
                className="absolute right-2 top-2 text-gray-400 hover:text-gray-600 cursor-pointer"
                title="清空关键字"
              >
                <X size={12} />
              </button>
            )}
          </div>

          <div className="flex items-center justify-between text-[11px] text-gray-500 pt-0.5">
            <span className="truncate max-w-[280px]">原项目名称: <strong className="text-gray-700">{itemName || '未填写'}</strong></span>
            <span className="truncate max-w-[280px]">当前关联清单: <strong className="text-gray-700">{baseName || '未匹配'}</strong></span>
          </div>
        </div>

        {/* 候选列表 */}
        <div className="max-h-64 overflow-y-auto">
          {loading ? (
            <div className="p-8 text-center text-sm text-blue-600 flex flex-col items-center justify-center gap-2">
              <Loader2 className="h-5 w-5 animate-spin" />
              <span>正在检索 "{searchQuery || '全部物料'}"...</span>
            </div>
          ) : candidates.length === 0 ? (
            <div className="p-6 text-center space-y-2">
              <div className="inline-flex p-2.5 rounded-full bg-amber-50 border border-amber-200 text-amber-600 mb-1">
                <AlertCircle className="h-5 w-5" />
              </div>
              <p className="text-sm font-medium text-gray-700">未找到与 "{searchQuery}" 匹配的协议物料</p>
              <p className="text-xs text-gray-400 max-w-sm mx-auto leading-relaxed">
                建议尝试使用更简短的关键词（如"地毯"、"找平"、"拆除"）重新搜索
              </p>
            </div>
          ) : (
            <table className="w-full text-xs">
              <thead>
                <tr className="border-b text-left text-gray-400 bg-gray-50/50">
                  <th className="px-3 py-2 font-medium">清单名称</th>
                  {candidates[0]?.columns.map((col, i) => (
                    <th key={i} className="px-3 py-2 font-medium">{col}</th>
                  ))}
                  <th className="px-3 py-2 font-medium">匹配度</th>
                </tr>
              </thead>
              <tbody>
                {candidates.map((c, i) => {
                  const scoreVal = Math.round(c.score);
                  return (
                    <tr key={i} className="border-b last:border-b-0 hover:bg-blue-50 cursor-pointer transition" onClick={() => handleSelect(c)}>
                      <td className="px-3 py-2 font-medium text-gray-800 max-w-[120px] truncate" title={c.name}>{c.name}</td>
                      {c.values.map((v, vi) => (
                        <td key={vi} className="px-3 py-2 text-gray-600 max-w-[80px] truncate" title={v}>{v || '-'}</td>
                      ))}
                      <td className="px-3 py-2 text-gray-500">
                        <span className={scoreVal === 100 ? 'text-emerald-600 font-semibold' : 'text-blue-600 font-semibold'}>
                          {scoreVal}%
                        </span>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          )}
        </div>

        {/* 底部：标记自定义 */}
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
