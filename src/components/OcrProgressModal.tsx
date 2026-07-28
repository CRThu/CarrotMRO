import { useEffect, useRef } from 'react';
import { Loader2, CheckCircle2, AlertTriangle, X, Terminal, Ban } from 'lucide-react';
import { Button } from '@/components/ui/button';

interface OcrProgressModalProps {
  isOpen: boolean;
  status: 'processing' | 'done' | 'error';
  imageCount: number;
  currentStep: string;
  logs: string[];
  errorMessage?: string;
  itemCount?: number;
  onClose: () => void;
  onCancel?: () => void;
}

export function OcrProgressModal({
  isOpen,
  status,
  imageCount,
  currentStep,
  logs,
  errorMessage,
  itemCount = 0,
  onClose,
  onCancel,
}: OcrProgressModalProps) {
  const logEndRef = useRef<HTMLDivElement>(null);

  // 自动滚动到最新日志底部
  useEffect(() => {
    if (isOpen) {
      logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs, isOpen]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden border border-gray-100 animate-in fade-in zoom-in duration-200">
        {/* 头部标题栏 */}
        <div className="px-6 py-4 border-b border-gray-100 flex items-center justify-between bg-slate-50/70">
          <div className="flex items-center gap-3">
            {status === 'processing' && (
              <div className="p-2 bg-blue-100 text-blue-600 rounded-xl">
                <Loader2 className="h-5 w-5 animate-spin shrink-0" />
              </div>
            )}
            {status === 'done' && (
              <div className="p-2 bg-emerald-100 text-emerald-600 rounded-xl">
                <CheckCircle2 className="h-5 w-5 shrink-0" />
              </div>
            )}
            {status === 'error' && (
              <div className="p-2 bg-red-100 text-red-600 rounded-xl">
                <AlertTriangle className="h-5 w-5 shrink-0" />
              </div>
            )}
            <div>
              <h3 className="text-lg font-semibold text-gray-800">
                {status === 'processing' && `AI 多图智能 OCR 提取中 (共 ${imageCount} 张图片)`}
                {status === 'done' && '识别完成！表格数据提取成功'}
                {status === 'error' && '识别中断或产生错误'}
              </h3>
              <p className="text-xs text-gray-500 mt-0.5">大模型多页表格提取与结构化解析引擎</p>
            </div>
          </div>

          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600 p-1.5 rounded-lg hover:bg-gray-100 transition-colors"
            title="关闭窗口"
          >
            <X className="h-5 w-5" />
          </button>
        </div>

        {/* 状态与日志主体 */}
        <div className="p-6 space-y-5">
          {/* 实时当前步骤 Banner */}
          {status === 'processing' && (
            <div className="bg-blue-50/80 border border-blue-200/80 p-4 rounded-xl flex items-center justify-between">
              <div className="flex items-center gap-3">
                <Loader2 className="h-5 w-5 text-blue-600 animate-spin shrink-0" />
                <span className="text-sm font-medium text-blue-900">{currentStep}</span>
              </div>
              <span className="text-xs text-blue-600 bg-blue-100/80 px-2.5 py-1 rounded-full font-mono">
                处理中...
              </span>
            </div>
          )}

          {status === 'done' && (
            <div className="bg-emerald-50 border border-emerald-200 p-4.5 rounded-xl text-emerald-900 flex items-start gap-3">
              <CheckCircle2 className="h-5 w-5 text-emerald-600 shrink-0 mt-0.5" />
              <div>
                <p className="text-sm font-semibold">✨ 表格数据成功合并充填！</p>
                <p className="text-xs text-emerald-700 mt-1">
                  大模型已完成全部 <strong className="text-emerald-900 font-semibold">{imageCount}</strong> 张多页图片的扫描，共提取 <strong className="text-emerald-900 font-semibold">{itemCount}</strong> 行条目并已自动计算数值公式写入当前报价单。
                </p>
              </div>
            </div>
          )}

          {status === 'error' && (
            <div className="bg-red-50 border border-red-200 p-4.5 rounded-xl text-red-900 flex items-start gap-3">
              <AlertTriangle className="h-5 w-5 text-red-600 shrink-0 mt-0.5" />
              <div className="flex-1">
                <p className="text-sm font-semibold text-red-800">识别终止与异常原因</p>
                <p className="text-xs text-red-700 mt-1.5 font-mono bg-white/80 p-3 rounded-lg border border-red-100 break-all leading-relaxed">
                  {errorMessage || '模型响应超时或网络阻断，请检查 API Key / BaseURL 设置。'}
                </p>
              </div>
            </div>
          )}

          {/* 开发者风格 Console Log 终端窗口 */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <div className="flex items-center gap-1.5 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                <Terminal className="h-3.5 w-3.5 text-gray-600" />
                <span>实时任务追踪终端 (Live Terminal Output)</span>
              </div>
              <span className="text-xs text-gray-400 font-mono">Total Logs: {logs.length}</span>
            </div>

            <div className="bg-slate-950 text-slate-200 p-4 rounded-xl font-mono text-xs h-64 overflow-y-auto space-y-2 border border-slate-800 shadow-inner">
              {logs.map((log, index) => (
                <div key={index} className="flex items-start gap-2 leading-relaxed">
                  <span className="text-blue-400 shrink-0 select-none">›</span>
                  <span className="break-all">{log}</span>
                </div>
              ))}
              {logs.length === 0 && (
                <div className="text-slate-600 italic">等待任务日志流建立...</div>
              )}
              <div ref={logEndRef} />
            </div>
          </div>
        </div>

        {/* 底部按钮栏 */}
        <div className="px-6 py-4 bg-slate-50/70 border-t border-gray-100 flex items-center justify-between">
          <div className="text-xs text-gray-500">
            {status === 'processing' && '关闭此窗口任务仍将在后台持续运行'}
          </div>

          <div className="flex items-center gap-2.5">
            {status === 'processing' && onCancel && (
              <Button
                variant="outline"
                size="sm"
                onClick={onCancel}
                className="text-red-600 border-red-200 hover:bg-red-50 hover:text-red-700"
              >
                <Ban className="h-4 w-4 mr-1.5" />
                取消任务
              </Button>
            )}

            <Button
              onClick={onClose}
              className={status === 'done' ? 'bg-emerald-600 hover:bg-emerald-700 text-white' : 'bg-slate-800 hover:bg-slate-900 text-white'}
            >
              {status === 'processing' ? '后台运行' : '关闭窗口'}
            </Button>
          </div>
        </div>
      </div>
    </div>
  );
}
