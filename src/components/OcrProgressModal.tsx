import { useEffect, useRef } from 'react';
import { Loader2, CheckCircle2, AlertTriangle, X, Terminal, Ban } from 'lucide-react';
import { Button } from '@/components/ui/button';

interface OcrProgressModalProps {
  isOpen: boolean;
  status: 'processing' | 'done' | 'error';
  imageCount: number;
  currentStep: string;
  logs: string[];
  streamText?: string;
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
  streamText = '',
  errorMessage,
  itemCount = 0,
  onClose,
  onCancel,
}: OcrProgressModalProps) {
  const logEndRef = useRef<HTMLDivElement>(null);
  const streamEndRef = useRef<HTMLDivElement>(null);

  // 自动滚动到最新日志与打字流底部
  useEffect(() => {
    if (isOpen) {
      logEndRef.current?.scrollIntoView({ behavior: 'smooth' });
      streamEndRef.current?.scrollIntoView({ behavior: 'smooth' });
    }
  }, [logs, streamText, isOpen]);

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-2xl max-w-3xl w-full overflow-hidden border border-gray-100 animate-in fade-in zoom-in duration-200">
        {/* 头部标题栏 */}
        <div className="px-5 py-3.5 border-b border-gray-100 flex items-center justify-between bg-slate-50/70">
          <div className="flex items-center gap-2.5">
            {status === 'processing' && (
              <div className="p-1.5 bg-blue-100 text-blue-600 rounded-lg">
                <Loader2 className="h-4 w-4 animate-spin shrink-0" />
              </div>
            )}
            {status === 'done' && (
              <div className="p-1.5 bg-emerald-100 text-emerald-600 rounded-lg">
                <CheckCircle2 className="h-4 w-4 shrink-0" />
              </div>
            )}
            {status === 'error' && (
              <div className="p-1.5 bg-red-100 text-red-600 rounded-lg">
                <AlertTriangle className="h-4 w-4 shrink-0" />
              </div>
            )}
            <div>
              <h3 className="text-base font-semibold text-gray-800">
                {status === 'processing' && `OCR 识别中 (共 ${imageCount} 张图片)`}
                {status === 'done' && '识别完成！表格数据提取成功'}
                {status === 'error' && '识别中断'}
              </h3>
            </div>
          </div>

          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600 p-1 rounded-lg hover:bg-gray-100 transition-colors"
            title="关闭"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {/* 状态与日志主体 */}
        <div className="p-5 space-y-4">
          {/* 实时当前步骤 Banner */}
          {status === 'processing' && (
            <div className="bg-blue-50/80 border border-blue-200/80 p-3 rounded-lg flex items-center justify-between text-xs">
              <div className="flex items-center gap-2.5">
                <Loader2 className="h-4 w-4 text-blue-600 animate-spin shrink-0" />
                <span className="font-medium text-blue-900">{currentStep}</span>
              </div>
              <span className="text-[11px] text-blue-600 bg-blue-100/80 px-2 py-0.5 rounded font-mono">
                流式处理中...
              </span>
            </div>
          )}

          {status === 'done' && (
            <div className="bg-emerald-50 border border-emerald-200 p-3.5 rounded-lg text-emerald-900 flex items-start gap-2.5 text-xs">
              <CheckCircle2 className="h-4 w-4 text-emerald-600 shrink-0 mt-0.5" />
              <div>
                <p className="font-semibold">✨ 提取完成！</p>
                <p className="text-emerald-700 mt-0.5">
                  已扫描 <strong>{imageCount}</strong> 张图片，成功提取 <strong>{itemCount}</strong> 行条目并填入表格。
                </p>
              </div>
            </div>
          )}

          {status === 'error' && (
            <div className="bg-red-50 border border-red-200 p-3.5 rounded-lg text-red-900 flex items-start gap-2.5 text-xs">
              <AlertTriangle className="h-4 w-4 text-red-600 shrink-0 mt-0.5" />
              <div className="flex-1">
                <p className="font-semibold text-red-800">错误日志</p>
                <p className="mt-1 font-mono bg-white/80 p-2.5 rounded border border-red-100 break-all leading-relaxed">
                  {errorMessage || '模型响应超时或网络错误'}
                </p>
              </div>
            </div>
          )}

          {/* 终端视图 */}
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            {/* 左侧/上方：任务步骤 Log */}
            <div className="md:col-span-1 border border-gray-200 rounded-lg p-3 bg-slate-50 flex flex-col justify-between">
              <div>
                <div className="flex items-center justify-between mb-2">
                  <span className="text-xs font-medium text-gray-600 flex items-center gap-1">
                    <Terminal className="h-3.5 w-3.5 text-gray-500" />
                    任务日志
                  </span>
                  <span className="text-[10px] text-gray-400 font-mono">{logs.length} 条</span>
                </div>
                <div className="font-mono text-[11px] space-y-1.5 max-h-52 overflow-y-auto pr-1">
                  {logs.map((log, index) => (
                    <div key={index} className="text-slate-700 leading-tight break-all">
                      {log}
                    </div>
                  ))}
                  <div ref={logEndRef} />
                </div>
              </div>
            </div>

            {/* 右侧/下方：AI 真实打字流 Stream Output */}
            <div className="md:col-span-2 border border-slate-800 rounded-lg p-3 bg-slate-950 text-slate-200 flex flex-col">
              <div className="flex items-center justify-between mb-2 border-b border-slate-800 pb-1.5">
                <span className="text-xs font-medium text-blue-400 flex items-center gap-1">
                  <span className="relative flex h-2 w-2">
                    {status === 'processing' && <span className="animate-ping absolute inline-flex h-full w-full rounded-full bg-blue-400 opacity-75"></span>}
                    <span className="relative inline-flex rounded-full h-2 w-2 bg-blue-500"></span>
                  </span>
                  AI 真实流式输出 (Realtime Stream)
                </span>
                <span className="text-[10px] text-slate-500 font-mono">
                  {streamText ? `${streamText.length} 字符` : '等待打字...'}
                </span>
              </div>
              <div className="font-mono text-xs max-h-52 overflow-y-auto text-emerald-400/90 whitespace-pre-wrap break-all leading-relaxed shadow-inner">
                {streamText ? (
                  <>
                    {streamText}
                    {status === 'processing' && <span className="inline-block w-2 h-4 bg-emerald-400 ml-0.5 animate-pulse align-middle" />}
                  </>
                ) : (
                  <div className="text-slate-600 italic text-xs py-4 text-center">
                    {status === 'processing' ? '正在与大模型建立打字流通道...' : '无实时数据'}
                  </div>
                )}
                <div ref={streamEndRef} />
              </div>
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
