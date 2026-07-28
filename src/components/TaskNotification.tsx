import { useEffect } from 'react';
import { CheckCircle, XCircle, Loader2 } from 'lucide-react';

type TaskStatus = 'processing' | 'done' | 'error';

interface TaskNotificationProps {
  status: { status: TaskStatus; progress?: string; message?: string } | null;
  labels?: Partial<Record<TaskStatus, string>>;
  autoDismissMs?: number;
  onDismiss: () => void;
}

const DEFAULT_LABELS: Record<TaskStatus, string> = {
  processing: '处理中...',
  done: '完成',
  error: '失败',
};

export function TaskNotification({ status, labels, autoDismissMs = 5000, onDismiss }: TaskNotificationProps) {
  const merged = { ...DEFAULT_LABELS, ...labels };

  useEffect(() => {
    if (status && status.status !== 'processing') {
      const timer = setTimeout(onDismiss, autoDismissMs);
      return () => clearTimeout(timer);
    }
  }, [status, onDismiss, autoDismissMs]);

  if (!status) return null;

  const baseLabel = merged[status.status];
  const displayText = status.progress
    ? status.progress
    : status.message
    ? `${baseLabel}: ${status.message}`
    : baseLabel;

  return (
    <div className="fixed bottom-6 right-6 z-50">
      {status.status === 'processing' && (
        <div className="flex items-center gap-3 bg-blue-600 text-white px-5 py-3 rounded-lg shadow-lg border border-blue-500 animate-pulse">
          <Loader2 size={18} className="animate-spin shrink-0" />
          <span className="text-sm font-medium">{displayText}</span>
        </div>
      )}
      {status.status === 'done' && (
        <div className="flex items-center gap-3 bg-green-600 text-white px-5 py-3 rounded-lg shadow-lg border border-green-500">
          <CheckCircle size={18} className="shrink-0" />
          <span className="text-sm font-medium">{displayText}</span>
        </div>
      )}
      {status.status === 'error' && (
        <div className="flex items-center gap-3 bg-red-600 text-white px-5 py-3 rounded-lg shadow-lg border border-red-500">
          <XCircle size={18} className="shrink-0" />
          <span className="text-sm font-medium">{displayText}</span>
        </div>
      )}
    </div>
  );
}
