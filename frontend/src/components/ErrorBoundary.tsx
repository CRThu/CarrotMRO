import { Component, ReactNode } from 'react';
import { AlertTriangle, Copy, Check } from 'lucide-react';

interface Props { children: ReactNode; }
interface State { hasError: boolean; error: Error | null; componentStack: string; copied: boolean; }

export class ErrorBoundary extends Component<Props, State> {
  state: State = { hasError: false, error: null, componentStack: '', copied: false };

  static getDerivedStateFromError(error: Error): State {
    return { hasError: true, error, componentStack: '', copied: false };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo) {
    this.setState({ componentStack: info.componentStack || '' });
    console.error('ErrorBoundary caught:', error, info);
  }

  handleCopy = () => {
    const { error, componentStack } = this.state;
    const text = [
      `Error: ${error?.message}`,
      '',
      'Component Stack:',
      componentStack,
      '',
      'Full Stack:',
      error?.stack,
    ].join('\n');
    navigator.clipboard.writeText(text).then(() => {
      this.setState({ copied: true });
      setTimeout(() => this.setState({ copied: false }), 2000);
    });
  };

  render() {
    if (this.state.hasError) {
      const { error, componentStack, copied } = this.state;
      return (
        <div className="flex h-screen items-center justify-center bg-gray-100 p-8">
          <div className="bg-white p-8 rounded-2xl shadow-lg max-w-2xl w-full">
            <div className="flex items-center gap-3 mb-4">
              <AlertTriangle className="text-red-500" size={28} />
              <h1 className="text-2xl font-bold text-red-500">应用发生错误</h1>
            </div>

            <div className="bg-red-50 border border-red-200 rounded-lg p-4 mb-4">
              <p className="text-red-700 font-medium text-sm">{error?.message || '未知错误'}</p>
            </div>

            {componentStack && (
              <details className="mb-4" open>
                <summary className="text-sm font-medium text-gray-600 cursor-pointer mb-2">组件调用栈</summary>
                <pre className="bg-gray-50 border rounded-lg p-4 text-xs text-gray-700 overflow-auto max-h-48 whitespace-pre-wrap">
                  {componentStack}
                </pre>
              </details>
            )}

            {error?.stack && (
              <details className="mb-6">
                <summary className="text-sm font-medium text-gray-600 cursor-pointer mb-2">完整堆栈</summary>
                <pre className="bg-gray-50 border rounded-lg p-4 text-xs text-gray-700 overflow-auto max-h-64 whitespace-pre-wrap">
                  {error.stack}
                </pre>
              </details>
            )}

            <div className="flex gap-3">
              <button
                onClick={this.handleCopy}
                className="flex items-center gap-2 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition text-sm"
              >
                {copied ? <Check size={14} /> : <Copy size={14} />}
                {copied ? '已复制' : '复制错误信息'}
              </button>
              <button
                onClick={() => window.location.reload()}
                className="px-4 py-2 bg-blue-500 text-white rounded-lg hover:bg-blue-600 transition text-sm"
              >
                重新加载
              </button>
            </div>
          </div>
        </div>
      );
    }
    return this.props.children;
  }
}
