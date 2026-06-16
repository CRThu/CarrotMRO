import { useRef } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Upload, Trash2, CheckSquare, Square } from 'lucide-react';

interface ProjectConfigWorkspaceProps {
  currentProject: string;
  projectRateCard: string | null;
  rateCards: string[];
  projectTemplate: string | null;
  templates: string[];
  availableColumns: string[];
  selectedColumns: string[];
  onUpdateRateCard: (name: string) => Promise<void>;
  onUpdateTemplate: (templateName: string) => Promise<void>;
  onUpdateColumns: (columns: string[]) => Promise<void>;
  onUploadTemplate: (file: File) => Promise<void>;
  onDeleteTemplate: (filename: string) => Promise<void>;
}

export function ProjectConfigWorkspace({
  currentProject,
  projectRateCard,
  rateCards,
  projectTemplate,
  templates,
  availableColumns,
  selectedColumns,
  onUpdateRateCard,
  onUpdateTemplate,
  onUpdateColumns,
  onUploadTemplate,
  onDeleteTemplate,
}: ProjectConfigWorkspaceProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onUploadTemplate(file);
      e.target.value = '';
    }
  };

  const toggleColumn = (col: string) => {
    const next = selectedColumns.includes(col)
      ? selectedColumns.filter(c => c !== col)
      : [...selectedColumns, col];
    onUpdateColumns(next);
  };

  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <Card>
          <CardHeader>
            <CardTitle>协议定价表</CardTitle>
          </CardHeader>
          <CardContent>
            <label className="block text-sm text-gray-600 mb-2">关联协议定价表:</label>
            <select
              value={projectRateCard || ''}
              onChange={(e) => onUpdateRateCard(e.target.value)}
              className="w-full p-2 border rounded-lg"
            >
              <option value="">未关联</option>
              {rateCards.map(rc => <option key={rc} value={rc}>{rc}</option>)}
            </select>
          </CardContent>
        </Card>

        <Card>
          <CardHeader className="flex flex-row items-center justify-between">
            <CardTitle>报价单模板</CardTitle>
            <div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx"
                onChange={handleFileChange}
                className="hidden"
              />
              <Button
                variant="outline"
                size="sm"
                onClick={() => fileInputRef.current?.click()}
              >
                <Upload className="h-4 w-4 mr-1" />
                上传模板
              </Button>
            </div>
          </CardHeader>
          <CardContent>
            <label className="block text-sm text-gray-600 mb-2">关联模板:</label>
            <select
              value={projectTemplate || ''}
              onChange={(e) => onUpdateTemplate(e.target.value)}
              className="w-full p-2 border rounded-lg mb-4"
            >
              <option value="">未关联</option>
              {templates.map(t => <option key={t} value={t}>{t}</option>)}
            </select>

            {templates.length > 0 && (
              <div className="border rounded-lg divide-y">
                {templates.map(t => (
                  <div key={t} className="flex items-center justify-between px-3 py-2">
                    <span className="text-sm truncate">{t}</span>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="h-7 w-7 text-gray-400 hover:text-red-500"
                      onClick={() => onDeleteTemplate(t)}
                    >
                      <Trash2 className="h-4 w-4" />
                    </Button>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      {projectTemplate && availableColumns.length > 0 && (
        <Card className="mt-6">
          <CardHeader>
            <CardTitle>OCR 识别列配置</CardTitle>
          </CardHeader>
          <CardContent>
            <p className="text-sm text-gray-500 mb-3">
              勾选需要识别的列，OCR 将只提取这些字段：
            </p>
            <div className="flex flex-wrap gap-3">
              {availableColumns.map(col => {
                const checked = selectedColumns.includes(col);
                return (
                  <button
                    key={col}
                    onClick={() => toggleColumn(col)}
                    className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg border text-sm transition-colors ${
                      checked
                        ? 'bg-blue-50 border-blue-300 text-blue-700'
                        : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                    }`}
                  >
                    {checked ? <CheckSquare className="h-4 w-4" /> : <Square className="h-4 w-4" />}
                    {col}
                  </button>
                );
              })}
            </div>
            {selectedColumns.length === 0 && (
              <p className="text-sm text-amber-600 mt-2">请至少选择一列，否则无法进行 OCR 识别</p>
            )}
          </CardContent>
        </Card>
      )}
    </>
  );
}
