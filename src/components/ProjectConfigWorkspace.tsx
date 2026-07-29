import { useRef } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Upload, Trash2, CheckSquare, Square } from 'lucide-react';
import { PRESET_COLUMNS, ProjectSettings } from '@/types';

interface ProjectConfigWorkspaceProps {
  currentProject: string;
  settings: ProjectSettings;
  rateCards: string[];
  templates: string[];
  onUpdateSettings: (updated: Partial<ProjectSettings>) => Promise<void>;
  onUploadTemplate: (file: File) => Promise<void>;
  onDeleteTemplate: (filename: string) => Promise<void>;
}

export function ProjectConfigWorkspace({
  currentProject,
  settings,
  rateCards,
  templates,
  onUpdateSettings,
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

  const toggleOcrColumn = (col: string) => {
    const current = settings.ocr_columns || [];
    const next = current.includes(col)
      ? current.filter(c => c !== col)
      : [...current, col];
    onUpdateSettings({ ocr_columns: next });
  };

  const toggleQuotationColumn = (col: string) => {
    const current = settings.quotation_columns || [];
    const next = current.includes(col)
      ? current.filter(c => c !== col)
      : [...current, col];
    onUpdateSettings({ quotation_columns: next });
  };

  return (
    <div className="space-y-5">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-light text-gray-800">项目设置: {currentProject}</h1>
          <p className="text-xs text-gray-500 mt-1">配置关联定价表、Excel 模板及展示与识别字段</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-5">
        {/* 关联协议定价表 */}
        <Card className="shadow-sm">
          <CardHeader className="py-4">
            <CardTitle className="text-base font-medium">关联协议定价表</CardTitle>
          </CardHeader>
          <CardContent>
            <select
              value={settings.ratecard_name || ''}
              onChange={(e) => onUpdateSettings({ ratecard_name: e.target.value || null })}
              className="w-full p-2 border border-gray-300 rounded-lg text-xs focus:ring-2 focus:ring-blue-500 focus:outline-none"
            >
              <option value="">未关联定价表</option>
              {rateCards.map(rc => <option key={rc} value={rc}>{rc}</option>)}
            </select>
          </CardContent>
        </Card>

        {/* 关联报价单 Excel 模板 */}
        <Card className="shadow-sm">
          <CardHeader className="flex flex-row items-center justify-between py-4">
            <CardTitle className="text-base font-medium">关联 Excel 导出模板</CardTitle>
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
                className="text-xs h-8"
              >
                <Upload className="h-3.5 w-3.5 mr-1" />
                上传模板
              </Button>
            </div>
          </CardHeader>
          <CardContent>
            <select
              value={settings.template_name || ''}
              onChange={(e) => onUpdateSettings({ template_name: e.target.value || null })}
              className="w-full p-2 border border-gray-300 rounded-lg text-xs mb-3 focus:ring-2 focus:ring-blue-500 focus:outline-none"
            >
              <option value="">未关联模板</option>
              {templates.map(t => <option key={t} value={t}>{t}</option>)}
            </select>

            {templates.length > 0 && (
              <div className="border border-gray-200 rounded-lg divide-y divide-gray-100 max-h-32 overflow-y-auto">
                {templates.map(t => (
                  <div key={t} className="flex items-center justify-between px-3 py-1.5 text-xs hover:bg-gray-50">
                    <span className="truncate text-gray-700">{t}</span>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="h-6 w-6 text-gray-400 hover:text-red-500"
                      onClick={() => onDeleteTemplate(t)}
                    >
                      <Trash2 className="h-3.5 w-3.5" />
                    </Button>
                  </div>
                ))}
              </div>
            )}
          </CardContent>
        </Card>
      </div>

      {/* OCR 图像抽取识别列配置 */}
      <Card className="shadow-sm">
        <CardHeader className="py-4">
          <CardTitle className="text-base font-medium flex items-center justify-between">
            <span>OCR 识别提取字段</span>
            <span className="text-xs font-normal text-gray-400">已选 {(settings.ocr_columns || []).length} / {PRESET_COLUMNS.length}</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2.5">
            {PRESET_COLUMNS.map(col => {
              const checked = (settings.ocr_columns || []).includes(col);
              return (
                <button
                  key={col}
                  type="button"
                  onClick={() => toggleOcrColumn(col)}
                  className={`flex items-center justify-between px-3 py-1.5 rounded-lg border text-xs font-medium transition-all ${
                    checked
                      ? 'bg-blue-50 border-blue-400 text-blue-700 shadow-sm'
                      : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                  }`}
                >
                  <span>{col}</span>
                  {checked ? <CheckSquare className="h-3.5 w-3.5 text-blue-600 ml-1" /> : <Square className="h-3.5 w-3.5 text-gray-300 ml-1" />}
                </button>
              );
            })}
          </div>
        </CardContent>
      </Card>

      {/* 报价单展示列配置 */}
      <Card className="shadow-sm">
        <CardHeader className="py-4">
          <CardTitle className="text-base font-medium flex items-center justify-between">
            <span>报价单表格展示字段</span>
            <span className="text-xs font-normal text-gray-400">已选 {(settings.quotation_columns || []).length} / {PRESET_COLUMNS.length}</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-2.5">
            {PRESET_COLUMNS.map(col => {
              const checked = (settings.quotation_columns || []).includes(col);
              return (
                <button
                  key={col}
                  type="button"
                  onClick={() => toggleQuotationColumn(col)}
                  className={`flex items-center justify-between px-3 py-1.5 rounded-lg border text-xs font-medium transition-all ${
                    checked
                      ? 'bg-emerald-50 border-emerald-400 text-emerald-700 shadow-sm'
                      : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                  }`}
                >
                  <span>{col}</span>
                  {checked ? <CheckSquare className="h-3.5 w-3.5 text-emerald-600 ml-1" /> : <Square className="h-3.5 w-3.5 text-gray-300 ml-1" />}
                </button>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
