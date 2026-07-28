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
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-light text-gray-800">项目设置: {currentProject}</h1>
          <p className="text-sm text-gray-500 mt-1">管理项目关联的定价单、导出模板以及提取与展示列规范</p>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* 关联协议定价表 */}
        <Card className="shadow-sm">
          <CardHeader>
            <CardTitle className="text-lg font-medium">关联协议定价表</CardTitle>
          </CardHeader>
          <CardContent>
            <label className="block text-sm text-gray-600 mb-2">选择在报价单中比对计价的定价表:</label>
            <select
              value={settings.ratecard_name || ''}
              onChange={(e) => onUpdateSettings({ ratecard_name: e.target.value || null })}
              className="w-full p-2.5 border border-gray-300 rounded-lg text-sm focus:ring-2 focus:ring-blue-500 focus:outline-none"
            >
              <option value="">未关联 (暂不比对)</option>
              {rateCards.map(rc => <option key={rc} value={rc}>{rc}</option>)}
            </select>
          </CardContent>
        </Card>

        {/* 关联报价单 Excel 模板 */}
        <Card className="shadow-sm">
          <CardHeader className="flex flex-row items-center justify-between">
            <CardTitle className="text-lg font-medium">关联报价单 Excel 模板</CardTitle>
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
                <Upload className="h-4 w-4 mr-1.5" />
                上传模板
              </Button>
            </div>
          </CardHeader>
          <CardContent>
            <label className="block text-sm text-gray-600 mb-2">选择导出 Excel 报价单时使用的模板:</label>
            <select
              value={settings.template_name || ''}
              onChange={(e) => onUpdateSettings({ template_name: e.target.value || null })}
              className="w-full p-2.5 border border-gray-300 rounded-lg text-sm mb-4 focus:ring-2 focus:ring-blue-500 focus:outline-none"
            >
              <option value="">未关联模板</option>
              {templates.map(t => <option key={t} value={t}>{t}</option>)}
            </select>

            {templates.length > 0 && (
              <div className="border border-gray-200 rounded-lg divide-y divide-gray-100 max-h-40 overflow-y-auto">
                {templates.map(t => (
                  <div key={t} className="flex items-center justify-between px-3 py-2 text-sm hover:bg-gray-50">
                    <span className="truncate text-gray-700">{t}</span>
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

      {/* OCR 图像抽取识别列配置 */}
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="text-lg font-medium flex items-center justify-between">
            <span>OCR 图像识别提取列规范</span>
            <span className="text-xs font-normal text-gray-400">已选中 {(settings.ocr_columns || []).length} / {PRESET_COLUMNS.length} 列</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-sm text-gray-500 mb-4">
            在报价单中使用“图片 OCR 识别导入”时，大模型将严格提取以下勾选的字段：
          </p>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-3">
            {PRESET_COLUMNS.map(col => {
              const checked = (settings.ocr_columns || []).includes(col);
              return (
                <button
                  key={col}
                  type="button"
                  onClick={() => toggleOcrColumn(col)}
                  className={`flex items-center justify-between px-3 py-2 rounded-lg border text-sm font-medium transition-all ${
                    checked
                      ? 'bg-blue-50 border-blue-400 text-blue-700 shadow-sm'
                      : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                  }`}
                >
                  <span>{col}</span>
                  {checked ? <CheckSquare className="h-4 w-4 text-blue-600 ml-1.5" /> : <Square className="h-4 w-4 text-gray-300 ml-1.5" />}
                </button>
              );
            })}
          </div>
        </CardContent>
      </Card>

      {/* 报价单展示列配置 */}
      <Card className="shadow-sm">
        <CardHeader>
          <CardTitle className="text-lg font-medium flex items-center justify-between">
            <span>报价单所需展示与编辑列规范</span>
            <span className="text-xs font-normal text-gray-400">已选中 {(settings.quotation_columns || []).length} / {PRESET_COLUMNS.length} 列</span>
          </CardTitle>
        </CardHeader>
        <CardContent>
          <p className="text-sm text-gray-500 mb-4">
            当前项目下的所有报价单表格将动态渲染以下勾选的列（支持包含/不包含税及价格公式实时联动）：
          </p>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-5 gap-3">
            {PRESET_COLUMNS.map(col => {
              const checked = (settings.quotation_columns || []).includes(col);
              return (
                <button
                  key={col}
                  type="button"
                  onClick={() => toggleQuotationColumn(col)}
                  className={`flex items-center justify-between px-3 py-2 rounded-lg border text-sm font-medium transition-all ${
                    checked
                      ? 'bg-emerald-50 border-emerald-400 text-emerald-700 shadow-sm'
                      : 'bg-white border-gray-200 text-gray-600 hover:border-gray-300'
                  }`}
                >
                  <span>{col}</span>
                  {checked ? <CheckSquare className="h-4 w-4 text-emerald-600 ml-1.5" /> : <Square className="h-4 w-4 text-gray-300 ml-1.5" />}
                </button>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
