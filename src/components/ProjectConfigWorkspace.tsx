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

      {/* 定价匹配与校验规则设置 */}
      <Card className="shadow-sm">
        <CardHeader className="py-4">
          <CardTitle className="text-base font-medium flex items-center justify-between">
            <span>定价单匹配带入与严格校验规则</span>
            <div className="flex items-center gap-4 text-xs font-normal text-gray-500">
              <span className="flex items-center gap-1">
                <span className="w-2.5 h-2.5 rounded bg-blue-500 inline-block"></span> 匹配带入 (
                {
                  (settings.match_validation_rules?.fill_columns ?? ['单位', '不含税单价', '含税单价', '税率', '说明']).length
                }
                )
              </span>
              <span className="flex items-center gap-1">
                <span className="w-2.5 h-2.5 rounded bg-purple-500 inline-block"></span> 严格校验 (
                {
                  (
                    settings.match_validation_rules?.check_columns ??
                    (settings.match_validation_rules?.strict_name_match !== false ? ['项目名称', '单位'] : ['单位'])
                  ).length
                }
                )
              </span>
            </div>
          </CardTitle>
        </CardHeader>
        <CardContent className="space-y-4">
          <div className="p-3 bg-blue-50/50 rounded-lg border border-blue-100 text-xs text-gray-600 space-y-1">
            <p className="font-semibold text-blue-900 flex items-center gap-1.5">
              <span>💡 功能区别说明</span>
            </p>
            <p>1. <strong>匹配带入 (📥)</strong>：在报价单中搜索并选择协议物料时，将选中的字段自动填充到报价单中并自动联动计算单价与总价。</p>
            <p>2. <strong>严格校验 (🔍)</strong>：在成品输出前点击顶部【校验】按钮时，对比报价单当前数值与协议定价表，防止关键校验列被误修改。</p>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-5 gap-3">
            {PRESET_COLUMNS.map(col => {
              const currentRules = settings.match_validation_rules || {
                strict_name_match: true,
                check_columns: ['项目名称', '单位'],
                fill_columns: ['单位', '不含税单价', '含税单价', '税率', '说明'],
              };
              const checkCols =
                currentRules.check_columns ??
                (currentRules.strict_name_match !== false ? ['项目名称', '单位'] : ['单位']);
              const fillCols = currentRules.fill_columns ?? ['单位', '不含税单价', '含税单价', '税率', '说明'];

              const isFill = fillCols.includes(col);
              const isCheck = checkCols.includes(col);

              const toggleFill = () => {
                const nextFill = isFill ? fillCols.filter(c => c !== col) : [...fillCols, col];
                onUpdateSettings({
                  match_validation_rules: {
                    ...currentRules,
                    strict_name_match: checkCols.includes('项目名称'),
                    check_columns: checkCols,
                    fill_columns: nextFill,
                  },
                });
              };

              const toggleCheck = () => {
                const nextCheck = isCheck ? checkCols.filter(c => c !== col) : [...checkCols, col];
                onUpdateSettings({
                  match_validation_rules: {
                    ...currentRules,
                    strict_name_match: nextCheck.includes('项目名称'),
                    check_columns: nextCheck,
                    fill_columns: fillCols,
                  },
                });
              };

              return (
                <div key={col} className="p-2.5 rounded-lg border border-gray-200 bg-white shadow-2xs space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="text-xs font-semibold text-gray-800">{col}</span>
                  </div>
                  <div className="grid grid-cols-2 gap-1.5">
                    <button
                      type="button"
                      onClick={toggleFill}
                      className={`flex items-center justify-center gap-1 px-2 py-1 rounded text-[11px] font-medium transition-all ${
                        isFill
                          ? 'bg-blue-50 border border-blue-400 text-blue-700'
                          : 'bg-gray-50 border border-gray-200 text-gray-500 hover:bg-gray-100'
                      }`}
                      title={`${col} - 匹配时自动填充带入`}
                    >
                      {isFill ? <CheckSquare className="h-3 w-3 text-blue-600" /> : <Square className="h-3 w-3 text-gray-400" />}
                      <span>带入</span>
                    </button>

                    <button
                      type="button"
                      onClick={toggleCheck}
                      className={`flex items-center justify-center gap-1 px-2 py-1 rounded text-[11px] font-medium transition-all ${
                        isCheck
                          ? 'bg-purple-50 border border-purple-400 text-purple-700'
                          : 'bg-gray-50 border border-gray-200 text-gray-500 hover:bg-gray-100'
                      }`}
                      title={`${col} - 输出前一键校验检查误改动`}
                    >
                      {isCheck ? <CheckSquare className="h-3 w-3 text-purple-600" /> : <Square className="h-3 w-3 text-gray-400" />}
                      <span>校验</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </CardContent>
      </Card>
    </div>
  );
}
