import { TableData, TableItem } from '@/types';
import { DataTable } from '@/components/DataTable';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';

interface ProjectWorkspaceProps {
  currentProject: string;
  projectRateCard: string | null;
  rateCards: string[];
  activeFilename: string | null;
  loading: boolean;
  tableData: TableData;
  onEdit: (index: number, field: keyof TableItem, value: string) => void;
  onSave: () => void;
  onUpdateRateCard: (name: string) => Promise<void>;
}

export function ProjectWorkspace({
  currentProject,
  projectRateCard,
  rateCards,
  activeFilename,
  loading,
  tableData,
  onEdit,
  onSave,
  onUpdateRateCard,
}: ProjectWorkspaceProps) {
  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

      <Card className="mb-6">
        <CardHeader>
          <CardTitle>项目配置</CardTitle>
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

      <Card className="mb-6">
        <CardContent className="pt-6">
          {loading && <p className="mb-4 text-blue-600 font-medium">AI 识别中...</p>}
          {activeFilename && <h3 className="mb-4 text-lg font-semibold text-gray-800">当前文件: {activeFilename}</h3>}
          <DataTable key={activeFilename} items={tableData?.items ?? []} onEdit={onEdit} />
          <Button onClick={onSave} className="mt-6 bg-green-600 hover:bg-green-700">保存修改</Button>
        </CardContent>
      </Card>
    </>
  );
}
