import { OcrTableData } from '@/types';
import { DataTable } from '@/components/DataTable';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';

interface ProjectWorkspaceProps {
  currentProject: string;
  activeFilename: string | null;
  loading: boolean;
  tableData: OcrTableData;
  onEdit: (index: number, field: string, value: string) => void;
  onAddRow?: (index?: number) => void;
  onDeleteRow?: (index: number) => void;
  onSave: () => void;
}

export function ProjectWorkspace({
  currentProject,
  activeFilename,
  loading,
  tableData,
  onEdit,
  onAddRow,
  onDeleteRow,
  onSave,
}: ProjectWorkspaceProps) {
  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

      <Card>
        <CardContent className="pt-6">
          {loading && <p className="mb-4 text-blue-600 font-medium">AI 识别中...</p>}
          {activeFilename && <h3 className="mb-4 text-lg font-semibold text-gray-800">当前文件: {activeFilename}</h3>}
          <DataTable
            key={activeFilename}
            columns={tableData?.columns ?? []}
            items={tableData?.items ?? []}
            onEdit={onEdit}
            onAddRow={onAddRow}
            onDeleteRow={onDeleteRow}
          />
          <Button onClick={onSave} className="mt-6 bg-green-600 hover:bg-green-700">保存修改</Button>
        </CardContent>
      </Card>
    </>
  );
}
