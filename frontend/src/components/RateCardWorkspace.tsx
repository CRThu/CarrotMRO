import { useRef } from 'react';
import { TableData, TableItem } from '@/types';
import { DataTable } from '@/components/DataTable';
import { Button } from '@/components/ui/button';
import { Card, CardContent } from '@/components/ui/card';

interface RateCardWorkspaceProps {
  currentRateCard: string;
  ratecardTableData: TableData;
  importing: boolean;
  onEdit: (index: number, field: keyof TableItem, value: string) => void;
  onImport: (file: File) => void;
}

export function RateCardWorkspace({
  currentRateCard,
  ratecardTableData,
  importing,
  onEdit,
  onImport,
}: RateCardWorkspaceProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">协议定价表: {currentRateCard}</h1>

      <Card>
        <CardContent className="pt-6">
          <div className="flex items-center justify-between mb-6">
            <div></div>
            <div className="flex gap-3">
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                className="hidden"
                onChange={(e) => {
                  const file = e.target.files?.[0];
                  if (file) onImport(file);
                  e.target.value = '';
                }}
              />
              <Button
                onClick={() => fileInputRef.current?.click()}
                disabled={importing}
                variant="outline"
                className="border-amber-500 text-amber-600 hover:bg-amber-50"
              >
                {importing ? '导入中...' : '导入 Excel / CSV'}
              </Button>
            </div>
          </div>

          <DataTable items={ratecardTableData?.items ?? []} onEdit={onEdit} />

          {ratecardTableData.remarks && (
            <div className="mt-4 p-3 bg-gray-50 rounded-lg text-sm text-gray-600">
              {ratecardTableData.remarks}
            </div>
          )}
        </CardContent>
      </Card>
    </>
  );
}
