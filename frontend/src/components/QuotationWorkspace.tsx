import { Card, CardContent } from '@/components/ui/card';

interface QuotationWorkspaceProps {
  currentProject: string;
  activeQuotationFilename: string | null;
}

export function QuotationWorkspace({
  currentProject,
  activeQuotationFilename,
}: QuotationWorkspaceProps) {
  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

      <Card>
        <CardContent className="pt-6">
          {activeQuotationFilename ? (
            <p className="text-gray-500">报价单: {activeQuotationFilename}</p>
          ) : (
            <p className="text-gray-400">请从左侧选择或创建一个报价单</p>
          )}
        </CardContent>
      </Card>
    </>
  );
}
