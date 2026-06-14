import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';

interface ProjectConfigWorkspaceProps {
  currentProject: string;
  projectRateCard: string | null;
  rateCards: string[];
  onUpdateRateCard: (name: string) => Promise<void>;
}

export function ProjectConfigWorkspace({
  currentProject,
  projectRateCard,
  rateCards,
  onUpdateRateCard,
}: ProjectConfigWorkspaceProps) {
  return (
    <>
      <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

      <Card>
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
    </>
  );
}
