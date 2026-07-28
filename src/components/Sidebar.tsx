import { SidebarSection, TreeItem, Folder, FolderOpen, FileText, Receipt } from '@/components/SidebarTree';
import { Settings, Sliders } from 'lucide-react';

interface SidebarProps {
  projects: string[];
  rateCards: string[];
  quotationFiles: string[];
  currentProject: string | null;
  currentRateCard: string | null;
  currentView: 'project_config' | 'quotation' | 'ratecard' | 'settings' | null;
  activeQuotationFilename: string | null;
  expandedSections: Set<string>;
  expandedProject: string | null;
  expandedQuotation: string | null;
  onToggleSection: (section: string) => void;
  onSelectProjectConfig: (name: string) => void;
  onSelectRateCard: (name: string) => void;
  onSelectSettings: () => void;
  onCreateProject: (name: string) => Promise<void>;
  onCreateRateCard: (name: string) => Promise<void>;
  onToggleProject: (name: string) => void;
  onToggleQuotation: (name: string) => void;
  onQuotationSelectFile: (filename: string) => void;
  onQuotationDeleteFile: (filename: string) => Promise<void>;
  onQuotationCreate: () => Promise<void>;
}

export function Sidebar({
  projects = [],
  rateCards = [],
  quotationFiles = [],
  currentProject,
  currentRateCard,
  currentView,
  activeQuotationFilename,
  expandedSections,
  expandedProject,
  expandedQuotation,
  onToggleSection,
  onSelectProjectConfig,
  onSelectRateCard,
  onSelectSettings,
  onCreateProject,
  onCreateRateCard,
  onToggleProject,
  onToggleQuotation,
  onQuotationSelectFile,
  onQuotationDeleteFile,
  onQuotationCreate,
}: SidebarProps) {
  return (
    <aside className="w-72 p-6 bg-slate-800 text-white flex flex-col overflow-y-auto shrink-0 select-none">
      <div className="flex items-center gap-2 mb-8">
        <h2 className="text-2xl font-light tracking-wide text-blue-400">CarrotMRO</h2>
      </div>

      <SidebarSection
        title="项目列表"
        icon={expandedSections.has('projects') ? <FolderOpen size={14} /> : <Folder size={14} />}
        expanded={expandedSections.has('projects')}
        onToggle={() => onToggleSection('projects')}
        onCreate={onCreateProject}
        createPlaceholder="新项目名称..."
      >
        {projects.map(p => {
          const isProjectExpanded = expandedProject === p;
          const isCurrentProject = currentProject === p;

          return (
            <TreeItem
              key={p}
              label={p}
              active={isCurrentProject && currentView === 'project_config'}
              expandable
              expanded={isProjectExpanded}
              onToggleExpand={() => onToggleProject(p)}
              onClick={() => onSelectProjectConfig(p)}
            >
              <TreeItem
                label="项目设置"
                icon={<Sliders size={14} />}
                active={isCurrentProject && currentView === 'project_config'}
                onClick={() => onSelectProjectConfig(p)}
              />

              <TreeItem
                label="报价单列表"
                icon={<Receipt size={14} />}
                active={isCurrentProject && currentView === 'quotation'}
                expandable
                expanded={expandedQuotation === p}
                onToggleExpand={() => onToggleQuotation(p)}
                onCreate={onQuotationCreate}
              >
                {quotationFiles.map(f => (
                  <TreeItem
                    key={f}
                    label={f}
                    icon={<FileText size={14} />}
                    active={isCurrentProject && currentView === 'quotation' && activeQuotationFilename === f}
                    onClick={() => onQuotationSelectFile(f)}
                    onDelete={() => onQuotationDeleteFile(f)}
                  />
                ))}
              </TreeItem>
            </TreeItem>
          );
        })}
      </SidebarSection>

      <SidebarSection
        title="协议定价表"
        icon={expandedSections.has('ratecards') ? <FolderOpen size={14} /> : <Folder size={14} />}
        expanded={expandedSections.has('ratecards')}
        onToggle={() => onToggleSection('ratecards')}
        onCreate={onCreateRateCard}
        createPlaceholder="新定价表名称..."
      >
        {rateCards.map(rc => (
          <TreeItem
            key={rc}
            label={rc}
            active={currentRateCard === rc && currentView === 'ratecard'}
            onClick={() => onSelectRateCard(rc)}
          />
        ))}
      </SidebarSection>

      <div className="mt-auto pt-6 border-t border-slate-700">
        <button
          onClick={onSelectSettings}
          className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors ${
            currentView === 'settings'
              ? 'bg-blue-600 text-white font-medium shadow'
              : 'text-slate-300 hover:bg-slate-700 hover:text-white'
          }`}
        >
          <Settings size={16} />
          系统 LLM 设置
        </button>
      </div>
    </aside>
  );
}
