import { SidebarSection, TreeItem, Folder, FolderOpen, ScanLine, FileText, Receipt } from '@/components/SidebarTree';

interface SidebarProps {
  projects: string[];
  rateCards: string[];
  ocrFiles: string[];
  quotationFiles: string[];
  currentProject: string | null;
  currentRateCard: string | null;
  currentView: 'project' | 'ratecard' | null;
  activeFilename: string | null;
  activeQuotationFilename: string | null;
  expandedSections: Set<string>;
  expandedProject: string | null;
  expandedOcr: string | null;
  expandedQuotation: string | null;
  onToggleSection: (section: string) => void;
  onSelectProject: (name: string) => void;
  onSelectRateCard: (name: string) => void;
  onCreateProject: (name: string) => Promise<void>;
  onCreateRateCard: (name: string) => Promise<void>;
  onToggleProject: (name: string) => void;
  onToggleOcr: (name: string) => void;
  onToggleQuotation: (name: string) => void;
  onOcrSelectFile: (filename: string) => void;
  onOcrDeleteFile: (filename: string) => Promise<void>;
  onOcrUpload: (files: FileList) => void;
  onQuotationSelectFile: (filename: string) => void;
  onQuotationDeleteFile: (filename: string) => Promise<void>;
  onQuotationCreate: () => Promise<void>;
}

export function Sidebar({
  projects,
  rateCards,
  ocrFiles,
  quotationFiles,
  currentProject,
  currentRateCard,
  currentView,
  activeFilename,
  activeQuotationFilename,
  expandedSections,
  expandedProject,
  expandedOcr,
  expandedQuotation,
  onToggleSection,
  onSelectProject,
  onSelectRateCard,
  onCreateProject,
  onCreateRateCard,
  onToggleProject,
  onToggleOcr,
  onToggleQuotation,
  onOcrSelectFile,
  onOcrDeleteFile,
  onOcrUpload,
  onQuotationSelectFile,
  onQuotationDeleteFile,
  onQuotationCreate,
}: SidebarProps) {
  return (
    <aside className="w-72 p-6 bg-slate-800 text-white flex flex-col overflow-y-auto">
      <h2 className="text-2xl font-light mb-8">CarrotMRO</h2>

      <SidebarSection
        title="项目"
        icon={expandedSections.has('projects') ? <FolderOpen size={14} /> : <Folder size={14} />}
        expanded={expandedSections.has('projects')}
        onToggle={() => onToggleSection('projects')}
        onCreate={onCreateProject}
        createPlaceholder="新项目名称..."
      >
        {projects.map(p => (
          <TreeItem
            key={p}
            label={p}
            active={currentProject === p && currentView === 'project' && !activeFilename && !activeQuotationFilename}
            expandable
            expanded={expandedProject === p}
            onToggleExpand={() => onToggleProject(p)}
            onClick={() => onSelectProject(p)}
            onSettings={() => onSelectProject(p)}
          >
            <TreeItem
              label="OCR"
              icon={<ScanLine size={14} />}
              active={currentProject === p && currentView === 'project' && activeFilename !== null}
              expandable
              expanded={expandedOcr === p}
              onToggleExpand={() => onToggleOcr(p)}
              onUpload={onOcrUpload}
            >
              {ocrFiles.map(f => (
                <TreeItem
                  key={f}
                  label={f}
                  icon={<FileText size={14} />}
                  active={activeFilename === f}
                  onClick={() => onOcrSelectFile(f)}
                  onDelete={() => onOcrDeleteFile(f)}
                />
              ))}
            </TreeItem>
            <TreeItem
              label="报价"
              icon={<Receipt size={14} />}
              active={currentProject === p && currentView === 'project' && activeQuotationFilename !== null}
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
                  active={activeQuotationFilename === f}
                  onClick={() => onQuotationSelectFile(f)}
                  onDelete={() => onQuotationDeleteFile(f)}
                />
              ))}
            </TreeItem>
          </TreeItem>
        ))}
      </SidebarSection>

      <SidebarSection
        title="协议基准价格清单"
        icon={expandedSections.has('ratecards') ? <FolderOpen size={14} /> : <Folder size={14} />}
        expanded={expandedSections.has('ratecards')}
        onToggle={() => onToggleSection('ratecards')}
        onCreate={onCreateRateCard}
        createPlaceholder="新清单名称..."
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
    </aside>
  );
}
