import { useState, useEffect } from 'react';
import * as api from '@/api';
import { RateCardTableData, QuotationItem, ProjectSettings } from '@/types';
import { Sidebar } from '@/components/Sidebar';
import { ProjectConfigWorkspace } from '@/components/ProjectConfigWorkspace';
import { RateCardWorkspace } from '@/components/RateCardWorkspace';
import { QuotationWorkspace } from '@/components/QuotationWorkspace';
import { SettingsWorkspace } from '@/components/SettingsWorkspace';

function App() {
  const [projects, setProjects] = useState<string[]>([]);
  const [rateCards, setRateCards] = useState<string[]>([]);
  const [templates, setTemplates] = useState<string[]>([]);
  const [quotationFiles, setQuotationFiles] = useState<string[]>([]);

  const [currentProject, setCurrentProject] = useState<string | null>(null);
  const [currentRateCard, setCurrentRateCard] = useState<string | null>(null);
  const [activeQuotationFilename, setActiveQuotationFilename] = useState<string | null>(null);

  const [currentView, setCurrentView] = useState<'project_config' | 'quotation' | 'ratecard' | 'settings' | null>(null);

  const [projectSettings, setProjectSettings] = useState<ProjectSettings>({
    name: '',
    created_at: '',
    ratecard_name: null,
    template_name: null,
    ocr_columns: [],
    quotation_columns: [],
  });

  const [quotationItems, setQuotationItems] = useState<QuotationItem[]>([]);
  const [quotationRemarks, setQuotationRemarks] = useState<string[]>([]);
  const [ratecardTableData, setRatecardTableData] = useState<RateCardTableData>({ columns: [], items: [] });

  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set(['projects', 'ratecards']));
  const [expandedProject, setExpandedProject] = useState<string | null>(null);
  const [expandedQuotation, setExpandedQuotation] = useState<string | null>(null);

  const fetchProjects = async () => {
    try {
      const res = await api.getProjects();
      setProjects(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('获取项目列表失败', err);
    }
  };

  const fetchRateCards = async () => {
    try {
      const res = await api.getRateCards();
      setRateCards(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('获取定价表列表失败', err);
    }
  };

  const fetchTemplates = async () => {
    try {
      const res = await api.getTemplates();
      setTemplates(Array.isArray(res.data) ? res.data : []);
    } catch (err) {
      console.error('获取模板列表失败', err);
    }
  };

  const fetchQuotationFiles = async (projectName: string) => {
    try {
      const res = await api.getQuotations(projectName);
      setQuotationFiles(Array.isArray(res.data.files) ? res.data.files : []);
    } catch (err) {
      console.error('获取报价单列表失败', err);
    }
  };

  const loadProjectInfo = async (projectName: string) => {
    try {
      const res = await api.getProjectInfo(projectName);
      setProjectSettings(res.data);
    } catch (err) {
      console.error('加载项目 Settings 失败', err);
    }
  };

  const refreshRateCardData = async (rcName: string) => {
    try {
      const res = await api.getRateCardData(rcName);
      setRatecardTableData({
        columns: Array.isArray(res.data.columns) ? res.data.columns : [],
        items: Array.isArray(res.data.items) ? res.data.items : [],
      });
    } catch {
      setRatecardTableData({ columns: [], items: [] });
    }
  };

  useEffect(() => {
    fetchProjects();
    fetchRateCards();
    fetchTemplates();
  }, []);

  useEffect(() => {
    if (currentProject) {
      fetchQuotationFiles(currentProject);
      loadProjectInfo(currentProject);
      setExpandedProject(currentProject);
      setExpandedQuotation(currentProject);
    }
  }, [currentProject]);

  useEffect(() => {
    if (currentRateCard) {
      refreshRateCardData(currentRateCard);
    }
  }, [currentRateCard]);

  // 选择切换为项目配置视图
  const handleSelectProjectConfig = (projectName: string) => {
    setCurrentProject(projectName);
    setCurrentRateCard(null);
    setActiveQuotationFilename(null);
    setCurrentView('project_config');
    loadProjectInfo(projectName);
  };

  // 切换选中特定报价单
  const handleQuotationSelectFile = async (filename: string) => {
    if (!currentProject) return;
    try {
      const res = await api.getQuotationData(currentProject, filename);
      const items: QuotationItem[] = Array.isArray(res.data.items) ? res.data.items : [];
      const remarks: string[] = Array.isArray(res.data.remarks)
        ? res.data.remarks
        : (res.data.remarks ? [String(res.data.remarks)] : []);

      setQuotationItems(items);
      setQuotationRemarks(remarks);
      setActiveQuotationFilename(filename);
      setCurrentRateCard(null);
      setCurrentView('quotation');

      // 预加载项目关联的定价表数据用于弹窗检索
      if (projectSettings.ratecard_name) {
        refreshRateCardData(projectSettings.ratecard_name);
      }
    } catch (err) {
      alert('读取报价单失败: ' + String(err));
    }
  };

  // 更新项目设置
  const handleUpdateProjectSettings = async (updated: Partial<ProjectSettings>) => {
    if (!currentProject) return;
    try {
      const res = await api.updateProjectSettings(currentProject, updated);
      setProjectSettings(res.data.settings || { ...projectSettings, ...updated });
    } catch (err: any) {
      alert('更新项目设置失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  // 创建新报价单
  const handleQuotationCreate = async () => {
    if (!currentProject) return;
    try {
      const res = await api.createQuotation(currentProject);
      fetchQuotationFiles(currentProject);
      handleQuotationSelectFile(res.data.file);
    } catch (err: any) {
      alert('创建报价单失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  // 删除报价单
  const handleQuotationDeleteFile = async (filename: string) => {
    if (!currentProject) return;
    if (!confirm(`确定删除报价单 ${filename} 吗？`)) return;
    try {
      await api.deleteQuotation(currentProject, filename);
      fetchQuotationFiles(currentProject);
      if (activeQuotationFilename === filename) {
        setActiveQuotationFilename(null);
        setCurrentView('project_config');
      }
    } catch (err) {
      alert('删除报价单失败: ' + String(err));
    }
  };

  // 报价单编辑相关
  const handleQuotationEdit = (index: number, field: string, value: string) => {
    if (index < 0 || index >= quotationItems.length) return;
    const next = [...quotationItems];
    next[index] = { ...next[index], [field]: value };
    setQuotationItems(next);
  };

  const handleQuotationAddRow = (index?: number) => {
    const emptyRow: QuotationItem = { _matchStatus: 'pending' };
    const next = [...quotationItems];
    const insertAt = index !== undefined ? index + 1 : next.length;
    next.splice(insertAt, 0, emptyRow);
    setQuotationItems(next);
  };

  const handleQuotationDeleteRow = (index: number) => {
    const next = quotationItems.filter((_, i) => i !== index);
    setQuotationItems(next);
  };

  const handleQuotationSave = async () => {
    if (!currentProject || !activeQuotationFilename) return;
    try {
      await api.saveQuotationData(currentProject, activeQuotationFilename, { items: quotationItems });
      alert('报价单保存成功！');
    } catch {
      alert('保存报价单失败');
    }
  };

  // 模板上传与删除
  const handleUploadTemplate = async (file: File) => {
    const formData = new FormData();
    formData.append('file', file);
    try {
      await api.uploadTemplate(formData);
      fetchTemplates();
    } catch (err: any) {
      alert('上传模板失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  const handleDeleteTemplate = async (filename: string) => {
    if (!confirm(`确定删除模板 ${filename} 吗？`)) return;
    try {
      await api.deleteTemplate(filename);
      fetchTemplates();
      if (projectSettings.template_name === filename) {
        handleUpdateProjectSettings({ template_name: null });
      }
    } catch (err: any) {
      alert('删除模板失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  // 报价单重命名
  const handleQuotationRenameFile = async (oldFilename: string, newFilename: string) => {
    if (!currentProject) return;
    try {
      const res = await api.renameQuotation(currentProject, oldFilename, newFilename);
      const updatedFilename = res.data.file;
      await fetchQuotationFiles(currentProject);
      if (activeQuotationFilename === oldFilename) {
        setActiveQuotationFilename(updatedFilename);
      }
    } catch (err: any) {
      alert('重命名报价单失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  // 定价表重命名
  const handleRateCardRename = async (oldName: string, newName: string) => {
    try {
      const res = await api.renameRateCard(oldName, newName);
      const updatedName = res.data.name;
      await fetchRateCards();
      if (currentRateCard === oldName) {
        setCurrentRateCard(updatedName);
      }
      if (currentProject) {
        await loadProjectInfo(currentProject);
      }
    } catch (err: any) {
      alert('重命名定价表失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  // 定价表删除
  const handleRateCardDelete = async (name: string) => {
    if (!confirm(`确定删除定价表 ${name} 吗？`)) return;
    try {
      await api.deleteRateCard(name);
      await fetchRateCards();
      if (currentRateCard === name) {
        setCurrentRateCard(null);
        setCurrentView(null);
      }
      if (currentProject) {
        await loadProjectInfo(currentProject);
      }
    } catch (err: any) {
      alert('删除定价表失败: ' + (err.response?.data?.detail || err.message));
    }
  };

  const handleToggleSection = (section: string) => {
    setExpandedSections(prev => {
      const next = new Set(prev);
      next.has(section) ? next.delete(section) : next.add(section);
      return next;
    });
  };

  return (
    <div className="flex h-screen font-sans text-gray-800 bg-gray-100 overflow-hidden">
      <Sidebar
        projects={projects}
        rateCards={rateCards}
        quotationFiles={quotationFiles}
        currentProject={currentProject}
        currentRateCard={currentRateCard}
        currentView={currentView}
        activeQuotationFilename={activeQuotationFilename}
        expandedSections={expandedSections}
        expandedProject={expandedProject}
        expandedQuotation={expandedQuotation}
        onToggleSection={handleToggleSection}
        onSelectProjectConfig={handleSelectProjectConfig}
        onSelectRateCard={(name) => {
          setCurrentRateCard(name);
          setActiveQuotationFilename(null);
          setCurrentView('ratecard');
        }}
        onSelectSettings={() => {
          setCurrentRateCard(null);
          setActiveQuotationFilename(null);
          setCurrentView('settings');
        }}
        onCreateProject={async (name) => {
          await api.createProject(name);
          await fetchProjects();
          handleSelectProjectConfig(name);
        }}
        onCreateRateCard={async (name) => {
          await api.createRateCard(name);
          await fetchRateCards();
          setCurrentRateCard(name);
          setCurrentView('ratecard');
        }}
        onToggleProject={(name) => {
          if (currentProject === name && currentView === 'project_config') {
            setExpandedProject(expandedProject === name ? null : name);
          } else {
            handleSelectProjectConfig(name);
            setExpandedProject(name);
          }
        }}
        onToggleQuotation={(name) => setExpandedQuotation(expandedQuotation === name ? null : name)}
        onQuotationSelectFile={handleQuotationSelectFile}
        onQuotationDeleteFile={handleQuotationDeleteFile}
        onQuotationRenameFile={handleQuotationRenameFile}
        onRateCardRename={handleRateCardRename}
        onRateCardDelete={handleRateCardDelete}
        onQuotationCreate={handleQuotationCreate}
      />

      <main className="flex-1 min-w-0 p-8 overflow-y-auto">
        {currentView === 'project_config' && currentProject && (
          <ProjectConfigWorkspace
            currentProject={currentProject}
            settings={projectSettings}
            rateCards={rateCards}
            templates={templates}
            onUpdateSettings={handleUpdateProjectSettings}
            onUploadTemplate={handleUploadTemplate}
            onDeleteTemplate={handleDeleteTemplate}
          />
        )}

        {currentView === 'quotation' && currentProject && activeQuotationFilename && (
          <QuotationWorkspace
            currentProject={currentProject}
            activeQuotationFilename={activeQuotationFilename}
            quotationItems={quotationItems}
            quotationRemarks={quotationRemarks}
            projectRateCard={projectSettings.ratecard_name}
            projectTemplate={projectSettings.template_name}
            quotationColumns={projectSettings.quotation_columns}
            matchValidationRules={projectSettings.match_validation_rules}
            onEdit={handleQuotationEdit}
            onAddRow={handleQuotationAddRow}
            onDeleteRow={handleQuotationDeleteRow}
            onSave={handleQuotationSave}
            onQuotationDataChange={setQuotationItems}
            onQuotationRemarksChange={setQuotationRemarks}
          />
        )}

        {currentView === 'ratecard' && currentRateCard && (
          <RateCardWorkspace
            currentRateCard={currentRateCard}
            ratecardTableData={ratecardTableData}
            onRefreshData={() => refreshRateCardData(currentRateCard)}
          />
        )}

        {currentView === 'settings' && (
          <SettingsWorkspace />
        )}

        {!currentView && (
          <div className="flex flex-col items-center justify-center h-full text-gray-400">
            <h2 className="text-2xl font-light mb-2">欢迎使用 CarrotMRO 管理系统</h2>
            <p className="text-sm">请从左侧选择一个项目或协议定价表开始工作</p>
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
