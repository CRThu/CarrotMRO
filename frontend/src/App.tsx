import { useState, useEffect, useRef } from 'react';
import * as api from '@/api';
import { OcrTableData, RateCardTableData } from '@/types';
import { Sidebar } from '@/components/Sidebar';
import { ProjectWorkspace } from '@/components/ProjectWorkspace';
import { RateCardWorkspace } from '@/components/RateCardWorkspace';

function App() {
  const [projects, setProjects] = useState<string[]>([]);
  const [currentProject, setCurrentProject] = useState<string | null>(null);
  const [projectRateCard, setProjectRateCard] = useState<string | null>(null);
  const [rateCards, setRateCards] = useState<string[]>([]);
  const [currentRateCard, setCurrentRateCard] = useState<string | null>(null);
  const [ocrFiles, setOcrFiles] = useState<string[]>([]);
  const [activeFilename, setActiveFilename] = useState<string | null>(null);
  const [tableData, setTableData] = useState<OcrTableData>({ columns: [], items: [], remarks: '' });
  const [loading, setLoading] = useState(false);
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set(['projects', 'ratecards']));
  const [expandedProject, setExpandedProject] = useState<string | null>(null);
  const [expandedOcr, setExpandedOcr] = useState<string | null>(null);
  const [currentView, setCurrentView] = useState<'project' | 'ratecard' | null>(null);
  const [ratecardTableData, setRatecardTableData] = useState<RateCardTableData>({ columns: [], items: [] });
  const [ratecardImporting, setRatecardImporting] = useState(false);
  const pollingRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const fetchProjects = async () => { try { const res = await api.getProjects(); setProjects(res.data.projects); } catch (err) { console.error(err); } };
  const fetchRateCards = async () => { try { const res = await api.getRateCards(); setRateCards(res.data.ratecards); } catch (err) { console.error(err); } };
  const fetchOcrFiles = async (projectName: string) => { try { const res = await api.getOcrFiles(projectName); setOcrFiles(res.data.files); } catch (err) { console.error(err); } };

  useEffect(() => { fetchProjects(); fetchRateCards(); }, []);

  useEffect(() => {
    if (currentProject) {
      fetchOcrFiles(currentProject);
      setExpandedProject(currentProject);
      setExpandedOcr(currentProject);
      api.getProjectInfo(currentProject).then(res => setProjectRateCard(res.data.ratecard_name)).catch((err: any) => { alert("获取项目信息失败: " + (err?.message || String(err))); });
      setTableData({ columns: [], items: [], remarks: '' });
    }
  }, [currentProject]);

  useEffect(() => {
    if (currentRateCard) {
      api.getRateCardData(currentRateCard).then(res => {
        const d = res.data;
        setRatecardTableData({
          columns: Array.isArray(d.columns) ? d.columns : [],
          items: Array.isArray(d.items) ? d.items : [],
        });
      }).catch(() => setRatecardTableData({ columns: [], items: [] }));
    }
  }, [currentRateCard]);

  useEffect(() => {
    return () => { if (pollingRef.current) { clearInterval(pollingRef.current); pollingRef.current = null; } };
  }, []);

  const handleEdit = (index: number, field: string, value: string) => {
    const items = tableData?.items;
    if (!items || index < 0 || index >= items.length) return;
    const newItems = [...items];
    newItems[index] = { ...newItems[index], [field]: value };
    setTableData({ ...tableData, items: newItems });
  };

  const handleSelectFile = async (filename: string) => {
    try {
      const res = await api.getOcrData(currentProject!, filename);
      const fileData = res.data.data || res.data;
      setTableData({
        columns: Array.isArray(fileData.columns) ? fileData.columns : [],
        items: Array.isArray(fileData.items) ? fileData.items : [],
        remarks: fileData.remarks || ''
      });
      setActiveFilename(filename);
    } catch (err) { alert("读取数据失败"); }
  };

  const handleSave = async () => {
    if (!activeFilename) return;
    try {
      await api.saveOcrData(currentProject!, activeFilename, tableData);
      alert("保存成功！");
    } catch (err) { alert("保存失败"); }
  };

  const handleOcrUpload = async (files: FileList) => {
    setLoading(true);
    const formData = new FormData();
    for (let i = 0; i < files.length; i++) formData.append('files', files[i]);

    try {
      const res = await api.uploadOcrFiles(currentProject!, formData);
      const { task_id } = res.data;

      pollingRef.current = setInterval(async () => {
        try {
          const statusRes = await api.checkTaskStatus(task_id);

          if (statusRes.data.status === 'done') {
            clearInterval(pollingRef.current!);
            pollingRef.current = null;
            setActiveFilename(statusRes.data.file);

            const rawResult = statusRes.data.result || {};
            const actualData = rawResult.data || rawResult;

            if (rawResult.success === false) {
              setLoading(false);
              alert("识别失败: " + (rawResult.error || "未知错误"));
              fetchOcrFiles(currentProject!);
              return;
            }

            setTableData({
              columns: Array.isArray(statusRes.data.columns) ? statusRes.data.columns : [],
              items: Array.isArray(actualData.items) ? actualData.items : [],
              remarks: actualData.remarks || ''
            });

            fetchOcrFiles(currentProject!);
            setLoading(false);
          } else if (statusRes.data.status === 'error') {
            clearInterval(pollingRef.current!);
            pollingRef.current = null;
            setLoading(false);
            alert("识别失败: " + statusRes.data.message);
          }
        } catch (pollErr) {
          clearInterval(pollingRef.current!);
          pollingRef.current = null;
          setLoading(false);
          alert("轮询状态失败: " + String(pollErr));
        }
      }, 2000);
    } catch (uploadErr) {
      setLoading(false);
      alert("上传失败: " + String(uploadErr));
    }
  };

  const handleOcrDeleteFile = async (filename: string) => {
    if (!confirm(`确定删除 ${filename} 吗？`)) return;
    try {
      await api.deleteOcrFile(currentProject!, filename);
      fetchOcrFiles(currentProject!);
      if (activeFilename === filename) {
        setActiveFilename(null);
      setTableData({ columns: [], items: [], remarks: '' });
      }
    } catch (err) {
      alert('删除失败');
    }
  };

  const handleRateCardEdit = (index: number, field: string, value: string) => {
    const items = ratecardTableData?.items;
    if (!items || index < 0 || index >= items.length) return;
    const newItems = [...items];
    newItems[index] = { ...newItems[index], [field]: value };
    setRatecardTableData({ ...ratecardTableData, items: newItems });
  };

  const handleRateCardImport = async (file: File) => {
    setRatecardImporting(true);
    const formData = new FormData();
    formData.append('file', file);
    try {
      const res = await api.importRateCardFile(currentRateCard!, formData);
      const d = res.data;
      setRatecardTableData({
        columns: Array.isArray(d.columns) ? d.columns : [],
        items: Array.isArray(d.items) ? d.items : [],
      });
    } catch (err: any) {
      alert("导入失败: " + (err.response?.data?.detail || err.message));
    }
    setRatecardImporting(false);
  };

  const handleToggleSection = (section: string) => {
    setExpandedSections(prev => {
      const next = new Set(prev);
      next.has(section) ? next.delete(section) : next.add(section);
      return next;
    });
  };

  const handleSelectProject = (name: string) => {
    setCurrentProject(name);
    setCurrentRateCard(null);
    setCurrentView('project');
    setActiveFilename(null);
  };

  const handleToggleProject = (name: string) => {
    if (currentProject === name) {
      setExpandedProject(expandedProject === name ? null : name);
    } else {
      handleSelectProject(name);
    }
  };

  return (
    <div className="flex h-screen font-sans text-gray-800">
      <Sidebar
        projects={projects}
        rateCards={rateCards}
        ocrFiles={ocrFiles}
        currentProject={currentProject}
        currentRateCard={currentRateCard}
        currentView={currentView}
        activeFilename={activeFilename}
        expandedSections={expandedSections}
        expandedProject={expandedProject}
        expandedOcr={expandedOcr}
        onToggleSection={handleToggleSection}
        onSelectProject={handleSelectProject}
        onSelectRateCard={(name) => { setCurrentRateCard(name); setCurrentProject(null); setCurrentView('ratecard'); }}
        onCreateProject={async (name) => { await api.createProject(name); fetchProjects(); }}
        onCreateRateCard={async (name) => { await api.createRateCard(name); fetchRateCards(); }}
        onToggleProject={handleToggleProject}
        onToggleOcr={(name) => setExpandedOcr(expandedOcr === name ? null : name)}
        onOcrSelectFile={handleSelectFile}
        onOcrDeleteFile={handleOcrDeleteFile}
        onOcrUpload={handleOcrUpload}
      />

      <main className="flex-1 p-10 bg-gray-100 overflow-y-auto">
        {currentView === 'project' && currentProject && (
          <ProjectWorkspace
            currentProject={currentProject}
            projectRateCard={projectRateCard}
            rateCards={rateCards}
            activeFilename={activeFilename}
            loading={loading}
            tableData={tableData}
            onEdit={handleEdit}
            onSave={handleSave}
            onUpdateRateCard={async (name) => {
              await api.updateProjectRateCard(currentProject!, name);
              setProjectRateCard(name);
            }}
          />
        )}

        {currentView === 'ratecard' && currentRateCard && (
          <RateCardWorkspace
            currentRateCard={currentRateCard}
            ratecardTableData={ratecardTableData}
            importing={ratecardImporting}
            onEdit={handleRateCardEdit}
            onImport={handleRateCardImport}
          />
        )}

        {!currentView && (
          <h1 className="text-3xl font-light mb-8 text-gray-700">请从左侧选择一个项目或协议定价表</h1>
        )}
      </main>
    </div>
  );
}

export default App;
