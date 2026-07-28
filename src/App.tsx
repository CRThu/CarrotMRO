import { useState, useEffect, useRef } from 'react';
import * as api from '@/api';
import { OcrTableData, RateCardTableData, RateCardColumn, TableItem, QuotationItem, PresetColumn, ColumnMappings, MappingScope } from '@/types';
import { Sidebar } from '@/components/Sidebar';
import { ProjectWorkspace } from '@/components/ProjectWorkspace';
import { ProjectConfigWorkspace } from '@/components/ProjectConfigWorkspace';
import { RateCardWorkspace } from '@/components/RateCardWorkspace';
import { QuotationWorkspace } from '@/components/QuotationWorkspace';
import { TaskNotification } from '@/components/TaskNotification';

function App() {
  const [projects, setProjects] = useState<string[]>([]);
  const [currentProject, setCurrentProject] = useState<string | null>(null);
  const [projectRateCard, setProjectRateCard] = useState<string | null>(null);
  const [rateCards, setRateCards] = useState<string[]>([]);
  const [currentRateCard, setCurrentRateCard] = useState<string | null>(null);
  const [ocrFiles, setOcrFiles] = useState<string[]>([]);
  const [quotationFiles, setQuotationFiles] = useState<string[]>([]);
  const [activeFilename, setActiveFilename] = useState<string | null>(null);
  const [activeQuotationFilename, setActiveQuotationFilename] = useState<string | null>(null);
  const [tableData, setTableData] = useState<OcrTableData>({ columns: [], items: [], remarks: '' });
  const [expandedSections, setExpandedSections] = useState<Set<string>>(new Set(['projects', 'ratecards']));
  const [expandedProject, setExpandedProject] = useState<string | null>(null);
  const [expandedOcr, setExpandedOcr] = useState<string | null>(null);
  const [expandedQuotation, setExpandedQuotation] = useState<string | null>(null);
  const [currentView, setCurrentView] = useState<'project' | 'ratecard' | null>(null);
  const [ratecardTableData, setRatecardTableData] = useState<RateCardTableData>({ columns: [], items: [] });
  const [ratecardImporting, setRatecardImporting] = useState(false);
  const [quotationData, setQuotationData] = useState<{ columns: RateCardColumn[]; items: QuotationItem[] }>({ columns: [], items: [] });
  const [ocrTaskStatus, setOcrTaskStatus] = useState<{ status: 'processing' | 'done' | 'error'; message?: string } | null>(null);
  const pollingRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const [projectTemplate, setProjectTemplate] = useState<string | null>(null);
  const [templates, setTemplates] = useState<string[]>([]);
  const [availableColumns, setAvailableColumns] = useState<string[]>([]);
  const [selectedColumns, setSelectedColumns] = useState<string[]>([]);
  const [presetColumns, setPresetColumns] = useState<PresetColumn[]>([]);
  const [columnMappings, setColumnMappings] = useState<ColumnMappings>({ ocr: {}, ratecard: {}, quotation: {} });

  const toColumns = (cols: any): RateCardColumn[] => {
    if (!Array.isArray(cols)) return [];
    return cols.map(c => typeof c === 'string' ? { name: c, strict: false, alias: null } : c);
  };

  const fetchProjects = async () => { try { const res = await api.getProjects(); setProjects(Array.isArray(res.data) ? res.data : (res.data?.projects || [])); } catch (err) { console.error(err); } };
  const fetchRateCards = async () => { try { const res = await api.getRateCards(); setRateCards(Array.isArray(res.data) ? res.data : (res.data?.ratecards || [])); } catch (err) { console.error(err); } };
  const fetchTemplates = async () => { try { const res = await api.getTemplates(); setTemplates(Array.isArray(res.data) ? res.data : (res.data?.files || [])); } catch (err) { console.error(err); } };
  const fetchOcrFiles = async (projectName: string) => { try { const res = await api.getOcrFiles(projectName); setOcrFiles(Array.isArray(res.data) ? res.data : (res.data?.files || [])); } catch (err) { console.error(err); } };
  const fetchQuotationFiles = async (projectName: string) => { try { const res = await api.getQuotations(projectName); setQuotationFiles(Array.isArray(res.data) ? res.data : (res.data?.files || [])); } catch (err) { console.error(err); } };

  useEffect(() => { fetchProjects(); fetchRateCards(); fetchTemplates(); fetchPresetColumns(); }, []);

  const fetchPresetColumns = async () => { try { const res = await api.getPresetColumns(); setPresetColumns(Array.isArray(res.data) ? res.data : (res.data?.columns || [])); } catch (err) { console.error(err); } };

  useEffect(() => {
    if (currentProject) {
      fetchOcrFiles(currentProject);
      fetchQuotationFiles(currentProject);
      setExpandedProject(currentProject);
      setExpandedOcr(currentProject);
      setExpandedQuotation(currentProject);
      setProjectTemplate(null);
      setAvailableColumns([]);
      setSelectedColumns([]);
      setColumnMappings({ ocr: {}, ratecard: {}, quotation: {} });
      api.getProjectInfo(currentProject).then(res => {
        setProjectRateCard(res.data.ratecard_name);
        setProjectTemplate(res.data.template_name);
      }).catch((err: any) => { alert("获取项目信息失败: " + (err?.message || String(err))); });
      api.getProjectColumns(currentProject).then(res => {
        setAvailableColumns(res.data.available_columns || []);
        setSelectedColumns(res.data.selected_columns || []);
        setColumnMappings(res.data.column_mappings || { ocr: {}, ratecard: {}, quotation: {} });
      }).catch(() => { setAvailableColumns([]); setSelectedColumns([]); setColumnMappings({ ocr: {}, ratecard: {}, quotation: {} }); });
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
    if (currentView === 'project' && projectRateCard) {
      api.getRateCardData(projectRateCard).then(res => {
        const d = res.data;
        setRatecardTableData({
          columns: Array.isArray(d.columns) ? d.columns : [],
          items: Array.isArray(d.items) ? d.items : [],
        });
      }).catch(() => setRatecardTableData({ columns: [], items: [] }));
    }
  }, [currentView, projectRateCard]);

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

  const handleAddRow = (index?: number) => {
    const emptyRow: Record<string, string> = {};
    tableData.columns.forEach((col) => { emptyRow[col.name] = ''; });
    const newItems = [...tableData.items];
    const insertAt = index !== undefined ? index + 1 : newItems.length;
    newItems.splice(insertAt, 0, emptyRow);
    setTableData({ ...tableData, items: newItems });
  };

  const handleDeleteRow = (index: number) => {
    const newItems = tableData.items.filter((_, i) => i !== index);
    setTableData({ ...tableData, items: newItems });
  };

  const handleSelectFile = async (filename: string) => {
    try {
      const res = await api.getOcrData(currentProject!, filename);
      const raw = res.data;
      const fileData = raw.data || raw;
      setTableData({
        columns: toColumns(raw.columns || fileData.columns),
        items: Array.isArray(fileData.items) ? fileData.items : [],
        remarks: fileData.remarks || ''
      });
      setActiveFilename(filename);
      setActiveQuotationFilename(null);
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
    if (!selectedColumns || selectedColumns.length === 0) {
      alert('请先在项目配置中关联模板并选择识别列');
      return;
    }
    setOcrTaskStatus({ status: 'processing' });
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
            setActiveQuotationFilename(null);

            const rawResult = statusRes.data.result || {};
            const actualData = rawResult.data || rawResult;

            if (rawResult.success === false) {
              setOcrTaskStatus({ status: 'error', message: rawResult.error || '未知错误' });
              fetchOcrFiles(currentProject!);
              return;
            }

            setTableData({
              columns: toColumns(statusRes.data.columns),
              items: Array.isArray(actualData.items) ? actualData.items : [],
              remarks: actualData.remarks || ''
            });

            fetchOcrFiles(currentProject!);
            setOcrTaskStatus({ status: 'done' });
          } else if (statusRes.data.status === 'error') {
            clearInterval(pollingRef.current!);
            pollingRef.current = null;
            setOcrTaskStatus({ status: 'error', message: statusRes.data.message });
          }
        } catch (pollErr) {
          clearInterval(pollingRef.current!);
          pollingRef.current = null;
          setOcrTaskStatus({ status: 'error', message: '轮询状态失败: ' + String(pollErr) });
        }
      }, 2000);
    } catch (uploadErr) {
      setOcrTaskStatus({ status: 'error', message: '上传失败: ' + String(uploadErr) });
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

  const handleQuotationCreate = async () => {
    if (!currentProject) return;
    try {
      const res = await api.createQuotation(currentProject);
      fetchQuotationFiles(currentProject);
      setActiveQuotationFilename(res.data.file);
      setActiveFilename(null);
      setExpandedQuotation(currentProject);
    } catch (err) {
      alert('创建报价单失败');
    }
  };

  const handleQuotationSelectFile = async (filename: string) => {
    if (!currentProject) return;
    try {
      const res = await api.getQuotationData(currentProject, filename);
      const data = res.data;
      const items: QuotationItem[] = (Array.isArray(data.items) ? data.items : []).map((item: TableItem) => ({
        ...item,
        _matchStatus: item._matchStatus || 'pending',
        '清单名称': item['清单名称'] || '',
      }));
      setQuotationData({
        columns: Array.isArray(data.columns) ? data.columns : [],
        items,
      });
      setActiveQuotationFilename(filename);
      setActiveFilename(null);
    } catch (err) {
      alert('读取报价单数据失败');
    }
  };

  const handleQuotationDeleteFile = async (filename: string) => {
    if (!currentProject) return;
    if (!confirm(`确定删除 ${filename} 吗？`)) return;
    try {
      await api.deleteQuotation(currentProject, filename);
      fetchQuotationFiles(currentProject);
      if (activeQuotationFilename === filename) {
        setActiveQuotationFilename(null);
        setQuotationData({ columns: [], items: [] });
      }
    } catch (err) {
      alert('删除报价单失败');
    }
  };

  const handleQuotationEdit = (index: number, field: string, value: string) => {
    const items = quotationData.items;
    if (index < 0 || index >= items.length) return;
    const newItems = [...items];
    newItems[index] = { ...newItems[index], [field]: value };
    setQuotationData({ ...quotationData, items: newItems });
  };

  const handleQuotationAddRow = (index?: number) => {
    const emptyRow: QuotationItem = { _matchStatus: 'pending', '清单名称': '' };
    quotationData.columns.forEach((col) => { emptyRow[col.name] = ''; });
    const newItems = [...quotationData.items];
    const insertAt = index !== undefined ? index + 1 : newItems.length;
    newItems.splice(insertAt, 0, emptyRow);
    setQuotationData({ ...quotationData, items: newItems });
  };

  const handleQuotationDeleteRow = (index: number) => {
    const newItems = quotationData.items.filter((_, i) => i !== index);
    setQuotationData({ ...quotationData, items: newItems });
  };

  const handleQuotationSave = async () => {
    if (!activeQuotationFilename || !currentProject) return;
    try {
      await api.saveQuotationData(currentProject, activeQuotationFilename, quotationData);
      alert('报价单保存成功！');
    } catch {
      alert('报价单保存失败');
    }
  };

  const handleQuotationDataChange = (items: QuotationItem[]) => {
    setQuotationData({ ...quotationData, items });
  };

  const handleUpdateProjectTemplate = async (templateName: string) => {
    if (!currentProject) return;
    try {
      await api.updateProjectTemplate(currentProject, templateName || null);
      setProjectTemplate(templateName || null);
      const colRes = await api.getProjectColumns(currentProject);
      setAvailableColumns(colRes.data.available_columns || []);
      setSelectedColumns(colRes.data.selected_columns || []);
    } catch (err: any) {
      alert("模板关联失败: " + (err?.response?.data?.detail || err.message));
    }
  };

  const handleUpdateSelectedColumns = async (columns: string[]) => {
    if (!currentProject) return;
    try {
      await api.updateProjectColumns(currentProject, { columns });
      setSelectedColumns(columns);
    } catch (err: any) {
      alert("列配置失败: " + (err?.response?.data?.detail || err.message));
    }
  };

  const handleUpdateColumnMapping = async (scope: MappingScope, mapping: Record<string, string>) => {
    if (!currentProject) return;
    try {
      await api.updateProjectColumns(currentProject, { columns: selectedColumns, scope, column_mapping: mapping });
      setColumnMappings(prev => ({ ...prev, [scope]: mapping }));
    } catch (err: any) {
      alert("列映射失败: " + (err?.response?.data?.detail || err.message));
    }
  };

  const handleUploadTemplate = async (file: File) => {
    const formData = new FormData();
    formData.append('file', file);
    try {
      await api.uploadTemplate(formData);
      fetchTemplates();
    } catch (err: any) {
      alert("模板上传失败: " + (err?.response?.data?.detail || err.message));
    }
  };

  const handleDeleteTemplate = async (filename: string) => {
    if (!confirm(`确定删除模板 ${filename} 吗？`)) return;
    try {
      await api.deleteTemplate(filename);
      fetchTemplates();
      if (projectTemplate === filename) {
        handleUpdateProjectTemplate('');
      }
    } catch (err: any) {
      alert("删除失败: " + (err?.response?.data?.detail || err.message));
    }
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
    setActiveQuotationFilename(null);
    setQuotationData({ columns: [], items: [] });
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
        quotationFiles={quotationFiles}
        currentProject={currentProject}
        currentRateCard={currentRateCard}
        currentView={currentView}
        activeFilename={activeFilename}
        activeQuotationFilename={activeQuotationFilename}
        expandedSections={expandedSections}
        expandedProject={expandedProject}
        expandedOcr={expandedOcr}
        expandedQuotation={expandedQuotation}
        onToggleSection={handleToggleSection}
        onSelectProject={handleSelectProject}
        onSelectRateCard={(name) => { setCurrentRateCard(name); setCurrentProject(null); setCurrentView('ratecard'); }}
        onCreateProject={async (name) => { await api.createProject(name); fetchProjects(); }}
        onCreateRateCard={async (name) => { await api.createRateCard(name); fetchRateCards(); }}
        onToggleProject={handleToggleProject}
        onToggleOcr={(name) => setExpandedOcr(expandedOcr === name ? null : name)}
        onToggleQuotation={(name) => setExpandedQuotation(expandedQuotation === name ? null : name)}
        onOcrSelectFile={handleSelectFile}
        onOcrDeleteFile={handleOcrDeleteFile}
        onOcrUpload={handleOcrUpload}
        onQuotationSelectFile={handleQuotationSelectFile}
        onQuotationDeleteFile={handleQuotationDeleteFile}
        onQuotationCreate={handleQuotationCreate}
      />

      <main className="flex-1 p-10 bg-gray-100 overflow-y-auto">
        {currentView === 'project' && currentProject && activeQuotationFilename && (
          <QuotationWorkspace
            currentProject={currentProject}
            activeQuotationFilename={activeQuotationFilename}
            quotationItems={quotationData.items}
            ocrFiles={ocrFiles}
            projectRateCard={projectRateCard}
            ratecardTableData={ratecardTableData}
            selectedColumns={selectedColumns}
            quotationMapping={columnMappings.quotation || {}}
            onEdit={handleQuotationEdit}
            onAddRow={handleQuotationAddRow}
            onDeleteRow={handleQuotationDeleteRow}
            onSave={handleQuotationSave}
            onQuotationDataChange={handleQuotationDataChange}
          />
        )}

        {currentView === 'project' && currentProject && activeFilename && !activeQuotationFilename && (
          <ProjectWorkspace
            currentProject={currentProject}
            activeFilename={activeFilename}
            tableData={tableData}
            onEdit={handleEdit}
            onAddRow={handleAddRow}
            onDeleteRow={handleDeleteRow}
            onSave={handleSave}
          />
        )}

        {currentView === 'project' && currentProject && !activeFilename && !activeQuotationFilename && (
          <ProjectConfigWorkspace
            currentProject={currentProject}
            projectRateCard={projectRateCard}
            rateCards={rateCards}
            projectTemplate={projectTemplate}
            templates={templates}
            availableColumns={availableColumns}
            selectedColumns={selectedColumns}
            presetColumns={presetColumns}
            columnMappings={columnMappings}
            onUpdateRateCard={async (name) => {
              await api.updateProjectRateCard(currentProject!, name);
              setProjectRateCard(name);
            }}
            onUpdateTemplate={handleUpdateProjectTemplate}
            onUpdateColumns={handleUpdateSelectedColumns}
            onUpdateColumnMapping={handleUpdateColumnMapping}
            onUploadTemplate={handleUploadTemplate}
            onDeleteTemplate={handleDeleteTemplate}
          />
        )}

        {currentView === 'ratecard' && currentRateCard && (
          <RateCardWorkspace
            currentRateCard={currentRateCard}
            ratecardTableData={ratecardTableData}
            importing={ratecardImporting}
            onImport={handleRateCardImport}
          />
        )}

        {!currentView && (
          <h1 className="text-3xl font-light mb-8 text-gray-700">请从左侧选择一个项目或协议定价表</h1>
        )}
      </main>

      <TaskNotification
        status={ocrTaskStatus}
        labels={{ processing: 'AI 识别中...', done: '识别完成', error: '识别失败' }}
        onDismiss={() => setOcrTaskStatus(null)}
      />
    </div>
  );
}

export default App;
