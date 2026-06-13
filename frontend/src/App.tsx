import { useState, useEffect, useRef } from 'react';
import * as api from './api';
import { TableData, TableItem } from './types';
import { DataTable } from './components/DataTable';

function App() {
  const [projects, setProjects] = useState<string[]>([]);
  const [currentProject, setCurrentProject] = useState<string | null>(null);
  const [projectRateCard, setProjectRateCard] = useState<string | null>(null);
  const [rateCards, setRateCards] = useState<string[]>([]);
  const [currentRateCard, setCurrentRateCard] = useState<string | null>(null);
  const [ocrFiles, setOcrFiles] = useState<string[]>([]);
  const [activeFilename, setActiveFilename] = useState<string | null>(null);
  const [tableData, setTableData] = useState<TableData>({ items: [], remarks: '' });
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<'data' | 'upload'>('data');
  const [projectName, setProjectName] = useState('');

  // 定价表视图状态（简洁版：一个定价表对应一个 data.json）
  const [currentView, setCurrentView] = useState<'project' | 'ratecard' | null>(null);
  const [ratecardTableData, setRatecardTableData] = useState<TableData>({ items: [], remarks: '' });
  const [ratecardImporting, setRatecardImporting] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pollingRef = useRef<ReturnType<typeof setInterval> | null>(null);

  const fetchProjects = async () => { try { const res = await api.getProjects(); setProjects(res.data.projects); } catch (err) { console.error(err); } };
  const fetchRateCards = async () => { try { const res = await api.getRateCards(); setRateCards(res.data.ratecards); } catch (err) { console.error(err); } };
  const fetchOcrFiles = async (projectName: string) => { try { const res = await api.getOcrFiles(projectName); setOcrFiles(res.data.files); } catch (err) { console.error(err); } };

  useEffect(() => { fetchProjects(); fetchRateCards(); }, []);

  useEffect(() => {
    if (currentProject) {
      fetchOcrFiles(currentProject);
      api.getProjectInfo(currentProject).then(res => setProjectRateCard(res.data.ratecard_name)).catch((err: any) => { alert("获取项目信息失败: " + (err?.message || String(err))); });
      setTableData({ items: [], remarks: '' });
      setActiveTab('data');
    }
  }, [currentProject]);

  // 定价表：选中后加载 data.json
  useEffect(() => {
    if (currentRateCard) {
      api.getRateCardData(currentRateCard).then(res => {
        const d = res.data.data || res.data;
        setRatecardTableData({
          items: Array.isArray(d.items) ? d.items : [],
          remarks: d.remarks || ''
        });
      }).catch(() => setRatecardTableData({ items: [], remarks: '' }));
    }
  }, [currentRateCard]);

  // 组件卸载时清理 OCR 轮询定时器
  useEffect(() => {
    return () => {
      if (pollingRef.current) {
        clearInterval(pollingRef.current);
        pollingRef.current = null;
      }
    };
  }, []);

  const handleEdit = (index: number, field: keyof TableItem, value: string) => {
    const items = tableData?.items;
    if (!items || index < 0 || index >= items.length) return;
    const newItems = [...items];
    newItems[index][field] = value;
    setTableData({ ...tableData, items: newItems });
  };

  const handleSelectFile = async (filename: string) => {
    try {
      const res = await api.getOcrData(currentProject!, filename);
      // OCR 文件两种可能格式:
      //   { success, data: { items, remarks }, timestamp }  — 成功/已保存
      //   { success: false, error: "..." }                    — 失败
      const fileData = res.data.data || res.data;
      setTableData({
        items: Array.isArray(fileData.items) ? fileData.items : [],
        remarks: fileData.remarks || ''
      });
      setActiveFilename(filename);
      setActiveTab('data');
    } catch (err) { alert("读取数据失败"); }
  };

  const handleSave = async () => {
    if (!activeFilename) return;
    try {
      await api.saveOcrData(currentProject!, activeFilename, tableData);
      alert("保存成功！");
    } catch (err) { alert("保存失败"); }
  };

  // ===== 定价表处理 =====
  const handleRateCardEdit = (index: number, field: keyof TableItem, value: string) => {
    const items = ratecardTableData?.items;
    if (!items || index < 0 || index >= items.length) return;
    const newItems = [...items];
    newItems[index][field] = value;
    setRatecardTableData({ ...ratecardTableData, items: newItems });
  };

  const handleRateCardImport = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = e.target.files;
    if (!files || files.length === 0) return;
    setRatecardImporting(true);
    const formData = new FormData();
    formData.append('file', files[0]);
    try {
      const res = await api.importRateCardFile(currentRateCard!, formData);
      const d = res.data.data || res.data;
      setRatecardTableData({
        items: Array.isArray(d.items) ? d.items : [],
        remarks: d.remarks || ''
      });
    } catch (err: any) {
      alert("导入失败: " + (err.response?.data?.detail || err.message));
    }
    setRatecardImporting(false);
    // 清空 input 以便重复选择同一文件
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  return (
    <div className="flex h-screen font-sans text-gray-800">
      {/* 左侧：项目管理 */}
      <aside className="w-72 p-6 bg-slate-800 text-white flex flex-col overflow-y-auto">
        <h2 className="text-2xl font-light mb-8">CarrotMRO</h2>

        {/* 创建项目 */}
        <div className="mb-8">
          <input
            value={projectName}
            onChange={(e) => setProjectName(e.target.value)}
            placeholder="新项目名称..."
            className="w-full p-2 rounded bg-white text-gray-800 placeholder-gray-400 mb-2"
          />
          <button
            onClick={async () => { await api.createProject(projectName); setProjectName(''); fetchProjects(); }}
            className="w-full p-2 rounded bg-blue-500 hover:bg-blue-600 transition"
          >
            创建项目
          </button>
        </div>

        <h3 className="text-sm font-semibold opacity-70 mb-4">项目列表</h3>
        <ul className="space-y-2">
          {projects.map(p => (
            <li
              key={p}
              onClick={() => { setCurrentProject(p); setCurrentRateCard(null); setCurrentView('project'); }}
              className={`p-3 cursor-pointer rounded transition ${currentProject === p ? 'bg-slate-700' : 'hover:bg-slate-700/50'}`}
            >
              {p}
            </li>
          ))}
        </ul>

        {/* 定价表管理 */}
        <div className="mt-8">
          <h3 className="text-sm font-semibold opacity-70 mb-4">协议定价表</h3>
          <div className="mb-4">
            <input
              id="new-ratecard-input"
              placeholder="新定价表名称..."
              className="w-full p-2 rounded bg-white text-gray-800 placeholder-gray-400 mb-2 text-sm"
              onKeyDown={async (e) => {
                if (e.key === 'Enter') {
                  const val = (e.target as HTMLInputElement).value.trim();
                  if (!val) return;
                  await api.createRateCard(val);
                  (e.target as HTMLInputElement).value = '';
                  fetchRateCards();
                }
              }}
            />
            <button
              onClick={async () => {
                const input = document.getElementById('new-ratecard-input') as HTMLInputElement;
                if (!input.value.trim()) return;
                await api.createRateCard(input.value);
                input.value = '';
                fetchRateCards();
              }}
              className="w-full p-2 rounded bg-blue-500 hover:bg-blue-600 transition text-sm"
            >
              新建定价表
            </button>
          </div>
          <ul className="space-y-1">
            {rateCards.map(rc => (
              <li
                key={rc}
                onClick={() => { setCurrentRateCard(rc); setCurrentProject(null); setCurrentView('ratecard'); }}
                className={`p-2 cursor-pointer rounded transition text-sm ${currentRateCard === rc ? 'bg-slate-700' : 'hover:bg-slate-700/50'}`}
              >
                {rc}
              </li>
            ))}
          </ul>
        </div>

        {currentProject && (
          <div className="mt-8">
            <h4 className="text-sm font-semibold opacity-70 mb-3">识别记录</h4>
            {ocrFiles.map(f => (
              <div
                key={f}
                className="group flex items-center justify-between cursor-pointer text-slate-400 hover:text-white mb-2 text-sm"
              >
                <span className="truncate flex-1" onClick={() => handleSelectFile(f)}>{f}</span>
                <span
                  className="ml-2 text-red-400 opacity-0 group-hover:opacity-100 hover:text-red-300 transition-opacity select-none"
                  title="删除"
                  onClick={async (e) => {
                    e.stopPropagation();
                    if (!confirm(`确定删除 ${f} 吗？`)) return;
                    try {
                      await api.deleteOcrFile(currentProject!, f);
                      fetchOcrFiles(currentProject!);
                      if (activeFilename === f) {
                        setActiveFilename(null);
                        setTableData({ items: [], remarks: '' });
                      }
                    } catch (err) {
                      alert('删除失败');
                    }
                  }}
                >
                  ×
                </span>
              </div>
            ))}
          </div>
        )}
      </aside>

      {/* 右侧：工作台 */}
      <main className="flex-1 p-10 bg-gray-100 overflow-y-auto">
        {/* ===== 项目视图 ===== */}
        {currentView === 'project' && currentProject && (
          <>
            <h1 className="text-3xl font-light mb-8 text-gray-700">项目: {currentProject}</h1>

            <div className="bg-white p-8 rounded-2xl shadow-sm mb-6">
              <h2 className="text-xl font-semibold mb-4">项目配置</h2>
              <label className="block text-sm text-gray-600 mb-2">关联协议定价表:</label>
              <select
                value={projectRateCard || ''}
                onChange={async (e) => {
                  const val = e.target.value;
                  await api.updateProjectRateCard(currentProject!, val);
                  setProjectRateCard(val);
                }}
                className="w-full p-2 border rounded-lg"
              >
                <option value="">未关联</option>
                {rateCards.map(rc => <option key={rc} value={rc}>{rc}</option>)}
              </select>
            </div>

            <div className="bg-white p-8 rounded-2xl shadow-sm mb-6">
              <div className="mb-6 flex gap-4">
                <button onClick={() => setActiveTab('data')} className={`px-5 py-2 rounded-lg transition ${activeTab === 'data' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>数据工作台</button>
                <button onClick={() => setActiveTab('upload')} className={`px-5 py-2 rounded-lg transition ${activeTab === 'upload' ? 'bg-blue-600 text-white' : 'bg-gray-200 text-gray-700 hover:bg-gray-300'}`}>上传新图片</button>
              </div>

              {activeTab === 'upload' ? (
                <div className="p-12 text-center border-2 border-dashed border-gray-300 rounded-xl bg-gray-50">
                  <label className="cursor-pointer bg-blue-500 hover:bg-blue-600 text-white px-6 py-3 rounded-lg font-medium transition">
                    选择图片文件
                    <input type="file" multiple className="hidden" onChange={async (e) => {
                      if (!e.target.files) return;
                      setLoading(true);
                      const formData = new FormData();
                      for (let i = 0; i < e.target.files.length; i++) formData.append('files', e.target.files[i]);

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
                              // OCR 返回: { success, data: { items, remarks }, timestamp }
                              const actualData = rawResult.data || rawResult;

                              // OCR 识别失败（如配额超限）
                              if (rawResult.success === false) {
                                setLoading(false);
                                alert("识别失败: " + (rawResult.error || "未知错误"));
                                fetchOcrFiles(currentProject!);
                                return;
                              }

                              setTableData({
                                items: Array.isArray(actualData.items) ? actualData.items : [],
                                remarks: actualData.remarks || ''
                              });

                              fetchOcrFiles(currentProject!);
                              setLoading(false);
                              setActiveTab('data');
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
                    }} />
                  </label>
                  {loading && <p className="mt-4 text-blue-600 font-medium">AI 识别中...</p>}
                </div>
              ) : (
                <div>
                  {activeFilename && <h3 className="mb-4 text-lg font-semibold text-gray-800">当前文件: {activeFilename}</h3>}
                  <DataTable key={activeFilename} items={tableData?.items ?? []} onEdit={handleEdit} />
                  <button onClick={handleSave} className="mt-6 px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition">保存修改</button>
                </div>
              )}
            </div>
          </>
        )}

        {/* ===== 定价表视图 ===== */}
        {currentView === 'ratecard' && currentRateCard && (
          <>
            <h1 className="text-3xl font-light mb-8 text-gray-700">协议定价表: {currentRateCard}</h1>

            <div className="bg-white p-8 rounded-2xl shadow-sm mb-6">
              <div className="flex items-center justify-between mb-6">
                <div></div>
                <div className="flex gap-3">
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx,.xls,.csv"
                    className="hidden"
                    onChange={handleRateCardImport}
                  />
                  <button
                    onClick={() => fileInputRef.current?.click()}
                    disabled={ratecardImporting}
                    className="px-4 py-2 bg-amber-500 text-white rounded-lg hover:bg-amber-600 transition text-sm disabled:opacity-50"
                  >
                    {ratecardImporting ? '导入中...' : '导入 Excel / CSV'}
                  </button>

                </div>
              </div>

              <DataTable items={ratecardTableData?.items ?? []} onEdit={handleRateCardEdit} />

              {ratecardTableData.remarks && (
                <div className="mt-4 p-3 bg-gray-50 rounded-lg text-sm text-gray-600">
                  {ratecardTableData.remarks}
                </div>
              )}
            </div>
          </>
        )}

        {!currentView && (
          <h1 className="text-3xl font-light mb-8 text-gray-700">请从左侧选择一个项目或协议定价表</h1>
        )}
      </main>
    </div>
  );
}

export default App;
