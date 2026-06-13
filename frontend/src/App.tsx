import { useState, useEffect } from 'react';
import * as api from './api';
import { TableData, TableItem } from './types';
import { DataTable } from './components/DataTable';

function App() {
  const [projects, setProjects] = useState<string[]>([]);
  const [currentProject, setCurrentProject] = useState<string | null>(null);
  const [ocrFiles, setOcrFiles] = useState<string[]>([]);
  const [activeFilename, setActiveFilename] = useState<string | null>(null);
  const [tableData, setTableData] = useState<TableData>({ items: [], remarks: '' });
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState<'data' | 'upload'>('data');
  const [projectName, setProjectName] = useState('');

  const fetchProjects = async () => { try { const res = await api.getProjects(); setProjects(res.data.projects); } catch (err) { console.error(err); } };
  const fetchOcrFiles = async (projectName: string) => { try { const res = await api.getOcrFiles(projectName); setOcrFiles(res.data.files); } catch (err) { console.error(err); } };

  useEffect(() => { fetchProjects(); }, []);

  useEffect(() => {
    if (currentProject) {
      fetchOcrFiles(currentProject);
      setTableData({ items: [], remarks: '' });
      setActiveTab('data');
    }
  }, [currentProject]);

  const handleEdit = (index: number, field: keyof TableItem, value: string) => {
    const newItems = [...tableData.items];
    newItems[index][field] = value;
    setTableData({ ...tableData, items: newItems });
  };

  const handleSelectFile = async (filename: string) => {
    try {
      const res = await api.getOcrData(currentProject!, filename);
      setTableData(res.data.data);
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

  return (
    <div className="flex h-screen font-sans text-gray-800">
      {/* 左侧：项目管理 */}
      <aside className="w-72 p-6 bg-slate-800 text-white flex flex-col overflow-y-auto">
        <h2 className="text-2xl font-light mb-8">CarrotMRO</h2>
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
              onClick={() => setCurrentProject(p)} 
              className={`p-3 cursor-pointer rounded transition ${currentProject === p ? 'bg-slate-700' : 'hover:bg-slate-700/50'}`}
            >
              {p}
            </li>
          ))}
        </ul>
        
        {currentProject && (
          <div className="mt-8">
            <h4 className="text-sm font-semibold opacity-70 mb-3">识别记录</h4>
            {ocrFiles.map(f => (
              <div 
                key={f} 
                onClick={() => handleSelectFile(f)} 
                className="cursor-pointer text-slate-400 hover:text-white mb-2 text-sm truncate"
              >
                {f}
              </div>
            ))}
          </div>
        )}
      </aside>

      {/* 右侧：工作台 */}
      <main className="flex-1 p-10 bg-gray-100 overflow-y-auto">
        <h1 className="text-3xl font-light mb-8 text-gray-700">{currentProject ? `项目: ${currentProject}` : '请选择一个项目'}</h1>
        
        {currentProject && (
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
                    
                    // 1. 提交任务，获取 task_id
                    const res = await api.uploadOcrFiles(currentProject!, formData);
                    const { task_id } = res.data;

                    // 2. 开启轮询
                    const interval = setInterval(async () => {
                      const statusRes = await api.checkTaskStatus(task_id);
                      
                      if (statusRes.data.status === 'done') {
                        clearInterval(interval);
                        
                        // 1. 先更新文件状态，确保上下文同步
                        setActiveFilename(statusRes.data.file);
                        
                        // 2. 更新表格数据
                        // 根据日志，数据结构为 { success: true, data: { items: [], ... }, ... }
                        const rawResult = statusRes.data.result || {};
                        const actualData = rawResult.data || rawResult;
                        
                        setTableData({
                          items: Array.isArray(actualData.items) ? actualData.items : [],
                          remarks: actualData.remarks || ''
                        });
                        
                        // 3. 刷新文件列表
                        fetchOcrFiles(currentProject!);
                        
                        // 4. 最后切换视图，强制触发渲染
                        setLoading(false);
                        setActiveTab('data');
                      } else if (statusRes.data.status === 'error') {
                        clearInterval(interval);
                        setLoading(false);
                        alert("识别失败: " + statusRes.data.message);
                      }
                    }, 2000);
                  }} />
                </label>
                {loading && <p className="mt-4 text-blue-600 font-medium">AI 识别中...</p>}
              </div>
            ) : (
              <div>
                {activeFilename && <h3 className="mb-4 text-lg font-semibold text-gray-800">当前文件: {activeFilename}</h3>}
                <DataTable key={activeFilename} items={tableData.items} onEdit={handleEdit} />
                <button onClick={handleSave} className="mt-6 px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition">保存修改</button>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
export default App;
