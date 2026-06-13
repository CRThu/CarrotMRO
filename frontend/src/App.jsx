import { useState, useEffect } from 'react';
import axios from 'axios';

// 样式定义
const styles = {
  container: { display: 'flex', height: '100vh', fontFamily: "'Segoe UI', Tahoma, sans-serif", color: '#333' },
  sidebar: { width: '280px', padding: '20px', background: '#34495e', color: '#fff', display: 'flex', flexDirection: 'column' },
  workspace: { flex: 1, padding: '40px', background: '#ecf0f1', overflowY: 'auto' },
  card: { background: '#fff', padding: '25px', borderRadius: '12px', boxShadow: '0 4px 6px rgba(0,0,0,0.05)', marginBottom: '20px' },
  button: { padding: '10px 16px', borderRadius: '6px', border: 'none', cursor: 'pointer', transition: '0.3s' },
  table: { width: '100%', borderCollapse: 'collapse', marginTop: '20px' },
  th: { padding: '12px', background: '#f4f6f8', textAlign: 'left', borderBottom: '2px solid #e1e4e8' },
  td: { padding: '12px', borderBottom: '1px solid #e1e4e8' },
  input: { padding: '8px', width: '90%', border: '1px solid #ddd', borderRadius: '4px' }
};

function App() {
  const [projects, setProjects] = useState([]);
  const [currentProject, setCurrentProject] = useState(null);
  const [ocrFiles, setOcrFiles] = useState([]);
  const [activeFilename, setActiveFilename] = useState(null);
  const [tableData, setTableData] = useState({ items: [], remarks: '' });
  const [loading, setLoading] = useState(false);
  const [activeTab, setActiveTab] = useState('data');
  const [projectName, setProjectName] = useState('');

  const fetchProjects = async () => { try { const res = await axios.get('/api/projects'); setProjects(res.data.projects); } catch (err) { console.error(err); } };
  const fetchOcrFiles = async (projectName) => { try { const res = await axios.get(`/api/ocr-files/${projectName}`); setOcrFiles(res.data.files); } catch (err) { console.error(err); } };

  useEffect(() => { fetchProjects(); }, []);

  useEffect(() => {
    if (currentProject) {
      fetchOcrFiles(currentProject);
      setTableData({ items: [], remarks: '' });
      setActiveTab('data');
    }
  }, [currentProject]);

  const handleEdit = (index, field, value) => {
    const newItems = [...tableData.items];
    newItems[index][field] = value;
    setTableData({ ...tableData, items: newItems });
  };

  const handleSelectFile = async (filename) => {
    try {
      const res = await axios.get(`/api/ocr-data/${currentProject}/${filename}`);
      setTableData(res.data.data);
      setActiveFilename(filename);
      setActiveTab('data');
    } catch (err) { alert("读取数据失败"); }
  };

  const handleSave = async () => {
    if (!activeFilename) return;
    try {
      await axios.post(`/api/save-ocr/${currentProject}/${activeFilename}`, tableData);
      alert("保存成功！");
    } catch (err) { alert("保存失败"); }
  };

  return (
    <div style={styles.container}>
      {/* 左侧：项目管理 */}
      <aside style={styles.sidebar}>
        <h2 style={{ marginBottom: '30px', fontWeight: '300' }}>CarrotMRO</h2>
        <div style={{ marginBottom: '30px' }}>
          <input value={projectName} onChange={(e) => setProjectName(e.target.value)} placeholder="新项目名称..." style={{ width: '100%', padding: '10px', borderRadius: '4px', border: 'none' }} />
          <button onClick={async () => { await axios.put(`/api/create-project/${projectName}`); setProjectName(''); fetchProjects(); }} style={{ ...styles.button, width: '100%', marginTop: '10px', background: '#3498db', color: '#fff' }}>创建项目</button>
        </div>
        
        <h3 style={{fontSize: '16px', opacity: 0.8}}>项目列表</h3>
        <ul style={{ listStyle: 'none', padding: 0 }}>
          {projects.map(p => (
            <li key={p} onClick={() => setCurrentProject(p)} 
                style={{ padding: '12px', cursor: 'pointer', background: currentProject === p ? '#2c3e50' : 'transparent', borderRadius: '6px' }}>
              {p}
            </li>
          ))}
        </ul>
        
        {currentProject && (
          <div style={{ marginTop: '20px' }}>
            <h4 style={{fontSize: '14px', opacity: 0.8}}>识别记录</h4>
            {ocrFiles.map(f => (
              <div key={f} onClick={() => handleSelectFile(f)} style={{ cursor: 'pointer', color: '#bdc3c7', marginBottom: '8px', fontSize: '13px' }}>{f}</div>
            ))}
          </div>
        )}
      </aside>

      {/* 右侧：工作台 */}
      <main style={styles.workspace}>
        <h1 style={{ fontWeight: '400', marginBottom: '30px' }}>{currentProject ? `项目: ${currentProject}` : '请选择一个项目'}</h1>
        
        {currentProject && (
          <div style={styles.card}>
            <div style={{ marginBottom: '20px' }}>
              <button onClick={() => setActiveTab('data')} style={{ ...styles.button, background: activeTab === 'data' ? '#3498db' : '#eee', color: activeTab === 'data' ? '#fff' : '#333' }}>数据工作台</button>
              <button onClick={() => setActiveTab('upload')} style={{ ...styles.button, background: activeTab === 'upload' ? '#3498db' : '#eee', color: activeTab === 'upload' ? '#fff' : '#333', marginLeft: '10px' }}>上传新图片</button>
            </div>

            {activeTab === 'upload' ? (
              <div style={{ padding: '40px', textAlign: 'center', border: '2px dashed #ccc', borderRadius: '8px' }}>
                <input type="file" multiple onChange={async (e) => {
                  setLoading(true); const formData = new FormData();
                  for (let i = 0; i < e.target.files.length; i++) formData.append('files', e.target.files[i]);
                  const res = await axios.post(`/api/ocr/${currentProject}`, formData);
                  setTableData(res.data.data.data);
                  setActiveFilename(res.data.file.split('\\').pop().split('/').pop());
                  fetchOcrFiles(currentProject); setLoading(false); setActiveTab('data');
                }} />
                {loading && <p>AI 识别中...</p>}
              </div>
            ) : (
              <div>
                {activeFilename && <h3 style={{marginBottom: '15px'}}>当前文件: {activeFilename}</h3>}
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={styles.th}>项目</th><th style={styles.th}>数量</th><th style={styles.th}>单位</th><th style={styles.th}>单价</th>
                    </tr>
                  </thead>
                  <tbody>
                    {tableData.items.map((item, i) => (
                      <tr key={i}>
                        <td style={styles.td}><input value={item.name} onChange={(e) => handleEdit(i, 'name', e.target.value)} style={styles.input} /></td>
                        <td style={styles.td}><input value={item.quantity} onChange={(e) => handleEdit(i, 'quantity', e.target.value)} style={styles.input} /></td>
                        <td style={styles.td}><input value={item.unit} onChange={(e) => handleEdit(i, 'unit', e.target.value)} style={styles.input} /></td>
                        <td style={styles.td}><input value={item.unit_price || ''} onChange={(e) => handleEdit(i, 'unit_price', e.target.value)} style={styles.input} /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <button onClick={handleSave} style={{ ...styles.button, marginTop: '20px', background: '#27ae60', color: '#fff' }}>保存修改</button>
              </div>
            )}
          </div>
        )}
      </main>
    </div>
  );
}
export default App;
