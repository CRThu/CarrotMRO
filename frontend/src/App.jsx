import { useState, useEffect } from 'react';
import axios from 'axios';

function App() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [projectName, setProjectName] = useState('');
  const [creating, setCreating] = useState(false);
  const [projects, setProjects] = useState([]);
  const [currentProject, setCurrentProject] = useState(null);

  const fetchProjects = async () => {
    try {
      const res = await axios.get('/api/projects');
      setProjects(res.data.projects);
    } catch (err) {
      console.error("获取项目列表失败", err);
    }
  };

  useEffect(() => {
    fetchProjects();
  }, []);

  const handleCreateProject = async () => {
    if (!projectName) {
      alert("请输入项目名称");
      return;
    }
    setCreating(true);
    try {
      await axios.put(`/api/create-project/${projectName}`);
      alert(`项目 "${projectName}" 创建成功`);
      setProjectName('');
      fetchProjects();
    } catch (err) {
      alert(err.response?.data?.detail || "创建项目失败");
    } finally {
      setCreating(false);
    }
  };

  const handleUpload = async (e) => {
    if (!currentProject) {
        alert("请先选择一个项目");
        return;
    }
    const file = e.target.files[0];
    if (!file) return;
    
    setLoading(true);
    const formData = new FormData();
    formData.append('file', file);
    
    try {
      const res = await axios.post('/api/ocr', formData);
      setData(res.data.data);
    } catch (err) {
      alert("上传失败");
    } finally {
      setLoading(false);
    }
  };

  const sidebarStyle = {
    width: '300px',
    padding: '20px',
    background: '#ffffff',
    borderRight: '1px solid #e1e4e8',
    height: '100vh',
    display: 'flex',
    flexDirection: 'column',
    boxShadow: '2px 0 5px rgba(0,0,0,0.05)'
  };

  const contentStyle = {
    flex: 1,
    padding: '40px',
    background: '#f9f9fb'
  };

  const buttonStyle = {
    padding: '10px 20px',
    background: '#007bff',
    color: 'white',
    border: 'none',
    borderRadius: '6px',
    cursor: 'pointer',
    marginTop: '10px'
  };

  return (
    <div style={{ display: 'flex', minHeight: '100vh' }}>
      <div style={sidebarStyle}>
        <h2>CarrotMRO</h2>
        <div style={{ marginBottom: '30px' }}>
          <h4>创建项目</h4>
          <input 
            type="text" 
            value={projectName} 
            onChange={(e) => setProjectName(e.target.value)} 
            placeholder="项目名称"
            style={{ width: '100%', padding: '8px', marginBottom: '10px', boxSizing: 'border-box' }}
            disabled={creating}
          />
          <button onClick={handleCreateProject} style={buttonStyle} disabled={creating}>
            {creating ? '创建中...' : '创建项目'}
          </button>
        </div>

        <div>
          <h4>项目列表</h4>
          <ul style={{ listStyle: 'none', padding: 0 }}>
            {projects.map(p => (
              <li 
                key={p} 
                onClick={() => setCurrentProject(p)}
                style={{
                  padding: '10px',
                  background: currentProject === p ? '#e1e4e8' : 'transparent',
                  cursor: 'pointer',
                  borderRadius: '4px'
                }}
              >
                {p}
              </li>
            ))}
          </ul>
        </div>
      </div>
      
      <div style={contentStyle}>
        <h1>{currentProject ? `当前项目: ${currentProject}` : '请选择或创建项目'}</h1>
        
        {currentProject && (
          <div style={{ background: 'white', padding: '20px', borderRadius: '8px', boxShadow: '0 2px 4px rgba(0,0,0,0.05)' }}>
            <h3>上传报价单图片</h3>
            <input type="file" onChange={handleUpload} disabled={loading} />
            
            {loading && <p>识别中...</p>}
            {data && (
              <div style={{marginTop: '20px'}}>
                <h3>识别结果:</h3>
                <pre style={{background: '#f0f0f0', padding: '10px', borderRadius: '4px'}}>{JSON.stringify(data, null, 2)}</pre>
                <button style={buttonStyle} onClick={() => alert("即将对接生成 XLSX 功能")}>
                  生成 XLSX 下载
                </button>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

export default App;
