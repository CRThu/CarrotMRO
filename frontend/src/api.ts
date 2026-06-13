import axios from 'axios';

// 定义项目相关的接口
export const getProjects = () => axios.get('/api/projects');
export const createProject = (name: string) => axios.put(`/api/create-project/${name}`);

// 定义OCR文件相关的接口
export const getOcrFiles = (projectName: string) => axios.get(`/api/ocr-files/${projectName}`);
export const getOcrData = (projectName: string, filename: string) => axios.get(`/api/ocr-data/${projectName}/${filename}`);
export const saveOcrData = (projectName: string, filename: string, data: any) => 
  axios.post(`/api/save-ocr/${projectName}/${filename}`, data);
export const uploadOcrFiles = (projectName: string, formData: FormData) => 
  axios.post(`/api/ocr/${projectName}`, formData);
export const checkTaskStatus = (taskId: string) => axios.get(`/api/task-status/${taskId}`);
