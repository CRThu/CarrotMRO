import axios from 'axios';

// ===== 项目相关 =====
export const getProjects = () => axios.get('/api/projects');
export const createProject = (name: string) => axios.put(`/api/create-project/${name}`);
export const getOcrFiles = (projectName: string) => axios.get(`/api/ocr-files/${projectName}`);
export const getOcrData = (projectName: string, filename: string) => axios.get(`/api/ocr-data/${projectName}/${filename}`);
export const saveOcrData = (projectName: string, filename: string, data: any) =>
  axios.post(`/api/save-ocr/${projectName}/${filename}`, data);
export const uploadOcrFiles = (projectName: string, formData: FormData) =>
  axios.post(`/api/ocr/${projectName}`, formData);
export const checkTaskStatus = (taskId: string) => axios.get(`/api/task-status/${taskId}`);
export const getProjectInfo = (projectName: string) => axios.get(`/api/ocr-data/${projectName}/project.json`);
export const updateProjectRateCard = (projectName: string, ratecardName: string) =>
  axios.post(`/api/update-project-ratecard/${projectName}`, { ratecard_name: ratecardName });

// ===== 协议定价表相关 =====
export const getRateCards = () => axios.get('/api/ratecards');
export const createRateCard = (name: string) => axios.put(`/api/create-ratecard/${name}`);
/** 获取定价表的 data.json（单文件） */
export const getRateCardData = (ratecardName: string) => axios.get(`/api/ratecard-data/${ratecardName}`);
/** 保存定价表编辑数据 */
export const saveRateCardData = (ratecardName: string, data: any) =>
  axios.post(`/api/save-ratecard-data/${ratecardName}`, data);
/** 导入 Excel / CSV 文件到定价表 */
export const importRateCardFile = (ratecardName: string, formData: FormData) =>
  axios.post(`/api/import-ratecard/${ratecardName}`, formData);
