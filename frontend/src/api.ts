import axios from 'axios';

// ===== 项目相关 =====
export const getProjects = () => axios.get('/api/projects');
export const createProject = (name: string) => axios.post(`/api/projects/${name}`);
export const getOcrFiles = (projectName: string) => axios.get(`/api/projects/${projectName}/ocr-files`);
export const getOcrData = (projectName: string, filename: string) => axios.get(`/api/projects/${projectName}/ocr-files/${filename}`);
export const saveOcrData = (projectName: string, filename: string, data: any) =>
  axios.put(`/api/projects/${projectName}/ocr-files/${filename}`, data);
export const deleteOcrFile = (projectName: string, filename: string) =>
  axios.delete(`/api/projects/${projectName}/ocr-files/${filename}`);
export const uploadOcrFiles = (projectName: string, formData: FormData) =>
  axios.post(`/api/projects/${projectName}/ocr`, formData);
export const checkTaskStatus = (taskId: string) => axios.get(`/api/tasks/${taskId}`);
export const getProjectInfo = (projectName: string) => axios.get(`/api/projects/${projectName}`);
export const updateProjectRateCard = (projectName: string, ratecardName: string) =>
  axios.patch(`/api/projects/${projectName}/ratecard`, { ratecard_name: ratecardName });

// ===== 协议定价表相关 =====
export const getRateCards = () => axios.get('/api/ratecards');
export const createRateCard = (name: string) => axios.post(`/api/ratecards/${name}`);
/** 获取定价表的 data.json（单文件） */
export const getRateCardData = (ratecardName: string) => axios.get(`/api/ratecards/${ratecardName}`);
/** 导入 Excel / CSV 文件到定价表 */
export const importRateCardFile = (ratecardName: string, formData: FormData) =>
  axios.post(`/api/ratecards/${ratecardName}/import`, formData);

// ===== 报价单相关 =====
export const getQuotations = (projectName: string) => axios.get(`/api/projects/${projectName}/quotations`);
export const createQuotation = (projectName: string) => axios.post(`/api/projects/${projectName}/quotations`);
export const getQuotationData = (projectName: string, filename: string) =>
  axios.get(`/api/projects/${projectName}/quotations/${filename}`);
export const saveQuotationData = (projectName: string, filename: string, data: any) =>
  axios.put(`/api/projects/${projectName}/quotations/${filename}`, data);
export const deleteQuotation = (projectName: string, filename: string) =>
  axios.delete(`/api/projects/${projectName}/quotations/${filename}`);

// ===== 定价表匹配 =====
export const matchRateCard = (ratecardName: string, queries: string[], limit?: number) =>
  axios.post('/api/match', { ratecard_name: ratecardName, queries, limit: limit ?? 1 });
