import axios from 'axios';
import { ProjectSettings } from './types';

// ===== 预制列相关 =====
export const getPresetColumns = () => axios.get<string[]>('/api/preset-columns');

// ===== 项目与设置相关 =====
export const getProjects = () => axios.get<string[]>('/api/projects');
export const createProject = (name: string) => axios.post<{ success: boolean; name: string }>(`/api/projects/${name}`);
export const getProjectInfo = (projectName: string) => axios.get<ProjectSettings>(`/api/projects/${projectName}`);
export const updateProjectSettings = (projectName: string, settings: Partial<ProjectSettings>) =>
  axios.patch<{ success: boolean; settings: ProjectSettings }>(`/api/projects/${projectName}/settings`, settings);

// 保留辅助匹配导出映射兼容接口
export const updateProjectRateCard = (projectName: string, ratecardName: string | null) =>
  updateProjectSettings(projectName, { ratecard_name: ratecardName });
export const updateProjectTemplate = (projectName: string, templateName: string | null) =>
  updateProjectSettings(projectName, { template_name: templateName });

// ===== 报价单相关 =====
export const getQuotations = (projectName: string) => axios.get<{ files: string[] }>(`/api/projects/${projectName}/quotations`);
export const createQuotation = (projectName: string) => axios.post<{ success: boolean; file: string }>(`/api/projects/${projectName}/quotations`);
export const getQuotationData = (projectName: string, filename: string) =>
  axios.get<{ items: any[]; created_at?: string; last_edit_time?: string }>(`/api/projects/${projectName}/quotations/${filename}`);
export const saveQuotationData = (projectName: string, filename: string, data: any) =>
  axios.put<{ success: boolean }>(`/api/projects/${projectName}/quotations/${filename}`, data);
export const deleteQuotation = (projectName: string, filename: string) =>
  axios.delete<{ success: boolean }>(`/api/projects/${projectName}/quotations/${filename}`);

// ===== 图片 OCR 识别导入（沉淀为报价单内部的导入功能） =====
export const uploadOcrFiles = (projectName: string, formData: FormData) =>
  axios.post<{ task_id: string }>(`/api/projects/${projectName}/ocr`, formData);
export const checkTaskStatus = (taskId: string) => axios.get(`/api/tasks/${taskId}`);
export const cancelTask = (taskId: string) => axios.delete(`/api/tasks/${taskId}`);

// ===== 协议定价表相关 =====
export const getRateCards = () => axios.get<string[]>('/api/ratecards');
export const createRateCard = (name: string) => axios.post<{ success: boolean }>(`/api/ratecards/${name}`);
export const getRateCardData = (ratecardName: string) =>
  axios.get<{ columns: string[]; items: Record<string, string>[] }>(`/api/ratecards/${ratecardName}`);

/** 定价表导入预览（解析原始表头与前 5 行样例） */
export const previewRateCardImport = (ratecardName: string, formData: FormData) =>
  axios.post<{ headers: string[]; sampleRows: Record<string, string>[] }>(`/api/ratecards/${ratecardName}/import-preview`, formData);

/** 确认表头映射后，最终导入保存定价表 */
export const importRateCardFile = (ratecardName: string, payload: { headers: string[]; items: Record<string, string>[]; mapping: Record<string, string> }) =>
  axios.post<{ success: boolean; columns: string[]; count: number }>(`/api/ratecards/${ratecardName}/import`, payload);

// ===== 模板相关 =====
export const getTemplates = () => axios.get<string[]>('/api/templates');
export const uploadTemplate = (formData: FormData) => axios.post<{ success: boolean; filename: string }>('/api/templates', filename => formData);
export const deleteTemplate = (filename: string) => axios.delete<{ success: boolean }>(`/api/templates/${filename}`);

// ===== 报价单导入导出 =====
export const exportQuotation = (projectName: string, filename: string) =>
  axios.post(`/api/projects/${projectName}/quotations/${filename}/export`, {}, { responseType: 'blob' });
export const importQuotation = (projectName: string, formData: FormData) =>
  axios.post(`/api/projects/${projectName}/quotations/import`, formData);

// ===== 定价表匹配 =====
export const matchRateCard = (ratecardName: string, queries: string[], limit: number = 5) =>
  axios.post<Record<string, [string, number, Record<string, string>][]>>('/api/match', {
    ratecard_name: ratecardName,
    queries,
    limit,
  });

// ===== 系统设置相关 =====
export const getSettings = () => axios.get('/api/settings');
export const saveSettings = (data: any) => axios.put('/api/settings', data);
export const updateSettings = saveSettings;
export const testLlm = (data: any) => axios.post('/api/settings/test-llm', data);
export const testLlmConfig = testLlm;
