/**
 * LLM / OCR 多模型统一接入服务
 * 本模块基于标准的 OpenAI Chat Completions 兼容 API 协议规范搭建。
 * 支持：Google Gemini (OpenAI 端点), DeepSeek, Xiaomi MiMo, 及任意 OpenAI 兼容中转站。
 * 参数支持：apiKey, model, baseUrl, proxy (HTTP/HTTPS 代理)
 */

import axios, { AxiosRequestConfig } from 'axios';
import { HttpsProxyAgent } from 'https-proxy-agent';
import { getActiveProviderConfig, ProviderConfig } from './settings.js';

export interface ImageInput {
  buffer: Buffer;
  mimeType: string;
}

export interface SimpleLlmConfig {
  apiKey: string;
  model: string;
  baseUrl: string;
  proxy?: string;
  provider?: string;
}

/**
 * 辅助函数：根据 provider 或 baseUrl 解析标准的 Base URL
 */
function resolveBaseUrl(config: SimpleLlmConfig): string {
  if (config.baseUrl && config.baseUrl.trim() !== '') {
    return config.baseUrl.trim().replace(/\/+$/, '');
  }
  const provider = (config.provider || 'google').toLowerCase();
  switch (provider) {
    case 'google':
      return 'https://generativelanguage.googleapis.com/v1beta/openai';
    case 'deepseek':
      return 'https://api.deepseek.com/v1';
    case 'mimo':
      return 'https://api.xiaomimimo.com/v1';
    default:
      return 'https://api.openai.com/v1';
  }
}

/**
 * 构造 Axios 请求配置（包含 HTTP/HTTPS 代理支持）
 */
function createAxiosConfig(config: SimpleLlmConfig): AxiosRequestConfig {
  const axiosConfig: AxiosRequestConfig = {
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${config.apiKey}`,
    },
    timeout: 60000,
  };

  if (config.proxy && config.proxy.trim() !== '') {
    const agent = new HttpsProxyAgent(config.proxy.trim());
    axiosConfig.httpAgent = agent;
    axiosConfig.httpsAgent = agent;
    axiosConfig.proxy = false; // 禁用 axios 自带代理处理机制，由 HttpsProxyAgent 接管
  }

  return axiosConfig;
}

/**
 * 动态提取返回文本中的 JSON 结构
 */
export function extractJsonFromText(text: string): any | null {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch {}

  // 提取 ```json ... ``` 代码块
  const jsonBlockMatch = text.match(/```(?:json)?\s*\n?([\s\S]*?)\n?```/);
  if (jsonBlockMatch) {
    try {
      return JSON.parse(jsonBlockMatch[1]);
    } catch {}
  }

  // 提取第一个 { ... } 闭合 JSON 对象
  const objectMatch = text.match(/\{[\s\S]*\}/);
  if (objectMatch) {
    try {
      return JSON.parse(objectMatch[0]);
    } catch {}
  }

  return null;
}

/**
 * 生成动态 OCR Prompt 引导 LLM 返回标准结构 JSON
 */
export function buildOcrPrompt(columns: string[]): string {
  const colList = columns.join('、');
  return `你是一个专业的工程报价单OCR识别助手。
你的任务是阅读并理解图片中的报价单内容，提取其中的项目信息。

任务要求：
1. 提取以下字段：${colList}。
2. 不要对项目进行序号编号。
3. 不要遗漏项目，也不要虚构项目。
4. 将所有不确定或有疑问的内容汇总在备注字段中。

请严格按照以下 JSON 格式返回结果，不要包含其他内容：
{
  "items": [
    { ${columns.map(col => `"${col}": ""`).join(', ')} }
  ],
  "remarks": ""
}`;
}

/**
 * 多模型 OCR 图片识别方法
 */
export async function runOcrWithLlm(images: ImageInput[], columns: string[], overrideConfig?: SimpleLlmConfig): Promise<{ success: boolean; data?: any; error?: string }> {
  const activeConfig = getActiveProviderConfig();
  const config: SimpleLlmConfig = overrideConfig || {
    apiKey: activeConfig.apiKey,
    model: activeConfig.model,
    baseUrl: activeConfig.baseUrl,
    proxy: activeConfig.proxy,
    provider: activeConfig.providerName,
  };

  if (!config.apiKey) {
    return { success: false, error: '未配置 当前激活模型服务商的 API Key，请先前往“系统设置”填写对应的 API Key。' };
  }
  if (!columns || columns.length === 0) {
    return { success: false, error: '未配置识别列，请先在项目中勾选识别列。' };
  }

  const baseUrl = resolveBaseUrl(config);
  const prompt = buildOcrPrompt(columns);

  const contentItems: any[] = [{ type: 'text', text: prompt }];

  for (const img of images) {
    const b64 = img.buffer.toString('base64');
    contentItems.push({
      type: 'image_url',
      image_url: {
        url: `data:${img.mimeType || 'image/jpeg'};base64,${b64}`,
      },
    });
  }

  const payload = {
    model: config.model || 'gemini-3.6-flash',
    messages: [
      {
        role: 'user',
        content: contentItems,
      },
    ],
    temperature: 0.1,
  };

  const url = `${baseUrl}/chat/completions`;
  const axiosConfig = createAxiosConfig(config);

  try {
    const response = await axios.post(url, payload, axiosConfig);
    const content = response.data?.choices?.[0]?.message?.content || '';
    const parsedData = extractJsonFromText(typeof content === 'string' ? content : JSON.stringify(content));

    if (!parsedData) {
      return { success: false, error: `无法解析模型返回的 JSON 内容: ${String(content).substring(0, 200)}` };
    }

    let finalData = parsedData;
    if (!finalData.items) {
      finalData = {
        items: Array.isArray(parsedData) ? parsedData : [parsedData],
        remarks: '',
      };
    }
    if (!finalData.remarks) {
      finalData.remarks = '';
    }

    return { success: true, data: finalData };
  } catch (err: any) {
    const errMsg = err.response?.data?.error?.message || err.message || String(err);
    console.error('LLM OCR 识别失败:', errMsg);
    return { success: false, error: `模型调用失败: ${errMsg}` };
  }
}

/**
 * 测试 LLM API 连通性
 */
export async function testLlmConnection(config: SimpleLlmConfig): Promise<{ success: boolean; message: string }> {
  if (!config.apiKey) {
    return { success: false, message: 'API Key 不能为空' };
  }
  const baseUrl = resolveBaseUrl(config);
  const url = `${baseUrl}/chat/completions`;
  const axiosConfig = createAxiosConfig(config);

  const payload = {
    model: config.model || 'gemini-3.6-flash',
    messages: [
      {
        role: 'user',
        content: 'Hello, respond with OK.',
      },
    ],
    max_tokens: 10,
  };

  try {
    const response = await axios.post(url, payload, axiosConfig);
    if (response.status === 200 && response.data?.choices) {
      return { success: true, message: '连接成功！API 接入测试正常。' };
    }
    return { success: false, message: `连接响应异常 (Status: ${response.status})` };
  } catch (err: any) {
    const errMsg = err.response?.data?.error?.message || err.message || String(err);
    return { success: false, message: `连接失败: ${errMsg}` };
  }
}
