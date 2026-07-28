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
    timeout: 45000, // 45 秒超时断线，防止错误端点无限卡死
  };

  if (config.proxy && config.proxy.trim() !== '') {
    const agent = new HttpsProxyAgent(config.proxy.trim());
    axiosConfig.httpAgent = agent;
    axiosConfig.httpsAgent = agent;
    axiosConfig.proxy = false;
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
 * 生成动态 OCR Prompt 引导 LLM 返回标准结构 JSON，支持多页/多图无缝拼接提取
 */
export function buildOcrPrompt(columns: string[], imageCount: number = 1): string {
  const colList = columns.join('、');
  const multiImageNote = imageCount > 1
    ? `特别强调：用户本次上传了 ${imageCount} 张多页/多图报价单图片。请按照图片排列的先后顺序（第 1 张到第 ${imageCount} 张），顺次提取每一页表格中的数据项，并将所有页面的表格行合并为单一完整的 "items" 数组输出，切勿漏掉后几页的数据！\n`
    : '';

  return `你是一个专业的工程与采购报价单 OCR 识别助手。
你的任务是阅读并理解用户上传的报价单图片内容，精确提取其中的表格项目数据。

${multiImageNote}任务要求：
1. 提取每个表格项的以下字段：${colList}。
2. 请按原始表格顺序依次提取，不要遗漏任何项目，也不要虚构项目。
3. 如果某些字段在图中未明示，保留为空字符串 ""。
4. 将所有不确定或有疑问的内容汇总在 "remarks" 备注字段中。

请严格按照以下 JSON 格式返回结果，不要包含任何 markdown 说明之外的代码块：
{
  "items": [
    { ${columns.map(col => `"${col}": ""`).join(', ')} }
  ],
  "remarks": ""
}`;
}

/**
 * 多模型 OCR 图片识别方法（单请求多图 + JSON 输出 + 自动容错降级 + 进度通知）
 */
export async function runOcrWithLlm(
  images: ImageInput[],
  columns: string[],
  overrideConfig?: SimpleLlmConfig,
  onProgress?: (step: string) => void
): Promise<{ success: boolean; data?: any; error?: string }> {
  const activeConfig = getActiveProviderConfig();
  const config: SimpleLlmConfig = overrideConfig || {
    apiKey: activeConfig.apiKey,
    model: activeConfig.model,
    baseUrl: activeConfig.baseUrl,
    proxy: activeConfig.proxy,
    provider: activeConfig.providerName,
  };

  if (!config.apiKey) {
    return { success: false, error: '未配置当前激活模型服务商的 API Key，请先前往“系统设置”填写 API Key。' };
  }
  if (!columns || columns.length === 0) {
    return { success: false, error: '未配置识别列，请先在项目中勾选识别列。' };
  }

  const baseUrl = resolveBaseUrl(config);
  const prompt = buildOcrPrompt(columns, images.length);

  onProgress?.(`已构建 Prompt（共 ${images.length} 张图片），正在打包至单次 API 请求...`);

  // 构建单次请求的多图 payload content 数组
  const contentItems: any[] = [{ type: 'text', text: prompt }];

  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    const b64 = img.buffer.toString('base64');
    contentItems.push({
      type: 'image_url',
      image_url: {
        url: `data:${img.mimeType || 'image/jpeg'};base64,${b64}`,
      },
    });
  }

  const url = `${baseUrl}/chat/completions`;
  const axiosConfig = createAxiosConfig(config);

  const providerName = config.provider || 'API';
  const modelName = config.model || 'default';

  onProgress?.(`正在发送请求至大模型 [${providerName} - ${modelName}]...`);

  try {
    let response;
    try {
      // 尝试包含 response_format 的严格模式
      response = await axios.post(
        url,
        {
          model: modelName,
          messages: [{ role: 'user', content: contentItems }],
          response_format: { type: 'json_object' },
          temperature: 0.1,
        },
        axiosConfig
      );
    } catch (firstErr: any) {
      // 如果服务商/模型不支持 response_format，尝试回退降级模式
      console.warn('含有 response_format 请求失败，尝试退回通用模式:', firstErr.message);
      onProgress?.(`服务商 [${providerName}] 响应异常: ${firstErr.message}。正在尝试通用模式无约束重发...`);

      response = await axios.post(
        url,
        {
          model: modelName,
          messages: [{ role: 'user', content: contentItems }],
          temperature: 0.1,
        },
        axiosConfig
      );
    }

    onProgress?.('大模型数据接收完毕，正在校验提取 JSON 结构...');

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
    const errMsg = err.response?.data?.error?.message || err.response?.data?.detail || err.message || String(err);
    console.error('LLM OCR 识别失败:', errMsg);
    return { success: false, error: `服务商 [${providerName}] 识别请求失败: ${errMsg}` };
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
