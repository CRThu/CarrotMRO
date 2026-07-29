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
  onProgress?: (step: string, rawChunk?: string) => void
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

  // 自动重试控制参数
  const maxRetries = 3;
  let lastError = '';

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const retryPrefix = attempt > 1 ? `[自动重试 ${attempt}/${maxRetries}] ` : '';
    onProgress?.(`${retryPrefix}正在发送请求至 [${providerName} - ${modelName}]...`);

    const startTime = Date.now();
    try {
      // 1. 尝试开启 response_format 强制 JSON 模式 + stream 流式
      let response;
      let isStreamSupported = true;

      try {
        response = await axios.post(
          url,
          {
            model: modelName,
            messages: [{ role: 'user', content: contentItems }],
            response_format: { type: 'json_object' }, // 协议级别标准 JSON 模式约束
            temperature: 0.1,
            stream: true,
          },
          {
            ...axiosConfig,
            responseType: 'stream',
          }
        );
      } catch (streamErr: any) {
        // 部分中转服务商不支持 response_format 或 stream，尝试降级
        console.warn('含有 response_format 的流式请求未获支持，降级尝试无约束流式:', streamErr.message);
        isStreamSupported = false;

        response = await axios.post(
          url,
          {
            model: modelName,
            messages: [{ role: 'user', content: contentItems }],
            temperature: 0.1,
            stream: true,
          },
          {
            ...axiosConfig,
            responseType: 'stream',
          }
        );
      }

      let fullText = '';
      let lastLogTime = Date.now();

      await new Promise<void>((resolve, reject) => {
        const stream = response.data;
        let buffer = '';

        stream.on('data', (chunk: Buffer) => {
          buffer += chunk.toString('utf-8');
          const lines = buffer.split('\n');
          buffer = lines.pop() || '';

          for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || trimmed === 'data: [DONE]') continue;
            if (trimmed.startsWith('data: ')) {
              try {
                const json = JSON.parse(trimmed.slice(6));
                const delta = json.choices?.[0]?.delta;
                const textChunk = delta?.content || delta?.reasoning_content || '';
                if (textChunk) {
                  fullText += textChunk;
                  const now = Date.now();
                  // 实时回调原始文本，供前端真实流打字渲染
                  onProgress?.('', textChunk);

                  if (now - lastLogTime > 400) {
                    lastLogTime = now;
                    const elapsedSec = ((now - startTime) / 1000).toFixed(1);
                    onProgress?.(`${retryPrefix}AI 实时提取中 (${elapsedSec}s) › 已输出 ${fullText.length} 字符`);
                  }
                }
              } catch {}
            }
          }
        });

        stream.on('end', () => resolve());
        stream.on('error', (err: any) => reject(err));
      });

      const parsedData = extractJsonFromText(fullText);
      if (!parsedData) {
        throw new Error(`无法从模型返回提取有效 JSON 格式数据 (返回长度: ${fullText.length})`);
      }

      let finalData = parsedData;
      if (!finalData.items) {
        finalData = {
          items: Array.isArray(parsedData) ? parsedData : [parsedData],
          remarks: '',
        };
      }
      if (!finalData.remarks) finalData.remarks = '';

      const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
      onProgress?.(`数据提取成功 (总耗时 ${elapsedSec}s)`);
      return { success: true, data: finalData };

    } catch (attemptErr: any) {
      lastError = attemptErr.response?.data?.error?.message || attemptErr.message || String(attemptErr);
      console.warn(`OCR 识别第 ${attempt} 次尝试失败:`, lastError);

      if (attempt < maxRetries) {
        onProgress?.(`[自动重试 ${attempt}/${maxRetries}] 遇到响应抖动: ${lastError}，等待 1 秒后自动重试...`);
        await new Promise(r => setTimeout(r, 1000));
      }
    }
  }

  return { success: false, error: `服务商 [${providerName}] 经过 ${maxRetries} 次重试依然失败: ${lastError}` };
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
