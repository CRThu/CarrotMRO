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
  };

  if (config.proxy && config.proxy.trim() !== '') {
    let rawProxy = config.proxy.trim();
    // 自动补全协议前缀，兼容用户输入的 127.0.0.1:7890 纯 IP 端口格式
    if (!rawProxy.startsWith('http://') && !rawProxy.startsWith('https://') && !rawProxy.startsWith('socks://')) {
      rawProxy = `http://${rawProxy}`;
    }

    try {
      const agent = new HttpsProxyAgent(rawProxy);
      axiosConfig.httpAgent = agent;
      axiosConfig.httpsAgent = agent;
      axiosConfig.proxy = false; // 禁用 Axios 自带 proxy 避免与 HttpsProxyAgent 发生冲突
    } catch (proxyErr) {
      console.warn('HttpsProxyAgent 代理设置失败:', proxyErr);
    }
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
你的任务是阅读并理解用户上传的报价单图片（包含打印件或手写装修/维保报价单），精确提取其中的表格项目数据。

${multiImageNote}【提取要求与规范】：
1. 字段提取：提取每个表格项的以下字段：${colList}。
2. 过滤序号：请自动去除“项目名称”内的前缀序号（例如将手写“1. 铺贴墙砖”提取为“铺贴墙砖”，切勿保留前缀“1.”或“一、”）。
3. 涂改与遗漏：请仔细辨认手写体字迹。若单据上有划线涂改废弃的项目请忽略；不要遗漏任何有效项目，也不要虚构项目。
4. 数值剥离：数量与单价尽量提取为纯数字或小数，单位提取为简短单位（例如将手写“15平米”剥离为数量 "15"、单位 "平米"）。
5. 缺失留空：若某些字段在图中未明示，保留为空字符串 ""。
6. 疑问汇总：将所有字迹模糊、缺失字段或无法确定的多条内容，逐条整理输出在 "remarks" 字符串数组列表中！

请严格按照以下标准 JSON 格式返回结果，不要包含任何 markdown 说明之外的代码块：
{
  "items": [
    { ${columns.map(col => `"${col}": ""`).join(', ')} }
  ],
  "remarks": [
    "识别提示或疑问列表项 1",
    "识别提示或疑问列表项 2"
  ]
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

  const maxRetries = 5;
  let lastError = '';

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    const retryPrefix = attempt > 1 ? `[自动重试 ${attempt}/${maxRetries}] ` : '';
    onProgress?.(`${retryPrefix}正在连接 API 服务商 [${providerName} - ${modelName}] (${baseUrl})...`);

    const startTime = Date.now();
    try {
      // 发送原生支持 response_format: { type: "json_object" } 的流式请求
      // 配置建连/握手超时为 15 秒，避免代理握手失败时卡死无限挂起
      const response = await axios.post(
        url,
        {
          model: modelName,
          messages: [{ role: 'user', content: contentItems }],
          response_format: { type: 'json_object' },
          temperature: 0.1,
          stream: true,
        },
        {
          ...axiosConfig,
          timeout: 15000, // 15 秒未建立 HTTP 握手连接立刻抛错进入下一次重试
          responseType: 'stream',
        }
      );

      onProgress?.(`${retryPrefix}已连接 API 端点，建立流打字传输通道...`);

      let fullText = '';
      let lastLogTime = Date.now();

      await new Promise<void>((resolve, reject) => {
        const stream = response.data;
        let buffer = '';

        // 智能流活动超时器：首包必须在 30 秒内产生；之后只要模型持续打字输出，每次新字符都会重置倒计时！
        let inactivityTimer: any = null;

        const resetInactivityTimer = (isFirst = false) => {
          if (inactivityTimer) clearTimeout(inactivityTimer);
          inactivityTimer = setTimeout(() => {
            stream.destroy();
            const timeType = isFirst ? '首包响应超时（30秒未收到首个字）' : '传输中断超时（30秒内无新字符输出）';
            reject(new Error(timeType));
          }, 30000);
        };

        // 启动首包等待倒计时 30秒
        resetInactivityTimer(true);

        stream.on('data', (chunk: Buffer) => {
          resetInactivityTimer(false); // 每次收到打字字符，立刻重置 30秒倒计时！打字不断流绝不中断！

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

                  if (now - lastLogTime > 800) {
                    lastLogTime = now;
                    const elapsedSec = ((now - startTime) / 1000).toFixed(1);
                    onProgress?.(`${retryPrefix}AI 实时提取中 (${elapsedSec}s) › 已输出 ${fullText.length} 字符`);
                  }
                }
              } catch {}
            }
          }
        });

        stream.on('end', () => {
          if (inactivityTimer) clearTimeout(inactivityTimer);
          resolve();
        });

        stream.on('error', (err: any) => {
          if (inactivityTimer) clearTimeout(inactivityTimer);
          reject(err);
        });
      });

      const parsedData = extractJsonFromText(fullText);
      if (!parsedData) {
        throw new Error(`无法从模型返回提取有效 JSON 格式数据 (返回长度: ${fullText.length})`);
      }

      let finalData = parsedData;
      if (!finalData.items) {
        finalData = {
          items: Array.isArray(parsedData) ? parsedData : [parsedData],
          remarks: [],
        };
      }

      // 归一化 remarks 为 string[]
      if (typeof finalData.remarks === 'string') {
        finalData.remarks = finalData.remarks.trim()
          ? finalData.remarks.split('\n').map((s: string) => s.trim()).filter(Boolean)
          : [];
      } else if (!Array.isArray(finalData.remarks)) {
        finalData.remarks = [];
      }

      const elapsedSec = ((Date.now() - startTime) / 1000).toFixed(1);
      onProgress?.(`数据提取成功 (总耗时 ${elapsedSec}s)`);
      return { success: true, data: finalData };

    } catch (attemptErr: any) {
      lastError = attemptErr.response?.data?.error?.message || attemptErr.message || String(attemptErr);
      console.warn(`OCR 识别第 ${attempt} 次尝试失败:`, lastError);

      if (attempt < maxRetries) {
        const delaySec = attempt + 1; // 递增退避：2s, 3s, 4s, 5s... 给网络充分恢复时间
        onProgress?.(`[自动重试 ${attempt}/${maxRetries}] 遇到网络抖动: ${lastError}，等待 ${delaySec} 秒后重置重试...`);
        await new Promise(r => setTimeout(r, delaySec * 1000));
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
