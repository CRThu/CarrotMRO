import { describe, it, expect, vi, beforeEach } from 'vitest';
import { extractJsonFromText, buildOcrPrompt, runOcrWithLlm } from './llm.js';
import axios from 'axios';
import { Readable } from 'stream';

vi.mock('axios');

describe('LLM Service Tools & OCR Engine', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('应该能正确解析各种 LLM 返回格式中的 JSON 提取', () => {
    // 1. 直接 JSON
    const json1 = extractJsonFromText('{"items": [{"物料": "螺丝"}]}');
    expect(json1).toEqual({ items: [{ 物料: '螺丝' }] });

    // 2. ```json ... ``` 包裹
    const json2 = extractJsonFromText('```json\n{"items": [{"物料": "螺母"}]}\n```');
    expect(json2).toEqual({ items: [{ 物料: '螺母' }] });

    // 3. 混合 Markdown 文本中的 JSON 提取
    const json3 = extractJsonFromText('识别结果如下：\n{\n  "items": [{"物料": "垫圈"}]\n}\n请核对。');
    expect(json3).toEqual({ items: [{ 物料: '垫圈' }] });
  });

  it('应该能够正确构建多字段动态 Prompt', () => {
    const columns = ['物料名称', '数量', '综合单价'];
    const prompt = buildOcrPrompt(columns, 2);
    expect(prompt).toContain('物料名称、数量、综合单价');
    expect(prompt).toContain('上传了 2 张多页/多图报价单图片');
    expect(prompt).toContain('"物料名称": ""');
  });

  it('runOcrWithLlm 在未配置 API Key 或列为空时能自动拦截', async () => {
    // 未配置 Key
    const resNoKey = await runOcrWithLlm([], ['项目名称'], { apiKey: '', model: 'm', baseUrl: '' });
    expect(resNoKey.success).toBe(false);
    expect(resNoKey.error).toContain('未配置当前激活模型服务商的 API Key');

    // 未配置列
    const resNoCol = await runOcrWithLlm([], [], { apiKey: 'sk-test', model: 'm', baseUrl: '' });
    expect(resNoCol.success).toBe(false);
    expect(resNoCol.error).toContain('未配置识别列');
  });

  it('图片 Base64 打包与流式 Stream 传输成功解析测试', async () => {
    (axios.post as any).mockImplementation(() => {
      const stream = new Readable({ read() {} });
      process.nextTick(() => {
        stream.push('data: {"choices":[{"delta":{"content":"{\\"items\\": [{\\"项目名称\\": \\"铜线\\", \\"不含税单价\\": \\"50\\"}]}"}}]}\n\n');
        stream.push('data: [DONE]\n\n');
        stream.push(null);
      });
      return Promise.resolve({ data: stream });
    });

    const mockImage = {
      buffer: Buffer.from('test-image-data'),
      mimeType: 'image/png',
    };

    const progressLogs: string[] = [];
    const onProgress = (msg: string) => progressLogs.push(msg);

    const config = {
      apiKey: 'sk-valid-key',
      model: 'gemini-3.6-flash',
      baseUrl: 'https://api.openai.com/v1',
      provider: 'google',
    };

    const result = await runOcrWithLlm([mockImage], ['项目名称', '不含税单价'], config, onProgress);

    expect(axios.post).toHaveBeenCalled();
    const requestPayload = (axios.post as any).mock.calls[0][1];
    expect(requestPayload.stream).toBe(true);
    // 验证图片 Base64 转换
    expect(requestPayload.messages[0].content[1].image_url.url).toContain('data:image/png;base64,');

    expect(result.success).toBe(true);
    expect(result.data.items[0]['项目名称']).toBe('铜线');
    expect(progressLogs.length).toBeGreaterThan(0);
  });

  it('当网络中断 / 401 鉴权错误 / 超时发生时，能优雅捕获并返回错误信息', async () => {
    (axios.post as any).mockRejectedValue({
      response: {
        data: {
          error: { message: 'Incorrect API key provided' },
        },
      },
    });

    const config = {
      apiKey: 'sk-invalid-key',
      model: 'deepseek-v4',
      baseUrl: 'https://api.deepseek.com/v1',
      provider: 'deepseek',
    };

    const result = await runOcrWithLlm(
      [{ buffer: Buffer.from('img'), mimeType: 'image/jpeg' }],
      ['项目名称'],
      config
    );

    expect(result.success).toBe(false);
    expect(result.error).toContain('Incorrect API key provided');
  }, 15000);

  it('当大模型输出非法 JSON 格式文本时，正确返回解析失败提示', async () => {
    (axios.post as any).mockImplementation(() => {
      const stream = new Readable({ read() {} });
      process.nextTick(() => {
        stream.push('data: {"choices":[{"delta":{"content":"抱歉，我无法识别这张图片的文字。"}}]}\n\n');
        stream.push('data: [DONE]\n\n');
        stream.push(null);
      });
      return Promise.resolve({ data: stream });
    });

    const config = { apiKey: 'sk-test', model: 'test', baseUrl: 'http://test' };
    const result = await runOcrWithLlm([{ buffer: Buffer.from('img'), mimeType: 'image/jpeg' }], ['项目名称'], config);

    expect(result.success).toBe(false);
    expect(result.error).toContain('无法从模型返回提取有效 JSON');
  });
});
