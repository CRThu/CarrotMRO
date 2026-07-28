import { describe, it, expect } from 'vitest';
import { extractJsonFromText, buildOcrPrompt } from './llm.js';

describe('LLM Service Tools', () => {
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
    const prompt = buildOcrPrompt(columns);
    expect(prompt).toContain('物料名称、数量、综合单价');
    expect(prompt).toContain('"物料名称": ""');
    expect(prompt).toContain('"数量": ""');
    expect(prompt).toContain('"综合单价": ""');
  });
});
