import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest';
import path from 'path';
import fs from 'fs/promises';
import fsSync from 'fs';

const TEST_SETTINGS_FILE = path.join(process.cwd(), 'data', 'settings.test.json');

// 强制为测试指定独立的 settings.test.json 路径
process.env.SETTINGS_FILE_PATH = TEST_SETTINGS_FILE;

import { loadSettings, saveSettings } from './settings.js';

describe('Settings Service (data/settings.test.json 隔离测试)', () => {
  beforeAll(async () => {
    // 确保清理旧的测试临时文件
    if (fsSync.existsSync(TEST_SETTINGS_FILE)) {
      await fs.unlink(TEST_SETTINGS_FILE).catch(() => {});
    }
  });

  afterAll(async () => {
    // 测试完成后清理测试文件
    if (fsSync.existsSync(TEST_SETTINGS_FILE)) {
      await fs.unlink(TEST_SETTINGS_FILE).catch(() => {});
    }
  });

  beforeEach(async () => {
    if (fsSync.existsSync(TEST_SETTINGS_FILE)) {
      await fs.unlink(TEST_SETTINGS_FILE).catch(() => {});
    }
  });

  it('应该能自动加载默认配置并创建 data/settings.test.json 文件', async () => {
    const settings = await loadSettings();
    expect(settings).toBeDefined();
    expect(settings.llm).toBeDefined();
    expect(settings.llm.activeProvider).toBe('google');
    expect(settings.llm.providers.google.model).toBe('gemini-3.6-flash');
    expect(fsSync.existsSync(TEST_SETTINGS_FILE)).toBe(true);
  });

  it('应该能够成功保存多 Provider 独立 Key 配置并持久化', async () => {
    await loadSettings();
    const updated = await saveSettings({
      llm: {
        activeProvider: 'deepseek',
        providers: {
          google: { apiKey: 'google-key-111', model: 'gemini-3.6-flash', baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai', proxy: '' },
          deepseek: { apiKey: 'deepseek-key-222', model: 'deepseek-v4', baseUrl: 'https://api.deepseek.com/v1', proxy: '' },
          mimo: { apiKey: 'mimo-key-333', model: 'mimo-v2.5', baseUrl: 'https://api.xiaomimimo.com/v1', proxy: '' },
          custom: { apiKey: '', model: '', baseUrl: '', proxy: '' },
        },
      },
    });

    expect(updated.llm.activeProvider).toBe('deepseek');
    expect(updated.llm.providers.google.apiKey).toBe('google-key-111');
    expect(updated.llm.providers.deepseek.apiKey).toBe('deepseek-key-222');

    // 重新读取磁盘文件，验证持久化正确性
    const diskContent = JSON.parse(await fs.readFile(TEST_SETTINGS_FILE, 'utf-8'));
    expect(diskContent.llm.activeProvider).toBe('deepseek');
    expect(diskContent.llm.providers.deepseek.apiKey).toBe('deepseek-key-222');
  });

  it('在内存/进程启动时应该能自动从多 Provider settings.test.json 完整恢复配置', async () => {
    // 模拟写入既有文件
    const mockSettings = {
      llm: {
        activeProvider: 'mimo',
        providers: {
          google: { apiKey: '', model: 'gemini-3.6-flash', baseUrl: '', proxy: '' },
          mimo: { apiKey: 'mimo-key-999', model: 'mimo-v2.5', baseUrl: 'https://api.xiaomimimo.com/v1', proxy: '' },
        },
      },
    };
    await fs.mkdir(path.dirname(TEST_SETTINGS_FILE), { recursive: true });
    await fs.writeFile(TEST_SETTINGS_FILE, JSON.stringify(mockSettings), 'utf-8');

    // 加载配置
    const loaded = await loadSettings();
    expect(loaded.llm.activeProvider).toBe('mimo');
    expect(loaded.llm.providers.mimo.apiKey).toBe('mimo-key-999');
  });
});

