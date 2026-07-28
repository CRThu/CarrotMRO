import { describe, it, expect, beforeAll, afterAll, beforeEach } from 'vitest';
import { loadSettings, saveSettings } from './settings.js';
import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';

const SETTINGS_FILE = path.join(process.cwd(), 'data', 'settings.json');
let originalContent: string | null = null;

describe('Settings Service (data/settings.json)', () => {
  beforeAll(async () => {
    // 备份用户真实的 settings.json 文件，避免测试执行破坏真实配置
    if (fsSync.existsSync(SETTINGS_FILE)) {
      originalContent = await fs.readFile(SETTINGS_FILE, 'utf-8').catch(() => null);
    }
  });

  afterAll(async () => {
    // 还原用户真实的 settings.json 文件
    if (originalContent !== null) {
      await fs.writeFile(SETTINGS_FILE, originalContent, 'utf-8').catch(() => {});
    }
  });

  beforeEach(async () => {
    // 测试前临时清理
    if (fsSync.existsSync(SETTINGS_FILE)) {
      await fs.unlink(SETTINGS_FILE).catch(() => {});
    }
  });

  it('应该能自动加载默认配置并创建 data/settings.json 文件', async () => {
    const settings = await loadSettings();
    expect(settings).toBeDefined();
    expect(settings.llm).toBeDefined();
    expect(settings.llm.activeProvider).toBe('google');
    expect(settings.llm.providers.google.model).toBe('gemini-3.6-flash');
    expect(fsSync.existsSync(SETTINGS_FILE)).toBe(true);
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
    const diskContent = JSON.parse(await fs.readFile(SETTINGS_FILE, 'utf-8'));
    expect(diskContent.llm.activeProvider).toBe('deepseek');
    expect(diskContent.llm.providers.deepseek.apiKey).toBe('deepseek-key-222');
  });

  it('在内存/进程启动时应该能自动从多 Provider settings.json 完整恢复配置', async () => {
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
    await fs.mkdir(path.dirname(SETTINGS_FILE), { recursive: true });
    await fs.writeFile(SETTINGS_FILE, JSON.stringify(mockSettings), 'utf-8');

    // 加载配置
    const loaded = await loadSettings();
    expect(loaded.llm.activeProvider).toBe('mimo');
    expect(loaded.llm.providers.mimo.apiKey).toBe('mimo-key-999');
  });
});
