import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';

export interface ProviderConfig {
  apiKey: string;
  model: string;
  baseUrl: string;
  proxy?: string;
}

export interface LlmConfig {
  activeProvider: string; // 'google' | 'deepseek' | 'mimo' | 'custom'
  providers: Record<string, ProviderConfig>;
}

export interface AppSettings {
  llm: LlmConfig;
}

export const DEFAULT_SETTINGS: AppSettings = {
  llm: {
    activeProvider: 'google',
    providers: {
      google: {
        apiKey: '',
        model: 'gemini-3.6-flash',
        baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai',
        proxy: '',
      },
      deepseek: {
        apiKey: '',
        model: 'deepseek-v4',
        baseUrl: 'https://api.deepseek.com/v1',
        proxy: '',
      },
      mimo: {
        apiKey: '',
        model: 'mimo-v2.5',
        baseUrl: 'https://api.xiaomimimo.com/v1',
        proxy: '',
      },
      custom: {
        apiKey: '',
        model: '',
        baseUrl: '',
        proxy: '',
      },
    },
  },
};

const DATA_DIR = path.join(process.cwd(), 'data');
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');

let cachedSettings: AppSettings | null = null;

function parseSettings(parsed: any): AppSettings {
  const settings: AppSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  if (!parsed || !parsed.llm) return settings;

  const rawLlm = parsed.llm;
  if (rawLlm.activeProvider) {
    settings.llm.activeProvider = rawLlm.activeProvider;
  }
  if (rawLlm.providers) {
    settings.llm.providers = {
      ...settings.llm.providers,
      ...rawLlm.providers,
    };
  }
  return settings;
}

/**
 * 启动或读取设置时，从 data/settings.json 加载
 */
export async function loadSettings(): Promise<AppSettings> {
  try {
    await fs.mkdir(DATA_DIR, { recursive: true });
    if (fsSync.existsSync(SETTINGS_FILE)) {
      const content = await fs.readFile(SETTINGS_FILE, 'utf-8');
      cachedSettings = parseSettings(JSON.parse(content));
    } else {
      cachedSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
      await fs.writeFile(SETTINGS_FILE, JSON.stringify(cachedSettings, null, 2), 'utf-8');
    }
  } catch (err) {
    console.error('加载 settings.json 失败，使用默认配置:', err);
    cachedSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
  }
  return cachedSettings || DEFAULT_SETTINGS;
}

/**
 * 获取当前内存/持久化设置
 */
export function getSettings(): AppSettings {
  if (!cachedSettings) {
    try {
      if (fsSync.existsSync(SETTINGS_FILE)) {
        const content = fsSync.readFileSync(SETTINGS_FILE, 'utf-8');
        cachedSettings = parseSettings(JSON.parse(content));
      } else {
        cachedSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
      }
    } catch {
      cachedSettings = JSON.parse(JSON.stringify(DEFAULT_SETTINGS));
    }
  }
  return cachedSettings || DEFAULT_SETTINGS;
}

/**
 * 获取当前激活生效的特定 Provider 配置
 */
export function getActiveProviderConfig(): ProviderConfig & { providerName: string } {
  const settings = getSettings();
  const active = settings.llm.activeProvider || 'google';
  const config = settings.llm.providers?.[active] || { apiKey: '', model: '', baseUrl: '', proxy: '' };
  return {
    ...config,
    providerName: active,
  };
}

/**
 * 保存配置更新并写入 data/settings.json
 */
export async function saveSettings(newSettings: Partial<AppSettings>): Promise<AppSettings> {
  const current = getSettings();
  const updated: AppSettings = {
    ...current,
    ...newSettings,
    llm: {
      ...current.llm,
      ...(newSettings.llm || {}),
      providers: {
        ...current.llm.providers,
        ...(newSettings.llm?.providers || {}),
      },
    },
  };

  await fs.mkdir(DATA_DIR, { recursive: true });
  await fs.writeFile(SETTINGS_FILE, JSON.stringify(updated, null, 2), 'utf-8');
  cachedSettings = updated;
  return updated;
}
