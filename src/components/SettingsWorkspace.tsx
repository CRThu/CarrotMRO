import { useState, useEffect } from 'react';
import { AppSettings, LlmConfig, ProviderConfig } from '@/types';
import * as api from '@/api';
import {
  Key, Server, Cpu, Globe, Network, Save,
  CheckCircle2, AlertCircle, Eye, EyeOff, Radio, Activity, ShieldCheck,
} from 'lucide-react';

// 各服务商预设配置（含快捷推荐模型药丸列表）
const PROVIDER_OPTIONS = [
  {
    id: 'google',
    name: 'Google Gemini',
    defaultModel: 'gemini-3.6-flash',
    defaultUrl: 'https://generativelanguage.googleapis.com/v1beta/openai',
    presets: ['gemini-3.6-flash', 'gemini-3.5-flash-lite', 'gemini-3.5-flash'],
  },
  {
    id: 'deepseek',
    name: 'DeepSeek',
    defaultModel: 'deepseek-v4',
    defaultUrl: 'https://api.deepseek.com/v1',
    presets: ['deepseek-v4', 'deepseek-v4-pro'],
  },
  {
    id: 'mimo',
    name: 'Xiaomi MiMo',
    defaultModel: 'mimo-v2.5',
    defaultUrl: 'https://api.xiaomimimo.com/v1',
    presets: ['mimo-v2.5', 'mimo-v2.5-pro'],
  },
  {
    id: 'custom',
    name: 'Custom 自定义模式',
    defaultModel: '',
    defaultUrl: '',
    presets: [],
  },
];

const DEFAULT_PROVIDERS: Record<string, ProviderConfig> = {
  google: { apiKey: '', model: 'gemini-3.6-flash', baseUrl: 'https://generativelanguage.googleapis.com/v1beta/openai', proxy: '' },
  deepseek: { apiKey: '', model: 'deepseek-v4', baseUrl: 'https://api.deepseek.com/v1', proxy: '' },
  mimo: { apiKey: '', model: 'mimo-v2.5', baseUrl: 'https://api.xiaomimimo.com/v1', proxy: '' },
  custom: { apiKey: '', model: '', baseUrl: '', proxy: '' },
};

export function SettingsWorkspace() {
  const [activeGroup, setActiveGroup] = useState<'api'>('api');
  // 当前正在编辑的服务商 tab（独立于"激活"的服务商）
  const [activeTabProvider, setActiveTabProvider] = useState<string>('google');

  const [llmConfig, setLlmConfig] = useState<LlmConfig>({
    activeProvider: 'google',
    providers: JSON.parse(JSON.stringify(DEFAULT_PROVIDERS)),
  });

  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const [testing, setTesting] = useState(false);
  const [showApiKey, setShowApiKey] = useState(false);
  const [statusMessage, setStatusMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);

  useEffect(() => {
    fetchSettings();
  }, []);

  const fetchSettings = async () => {
    setLoading(true);
    try {
      const res = await api.getSettings();
      if (res.data?.llm) {
        const raw = res.data.llm;
        const active = raw.activeProvider || 'google';
        setLlmConfig({
          activeProvider: active,
          providers: {
            ...JSON.parse(JSON.stringify(DEFAULT_PROVIDERS)),
            ...(raw.providers || {}),
          },
        });
        setActiveTabProvider(active);
      }
    } catch (err: any) {
      setStatusMessage({ type: 'error', text: '加载设置失败: ' + (err.message || String(err)) });
    } finally {
      setLoading(false);
    }
  };

  // 当前正在编辑的 Provider 配置
  const currentProviderConfig: ProviderConfig = llmConfig.providers?.[activeTabProvider] || {
    apiKey: '',
    model: '',
    baseUrl: '',
    proxy: '',
  };

  // 快捷保存函数
  const autoSaveConfig = async (newConfig: LlmConfig) => {
    try {
      await api.updateSettings({ llm: newConfig });
      const activeName = PROVIDER_OPTIONS.find(p => p.id === newConfig.activeProvider)?.name || newConfig.activeProvider;
      setStatusMessage({ type: 'success', text: `配置已自动保存 (激活服务商: ${activeName})` });
    } catch (err: any) {
      setStatusMessage({ type: 'error', text: '自动保存失败: ' + (err.message || String(err)) });
    }
  };

  // 更新当前编辑 Provider 的某一字段
  const updateCurrentProviderField = (field: keyof ProviderConfig, value: string) => {
    setLlmConfig(prev => {
      const next = {
        ...prev,
        providers: {
          ...prev.providers,
          [activeTabProvider]: {
            ...(prev.providers[activeTabProvider] || { apiKey: '', model: '', baseUrl: '', proxy: '' }),
            [field]: value,
          },
        },
      };
      autoSaveConfig(next);
      return next;
    });
  };

  const handleSelectProvider = (id: string) => {
    setActiveTabProvider(id);
    setLlmConfig(prev => {
      const next = { ...prev, activeProvider: id };
      autoSaveConfig(next);
      return next;
    });
  };

  const handleSave = async () => {
    setSaving(true);
    setStatusMessage(null);
    try {
      const payload: AppSettings = { llm: llmConfig };
      const res = await api.updateSettings(payload);
      if (res.data?.success) {
        setStatusMessage({ type: 'success', text: '配置已成功保存！' });
      } else {
        setStatusMessage({ type: 'error', text: '保存失败: ' + (res.data?.detail || '未知错误') });
      }
    } catch (err: any) {
      setStatusMessage({ type: 'error', text: '保存失败: ' + (err.message || String(err)) });
    } finally {
      setSaving(false);
    }
  };

  const handleTestConnection = async () => {
    setTesting(true);
    setStatusMessage(null);
    try {
      const testPayload = {
        apiKey: currentProviderConfig.apiKey,
        model: currentProviderConfig.model,
        baseUrl: currentProviderConfig.baseUrl,
        proxy: currentProviderConfig.proxy,
        provider: activeTabProvider,
      };
      const res = await api.testLlmConfig(testPayload);
      if (res.data?.success) {
        setStatusMessage({ type: 'success', text: res.data.message || '连接测试成功！' });
      } else {
        setStatusMessage({ type: 'error', text: res.data?.message || '连接测试失败' });
      }
    } catch (err: any) {
      setStatusMessage({ type: 'error', text: '测试异常: ' + (err.message || String(err)) });
    } finally {
      setTesting(false);
    }
  };

  const activeProviderObj = PROVIDER_OPTIONS.find(p => p.id === activeTabProvider);

  return (
    <div className="bg-white rounded-lg shadow-sm border border-gray-200 overflow-hidden min-h-[600px] flex flex-col md:flex-row">
      {/* 左侧设置分组导览 */}
      <div className="w-full md:w-52 bg-slate-50 border-r border-gray-200 p-4 flex flex-col gap-2 shrink-0">
        <h2 className="text-base font-semibold text-slate-800 mb-2 flex items-center gap-2">
          <Server className="w-4 h-4 text-indigo-600" />
          系统设置
        </h2>
        <button
          onClick={() => setActiveGroup('api')}
          className={`flex items-center gap-2 px-3 py-2 rounded-lg text-sm font-medium transition-colors text-left ${
            activeGroup === 'api'
              ? 'bg-indigo-50 text-indigo-700 border border-indigo-200 shadow-sm'
              : 'text-gray-600 hover:bg-gray-100 hover:text-gray-900'
          }`}
        >
          <Cpu className="w-4 h-4" />
          API 服务商设置
        </button>
      </div>

      {/* 右侧设置表单 */}
      <div className="flex-1 p-6 overflow-y-auto">
        {loading ? (
          <div className="flex items-center justify-center h-48 text-gray-500">加载设置中...</div>
        ) : (
          <div className="max-w-2xl space-y-5">
            <div>
              <h3 className="text-base font-semibold text-gray-900">LLM 模型服务商接入</h3>
              <p className="text-xs text-gray-500 mt-0.5">
                配置各大模型 API 参数，点击选项卡自动保存并切换激活。
              </p>
            </div>

            {/* 状态提示 */}
            {statusMessage && (
              <div
                className={`px-3.5 py-2.5 rounded-lg flex items-center gap-2 text-xs ${
                  statusMessage.type === 'success'
                    ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                    : 'bg-rose-50 text-rose-800 border border-rose-200'
                }`}
              >
                {statusMessage.type === 'success' ? (
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                ) : (
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                )}
                <span>{statusMessage.text}</span>
              </div>
            )}

            {/* 服务商 Tab 切换卡片 */}
            <div className="space-y-4 bg-slate-50 p-5 rounded-xl border border-slate-200">
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-2 flex items-center gap-1.5">
                  <Globe className="w-3.5 h-3.5 text-indigo-500" />
                  切换服务商（点击自动保存生效）
                </label>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                  {PROVIDER_OPTIONS.map(opt => {
                    const isActive = llmConfig.activeProvider === opt.id;
                    return (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => handleSelectProvider(opt.id)}
                        className={`p-2.5 rounded-lg border text-left flex flex-col justify-between transition-all ${
                          isActive
                            ? 'bg-white border-indigo-600 shadow-sm ring-2 ring-indigo-600/30'
                            : 'bg-slate-100 border-gray-200 hover:bg-white text-gray-700'
                        }`}
                      >
                        <div className="font-medium text-xs flex items-center justify-between">
                          <span>{opt.name}</span>
                          {isActive && (
                            <span className="bg-emerald-100 text-emerald-800 text-[10px] px-1.5 py-0.5 rounded-full font-semibold flex items-center gap-1">
                              <Radio className="w-2.5 h-2.5 animate-pulse" />
                              激活
                            </span>
                          )}
                        </div>
                        <div className="text-[11px] text-gray-400 mt-1.5 truncate">
                          {llmConfig.providers[opt.id]?.apiKey ? '已配置 Key' : '未配置'}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* API Key */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Key className="w-3.5 h-3.5 text-indigo-500" />
                  API Key ({activeProviderObj?.name})
                </label>
                <div className="relative">
                  <input
                    type={showApiKey ? 'text' : 'password'}
                    value={currentProviderConfig.apiKey || ''}
                    onChange={e => updateCurrentProviderField('apiKey', e.target.value)}
                    placeholder={`输入 ${activeProviderObj?.name} API Key`}
                    className="w-full px-3 py-1.5 pr-10 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-xs"
                  />
                  <button
                    type="button"
                    onClick={() => setShowApiKey(!showApiKey)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                  >
                    {showApiKey ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                  </button>
                </div>
              </div>

              {/* Model 名称 + 快捷推荐模型 */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Cpu className="w-3.5 h-3.5 text-indigo-500" />
                  Model 名称
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.model || ''}
                  onChange={e => updateCurrentProviderField('model', e.target.value)}
                  placeholder={`输入 Model (默认: ${activeProviderObj?.defaultModel})`}
                  className="w-full px-3 py-1.5 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-xs"
                />
                {activeProviderObj?.presets && activeProviderObj.presets.length > 0 && (
                  <div className="flex flex-wrap gap-1.5 mt-2">
                    <span className="text-[11px] text-gray-400 self-center">推荐:</span>
                    {activeProviderObj.presets.map(m => (
                      <button
                        key={m}
                        type="button"
                        onClick={() => updateCurrentProviderField('model', m)}
                        className={`text-[11px] px-2 py-0.5 rounded-full border transition-colors ${
                          currentProviderConfig.model === m
                            ? 'bg-indigo-600 text-white font-medium border-indigo-600'
                            : 'bg-white border-gray-300 text-gray-700 hover:bg-indigo-50 hover:text-indigo-600'
                        }`}
                      >
                        {m}
                      </button>
                    ))}
                  </div>
                )}
              </div>

              {/* Base URL */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Globe className="w-3.5 h-3.5 text-indigo-500" />
                  Base URL
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.baseUrl || ''}
                  onChange={e => updateCurrentProviderField('baseUrl', e.target.value)}
                  placeholder={`默认: ${activeProviderObj?.defaultUrl || '接口地址'}`}
                  className="w-full px-3 py-1.5 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-xs font-mono"
                />
              </div>

              {/* Proxy */}
              <div>
                <label className="block text-xs font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <ShieldCheck className="w-3.5 h-3.5 text-indigo-500" />
                  HTTP 代理 (可选)
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.proxy || ''}
                  onChange={e => updateCurrentProviderField('proxy', e.target.value)}
                  placeholder="如 http://127.0.0.1:7890"
                  className="w-full px-3 py-1.5 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-xs font-mono"
                />
              </div>
            </div>

            {/* 操作按钮 */}
            <div className="flex items-center gap-3 pt-1">
              <button
                type="button"
                onClick={handleSave}
                disabled={saving}
                className="flex items-center gap-1.5 px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md text-xs font-medium shadow-sm transition-colors disabled:opacity-50"
              >
                <Save className="w-3.5 h-3.5" />
                {saving ? '保存中...' : '保存配置'}
              </button>

              <button
                type="button"
                onClick={handleTestConnection}
                disabled={testing || !currentProviderConfig.apiKey}
                className="flex items-center gap-1.5 px-4 py-2 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 rounded-md text-xs font-medium transition-colors disabled:opacity-50 shadow-sm"
              >
                <Activity className="w-3.5 h-3.5 text-indigo-600" />
                {testing ? '测试中...' : '测试连接'}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
