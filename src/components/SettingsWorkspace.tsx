import { useState, useEffect } from 'react';
import { AppSettings, LlmConfig, ProviderConfig } from '@/types';
import * as api from '@/api';
import { Key, Server, Cpu, Globe, Network, Save, CheckCircle2, AlertCircle, Eye, EyeOff, Radio } from 'lucide-react';

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
        setLlmConfig({
          activeProvider: raw.activeProvider || 'google',
          providers: {
            ...JSON.parse(JSON.stringify(DEFAULT_PROVIDERS)),
            ...(raw.providers || {}),
          },
        });
        setActiveTabProvider(raw.activeProvider || 'google');
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

  const updateCurrentProviderField = (field: keyof ProviderConfig, value: string) => {
    setLlmConfig(prev => ({
      ...prev,
      providers: {
        ...prev.providers,
        [activeTabProvider]: {
          ...(prev.providers[activeTabProvider] || { apiKey: '', model: '', baseUrl: '', proxy: '' }),
          [field]: value,
        },
      },
    }));
  };

  const handleSave = async () => {
    setSaving(true);
    setStatusMessage(null);
    try {
      const payload: AppSettings = { llm: llmConfig };
      const res = await api.updateSettings(payload);
      if (res.data?.success) {
        setStatusMessage({ type: 'success', text: '配置已成功保存！多服务商 API Key 已成功独立持久化至 data/settings.json。' });
      } else {
        setStatusMessage({ type: 'error', text: '保存失败: ' + (res.data?.detail || '未知错误') });
      }
    } catch (err: any) {
      setStatusMessage({ type: 'error', text: '保存网络错误: ' + (err.message || String(err)) });
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
        setStatusMessage({ type: 'success', text: res.data.message || `${activeTabProvider} API 连接成功！` });
      } else {
        setStatusMessage({ type: 'error', text: res.data?.message || 'API 测试失败' });
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
      <div className="w-full md:w-64 bg-slate-50 border-r border-gray-200 p-6 flex flex-col gap-2">
        <h2 className="text-xl font-medium text-slate-800 mb-4 flex items-center gap-2">
          <Server className="w-5 h-5 text-indigo-600" />
          系统设置
        </h2>
        <button
          onClick={() => setActiveGroup('api')}
          className={`flex items-center gap-3 px-4 py-3 rounded-lg text-sm font-medium transition-colors text-left ${
            activeGroup === 'api'
              ? 'bg-indigo-50 text-indigo-700 border border-indigo-200 shadow-sm'
              : 'text-gray-600 hover:bg-gray-100 hover:text-gray-900'
          }`}
        >
          <Cpu className="w-4 h-4" />
          API 接入设置
        </button>
      </div>

      {/* 右侧设置表单 */}
      <div className="flex-1 p-8 overflow-y-auto">
        {loading ? (
          <div className="flex items-center justify-center h-48 text-gray-500">加载设置中...</div>
        ) : (
          <div className="max-w-2xl space-y-6">
            <div>
              <h3 className="text-lg font-semibold text-gray-900">LLM 多模型 API 接入配置（支持多 Key 独立持久化）</h3>
              <p className="text-sm text-gray-500 mt-1">
                支持为 Google Gemini、DeepSeek、Xiaomi MiMo 及 Custom 模式独立配置与持久化对应的 API Key、Model 和接口参数。
              </p>
            </div>

            {statusMessage && (
              <div
                className={`p-4 rounded-md flex items-start gap-3 text-sm ${
                  statusMessage.type === 'success'
                    ? 'bg-emerald-50 text-emerald-800 border border-emerald-200'
                    : 'bg-rose-50 text-rose-800 border border-rose-200'
                }`}
              >
                {statusMessage.type === 'success' ? (
                  <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0 mt-0.5" />
                ) : (
                  <AlertCircle className="w-5 h-5 text-rose-600 shrink-0 mt-0.5" />
                )}
                <span>{statusMessage.text}</span>
              </div>
            )}

            {/* 服务商 Tab 切换与生效配置 */}
            <div className="space-y-4 bg-slate-50 p-6 rounded-xl border border-slate-200">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2 flex items-center gap-1.5">
                  <Globe className="w-4 h-4 text-indigo-500" />
                  选择配置服务商
                </label>
                <div className="grid grid-cols-2 md:grid-cols-4 gap-2">
                  {PROVIDER_OPTIONS.map(opt => {
                    const isActiveTarget = llmConfig.activeProvider === opt.id;
                    const isTabSelected = activeTabProvider === opt.id;
                    return (
                      <button
                        key={opt.id}
                        type="button"
                        onClick={() => setActiveTabProvider(opt.id)}
                        className={`p-3 rounded-lg border text-left flex flex-col justify-between transition-all ${
                          isTabSelected
                            ? 'bg-white border-indigo-600 shadow-sm ring-1 ring-indigo-600'
                            : 'bg-slate-100 border-gray-200 hover:bg-white text-gray-700'
                        }`}
                      >
                        <div className="font-medium text-xs flex items-center justify-between">
                          <span>{opt.name}</span>
                          {isActiveTarget && (
                            <span className="bg-emerald-100 text-emerald-800 text-[10px] px-1.5 py-0.5 rounded-full font-bold">
                              生效中
                            </span>
                          )}
                        </div>
                        <div className="text-[11px] text-gray-500 mt-2 truncate">
                          {llmConfig.providers[opt.id]?.apiKey ? 'Key 已配置' : '未配置 Key'}
                        </div>
                      </button>
                    );
                  })}
                </div>
              </div>

              {/* 设为当前系统默认生效的服务商 */}
              <div className="flex items-center justify-between bg-indigo-50 p-3 rounded-lg border border-indigo-100">
                <span className="text-xs text-indigo-900 font-medium">
                  当前编辑: <span className="font-bold">{activeProviderObj?.name}</span>
                  {llmConfig.activeProvider === activeTabProvider ? ' (正在作为系统实际生效服务商)' : ''}
                </span>
                {llmConfig.activeProvider !== activeTabProvider && (
                  <button
                    type="button"
                    onClick={() => setLlmConfig({ ...llmConfig, activeProvider: activeTabProvider })}
                    className="text-xs bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded font-medium transition-colors flex items-center gap-1 shadow-xs"
                  >
                    <Radio className="w-3.5 h-3.5" />
                    设为默认生效服务商
                  </button>
                )}
              </div>

              {/* API Key */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Key className="w-4 h-4 text-indigo-500" />
                  API Key ({activeProviderObj?.name})
                </label>
                <div className="relative">
                  <input
                    type={showApiKey ? 'text' : 'password'}
                    value={currentProviderConfig.apiKey || ''}
                    onChange={e => updateCurrentProviderField('apiKey', e.target.value)}
                    placeholder={`输入 ${activeProviderObj?.name} 对应的 API Key`}
                    className="w-full px-3 py-2 pr-10 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-sm"
                  />
                  <button
                    type="button"
                    onClick={() => setShowApiKey(!showApiKey)}
                    className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                  >
                    {showApiKey ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                  </button>
                </div>
              </div>

              {/* Model 名称 */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Cpu className="w-4 h-4 text-indigo-500" />
                  Model 名称
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.model || ''}
                  onChange={e => updateCurrentProviderField('model', e.target.value)}
                  placeholder={`输入 Model 名称 (默认: ${activeProviderObj?.defaultModel})`}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-sm"
                />
                {/* 快捷推荐模型 Pill 列表 */}
                {activeProviderObj?.presets && activeProviderObj.presets.length > 0 && (
                  <div className="flex flex-wrap gap-2 mt-2">
                    <span className="text-xs text-gray-500 self-center">推荐模型:</span>
                    {activeProviderObj.presets.map(m => (
                      <button
                        key={m}
                        type="button"
                        onClick={() => updateCurrentProviderField('model', m)}
                        className={`text-xs px-2 py-1 rounded transition-colors ${
                          currentProviderConfig.model === m
                            ? 'bg-indigo-600 text-white font-medium'
                            : 'bg-white border border-gray-300 text-gray-700 hover:bg-indigo-50 hover:text-indigo-600'
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
                <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Globe className="w-4 h-4 text-indigo-500" />
                  Base URL (接口基础地址)
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.baseUrl || ''}
                  onChange={e => updateCurrentProviderField('baseUrl', e.target.value)}
                  placeholder={`默认: ${activeProviderObj?.defaultUrl || '自定义端点 URL'}`}
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-sm"
                />
              </div>

              {/* Optional Proxy */}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1 flex items-center gap-1.5">
                  <Network className="w-4 h-4 text-indigo-500" />
                  Proxy 网络代理 (可选)
                </label>
                <input
                  type="text"
                  value={currentProviderConfig.proxy || ''}
                  onChange={e => updateCurrentProviderField('proxy', e.target.value)}
                  placeholder="如 http://127.0.0.1:7890，无代理请留空"
                  className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-white text-sm"
                />
              </div>
            </div>

            {/* 操作按钮区 */}
            <div className="flex items-center gap-4 pt-2">
              <button
                type="button"
                onClick={handleSave}
                disabled={saving}
                className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-md font-medium shadow-sm transition-colors disabled:opacity-50"
              >
                <Save className="w-4 h-4" />
                {saving ? '保存中...' : '保存所有配置'}
              </button>

              <button
                type="button"
                onClick={handleTestConnection}
                disabled={testing || !currentProviderConfig.apiKey}
                className="flex items-center gap-2 px-5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-md font-medium border border-slate-300 transition-colors disabled:opacity-50 text-sm"
              >
                {testing ? '测试中...' : `测试 ${activeProviderObj?.name} 连接`}
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
