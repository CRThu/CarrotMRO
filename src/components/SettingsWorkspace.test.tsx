import { describe, it, expect, vi, beforeEach } from 'vitest';
import { render, screen, fireEvent, waitFor } from '@testing-library/react';
import { SettingsWorkspace } from './SettingsWorkspace';
import * as api from '@/api';

vi.mock('@/api', () => ({
  getSettings: vi.fn(),
  updateSettings: vi.fn(),
  testLlmConfig: vi.fn(),
}));

describe('SettingsWorkspace Component', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('渲染设置界面并从服务端获取多 Key LLM 配置', async () => {
    (api.getSettings as any).mockResolvedValue({
      data: {
        llm: {
          activeProvider: 'google',
          providers: {
            google: { apiKey: 'sk-google-123', model: 'gemini-3.6-flash', baseUrl: '', proxy: '' },
            deepseek: { apiKey: 'sk-deepseek-456', model: 'deepseek-v4', baseUrl: '', proxy: '' },
          },
        },
      },
    });

    render(<SettingsWorkspace />);

    await waitFor(() => {
      expect(screen.getByPlaceholderText(/输入 Google Gemini API Key/i)).toHaveValue('sk-google-123');
    });

    expect(screen.getByText(/LLM 模型服务商接入/i)).toBeInTheDocument();
  });

  it('能够修改当前 Provider Key 配置并保存所有设置', async () => {
    (api.getSettings as any).mockResolvedValue({
      data: {
        llm: {
          activeProvider: 'google',
          providers: {
            google: { apiKey: '', model: 'gemini-3.6-flash', baseUrl: '', proxy: '' },
          },
        },
      },
    });

    (api.updateSettings as any).mockResolvedValue({
      data: { success: true },
    });

    render(<SettingsWorkspace />);

    await waitFor(() => {
      expect(screen.getByPlaceholderText(/输入 Google Gemini API Key/i)).toBeInTheDocument();
    });

    const apiKeyInput = screen.getByPlaceholderText(/输入 Google Gemini API Key/i);
    fireEvent.change(apiKeyInput, { target: { value: 'sk-my-new-key' } });

    const saveBtn = screen.getByRole('button', { name: /保存配置/i });
    fireEvent.click(saveBtn);

    await waitFor(() => {
      expect(api.updateSettings).toHaveBeenCalled();
      expect(screen.getByText(/配置已/i)).toBeInTheDocument();
    });
  });

  it('点击切换服务商卡片时能够触发即选即存自动保存', async () => {
    (api.getSettings as any).mockResolvedValue({
      data: {
        llm: {
          activeProvider: 'google',
          providers: {
            google: { apiKey: 'sk-google-123', model: 'gemini-3.6-flash', baseUrl: '', proxy: '' },
            deepseek: { apiKey: 'sk-deepseek-456', model: 'deepseek-v4', baseUrl: '', proxy: '' },
          },
        },
      },
    });

    (api.updateSettings as any).mockResolvedValue({
      data: { success: true },
    });

    render(<SettingsWorkspace />);

    await waitFor(() => {
      expect(screen.getByText('Google Gemini')).toBeInTheDocument();
    });

    // 点击 DeepSeek 服务商卡片
    const deepseekCard = screen.getByText('DeepSeek');
    fireEvent.click(deepseekCard);

    await waitFor(() => {
      expect(api.updateSettings).toHaveBeenCalledWith({
        llm: expect.objectContaining({
          activeProvider: 'deepseek',
        }),
      });
      expect(screen.getByText(/配置已自动保存/i)).toBeInTheDocument();
    });
  });
});
