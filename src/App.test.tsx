import { render, screen, waitFor } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import App from './App'

vi.mock('@/api', () => ({
  getProjects: vi.fn().mockResolvedValue({ data: ['测试项目A', '测试项目B'] }),
  getRateCards: vi.fn().mockResolvedValue({ data: ['2026协议定价表.json'] }),
  getTemplates: vi.fn().mockResolvedValue({ data: ['标准模板.xlsx'] }),
  getQuotations: vi.fn().mockResolvedValue({
    data: { files: ['quotation-1.json'] }
  }),
  getProjectInfo: vi.fn().mockResolvedValue({
    data: {
      name: '测试项目A',
      created_at: new Date().toISOString(),
      ratecard_name: '2026协议定价表.json',
      template_name: '标准模板.xlsx',
      ocr_columns: ['项目名称', '单位', '数量', '不含税单价', '说明'],
      quotation_columns: ['项目组', '项目名称', '单位', '数量', '不含税单价', '不含税总价', '税率', '含税单价', '含税总价', '说明'],
    }
  }),
  getRateCardData: vi.fn().mockResolvedValue({
    data: { columns: ['项目名称', '单位', '不含税单价'], items: [] }
  }),
  createProject: vi.fn().mockResolvedValue({ data: { success: true } }),
  updateProjectSettings: vi.fn().mockResolvedValue({ data: { success: true } }),
  getSettings: vi.fn().mockResolvedValue({
    data: {
      llm: {
        activeProvider: 'google',
        providers: {
          google: { apiKey: 'sk-test', model: 'gemini-3.6-flash', baseUrl: '' }
        }
      }
    }
  })
}))

describe('App Main Application Component', () => {
  it('主应用初始化时应渲染侧边栏标题与全局主要容器', async () => {
    render(<App />)
    expect(screen.getByText('欢迎使用 CarrotMRO 管理系统')).toBeDefined()

    await waitFor(() => {
      expect(screen.getByText('项目列表')).toBeDefined()
    })
  })
})
