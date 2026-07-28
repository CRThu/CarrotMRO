import { describe, it, expect, vi, beforeEach } from 'vitest'

const mockGet = vi.fn()
const mockPost = vi.fn()
const mockPut = vi.fn()
const mockPatch = vi.fn()
const mockDelete = vi.fn()

vi.mock('axios', () => ({
  default: {
    get: (...args: any[]) => mockGet(...args),
    post: (...args: any[]) => mockPost(...args),
    put: (...args: any[]) => mockPut(...args),
    patch: (...args: any[]) => mockPatch(...args),
    delete: (...args: any[]) => mockDelete(...args),
  },
}))

import * as api from './api'

describe('API', () => {
  beforeEach(() => {
    vi.clearAllMocks()
  })

  describe('Preset Columns', () => {
    it('getPresetColumns calls GET /api/preset-columns', async () => {
      mockGet.mockResolvedValue({ data: ['项目组', '项目名称'] })
      await api.getPresetColumns()
      expect(mockGet).toHaveBeenCalledWith('/api/preset-columns')
    })
  })

  describe('Projects and Settings', () => {
    it('getProjects calls GET /api/projects', async () => {
      mockGet.mockResolvedValue({ data: ['proj1'] })
      await api.getProjects()
      expect(mockGet).toHaveBeenCalledWith('/api/projects')
    })

    it('createProject calls POST /api/projects/{name}', async () => {
      mockPost.mockResolvedValue({ data: { success: true, name: 'test-project' } })
      await api.createProject('test-project')
      expect(mockPost).toHaveBeenCalledWith('/api/projects/test-project')
    })

    it('getProjectInfo calls GET /api/projects/{name}', async () => {
      mockGet.mockResolvedValue({ data: { name: 'test-project', ocr_columns: [] } })
      await api.getProjectInfo('test-project')
      expect(mockGet).toHaveBeenCalledWith('/api/projects/test-project')
    })

    it('updateProjectSettings calls PATCH /api/projects/{name}/settings', async () => {
      mockPatch.mockResolvedValue({ data: { success: true } })
      await api.updateProjectSettings('test-project', { ratecard_name: 'rc1' })
      expect(mockPatch).toHaveBeenCalledWith('/api/projects/test-project/settings', { ratecard_name: 'rc1' })
    })
  })

  describe('RateCards', () => {
    it('getRateCards calls GET /api/ratecards', async () => {
      mockGet.mockResolvedValue({ data: ['rc1'] })
      await api.getRateCards()
      expect(mockGet).toHaveBeenCalledWith('/api/ratecards')
    })

    it('previewRateCardImport calls POST /api/ratecards/{name}/import-preview', async () => {
      mockPost.mockResolvedValue({ data: { headers: ['名称'], sampleRows: [] } })
      const formData = new FormData()
      await api.previewRateCardImport('rc1', formData)
      expect(mockPost).toHaveBeenCalledWith('/api/ratecards/rc1/import-preview', formData)
    })

    it('importRateCardFile calls POST /api/ratecards/{name}/import', async () => {
      mockPost.mockResolvedValue({ data: { success: true, count: 10 } })
      const payload = { headers: ['名称'], items: [], mapping: { 名称: '项目名称' } }
      await api.importRateCardFile('rc1', payload)
      expect(mockPost).toHaveBeenCalledWith('/api/ratecards/rc1/import', payload)
    })
  })

  describe('System Settings and LLM aliases', () => {
    it('updateSettings calls PUT /api/settings', async () => {
      mockPut.mockResolvedValue({ data: { success: true } })
      await api.updateSettings({ llm: {} })
      expect(mockPut).toHaveBeenCalledWith('/api/settings', { llm: {} })
    })

    it('testLlmConfig calls POST /api/settings/test-llm', async () => {
      mockPost.mockResolvedValue({ data: { success: true } })
      await api.testLlmConfig({ apiKey: '123' })
      expect(mockPost).toHaveBeenCalledWith('/api/settings/test-llm', { apiKey: '123' })
    })
  })
})
