import { describe, it, expect, vi, beforeEach } from 'vitest'

// Mock functions must be defined inside vi.mock factory
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
      mockGet.mockResolvedValue({ data: { columns: [] } })
      await api.getPresetColumns()
      expect(mockGet).toHaveBeenCalledWith('/api/preset-columns')
    })
  })

  describe('Projects', () => {
    it('getProjects calls GET /api/projects', async () => {
      mockGet.mockResolvedValue({ data: { projects: [] } })
      await api.getProjects()
      expect(mockGet).toHaveBeenCalledWith('/api/projects')
    })

    it('createProject calls POST /api/projects/{name}', async () => {
      mockPost.mockResolvedValue({ data: { message: 'ok' } })
      await api.createProject('test-project')
      expect(mockPost).toHaveBeenCalledWith('/api/projects/test-project')
    })
  })

  describe('Templates', () => {
    it('getTemplates calls GET /api/templates', async () => {
      mockGet.mockResolvedValue({ data: { files: [] } })
      await api.getTemplates()
      expect(mockGet).toHaveBeenCalledWith('/api/templates')
    })
  })

  describe('RateCards', () => {
    it('getRateCards calls GET /api/ratecards', async () => {
      mockGet.mockResolvedValue({ data: { ratecards: [] } })
      await api.getRateCards()
      expect(mockGet).toHaveBeenCalledWith('/api/ratecards')
    })
  })

  describe('Project Columns', () => {
    it('getProjectColumns calls GET /api/projects/{name}/columns', async () => {
      mockGet.mockResolvedValue({ data: { available_columns: [], selected_columns: [], column_mappings: {} } })
      await api.getProjectColumns('test-project')
      expect(mockGet).toHaveBeenCalledWith('/api/projects/test-project/columns')
    })

    it('updateProjectColumns calls PATCH /api/projects/{name}/columns', async () => {
      mockPatch.mockResolvedValue({ data: { message: 'ok' } })
      await api.updateProjectColumns('test-project', { columns: ['col1'] })
      expect(mockPatch).toHaveBeenCalledWith('/api/projects/test-project/columns', { columns: ['col1'] })
    })
  })
})
