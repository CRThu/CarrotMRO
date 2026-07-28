import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { Sidebar } from './Sidebar'

describe('Sidebar Component', () => {
  const defaultProps = {
    projects: ['项目A', '项目B'],
    rateCards: ['定价表1'],
    ocrFiles: ['ocr-1.json'],
    quotationFiles: ['quotation-1.json'],
    currentProject: '项目A',
    currentRateCard: null,
    currentView: 'project' as const,
    activeFilename: null,
    activeQuotationFilename: null,
    expandedSections: new Set(['projects', 'ratecards']),
    expandedProject: '项目A',
    expandedOcr: null,
    expandedQuotation: null,
    onToggleSection: vi.fn(),
    onSelectProject: vi.fn(),
    onSelectRateCard: vi.fn(),
    onCreateProject: vi.fn(),
    onCreateRateCard: vi.fn(),
    onDeleteOcrFile: vi.fn(),
    onDeleteQuotationFile: vi.fn(),
    onSelectOcrFile: vi.fn(),
    onSelectQuotationFile: vi.fn(),
    onCreateQuotation: vi.fn(),
    onToggleProject: vi.fn(),
    onToggleOcrFolder: vi.fn(),
    onToggleQuotationFolder: vi.fn()
  }

  it('1. 边界防御测试：当 API 数据为 undefined 时拥有防崩溃默认空数组机制', () => {
    render(
      <Sidebar
        {...defaultProps}
        projects={undefined as any}
        rateCards={undefined as any}
        ocrFiles={undefined as any}
        quotationFiles={undefined as any}
      />
    )
    expect(screen.getByText('项目')).toBeInTheDocument()
    expect(screen.getByText('协议基准价格清单')).toBeInTheDocument()
  })

  it('2. 正常渲染与点击分类展开', () => {
    render(<Sidebar {...defaultProps} />)

    expect(screen.getByText('项目')).toBeInTheDocument()
    expect(screen.getByText('协议基准价格清单')).toBeInTheDocument()

    // 点击切换分类
    fireEvent.click(screen.getByText('项目'))
    expect(defaultProps.onToggleSection).toHaveBeenCalledWith('projects')
  })
})
