import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { Sidebar } from './Sidebar'

describe('Sidebar Component', () => {
  const defaultProps = {
    projects: ['项目A', '项目B'],
    rateCards: ['定价表1'],
    quotationFiles: ['quotation-1.json'],
    currentProject: '项目A',
    currentRateCard: null,
    currentView: 'project_config' as const,
    activeQuotationFilename: null,
    expandedSections: new Set(['projects', 'ratecards']),
    expandedProject: '项目A',
    expandedQuotation: null,
    onToggleSection: vi.fn(),
    onSelectProjectConfig: vi.fn(),
    onSelectRateCard: vi.fn(),
    onSelectSettings: vi.fn(),
    onCreateProject: vi.fn(),
    onCreateRateCard: vi.fn(),
    onToggleProject: vi.fn(),
    onToggleQuotation: vi.fn(),
    onQuotationSelectFile: vi.fn(),
    onQuotationDeleteFile: vi.fn(),
    onQuotationCreate: vi.fn(),
  }

  it('1. 边界防御测试：当 API 数据为 undefined 时拥有防崩溃默认空数组机制', () => {
    render(
      <Sidebar
        {...defaultProps}
        projects={undefined as any}
        rateCards={undefined as any}
        quotationFiles={undefined as any}
      />
    )
    expect(screen.getByText('项目列表')).toBeInTheDocument()
    expect(screen.getByText('协议定价表')).toBeInTheDocument()
  })

  it('2. 正常渲染与点击分类展开', () => {
    render(<Sidebar {...defaultProps} />)

    expect(screen.getByText('项目列表')).toBeInTheDocument()
    expect(screen.getByText('协议定价表')).toBeInTheDocument()

    // 点击切换分类
    fireEvent.click(screen.getByText('项目列表'))
    expect(defaultProps.onToggleSection).toHaveBeenCalledWith('projects')
  })

  it('3. 视图切换验证：在系统设置视图下点击侧边栏已有项目触发 onToggleProject', () => {
    render(
      <Sidebar
        {...defaultProps}
        currentView="settings"
        currentProject="项目A"
      />
    )

    // 点击侧边栏项目 A 触发选择与视图恢复
    const projItem = screen.getByText('项目A')
    fireEvent.click(projItem)
    expect(defaultProps.onToggleProject).toHaveBeenCalledWith('项目A')
  })

  it('4. 点击系统 LLM 设置选项触发 onSelectSettings', () => {
    render(<Sidebar {...defaultProps} />)
    const settingsBtn = screen.getByText('系统 LLM 设置')
    fireEvent.click(settingsBtn)
    expect(defaultProps.onSelectSettings).toHaveBeenCalled()
  })
})
