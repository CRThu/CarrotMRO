import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { MatchPopover } from './MatchPopover'

describe('MatchPopover Component', () => {
  const mockCandidates = [
    {
      name: '国标铜芯电缆',
      score: 1.0,
      columns: ['单位', '不含税单价'],
      values: ['米', '45.00'],
      itemData: { 项目名称: '国标铜芯电缆', 单位: '米', 不含税单价: '45.00' }
    },
    {
      name: '铝芯电缆',
      score: 0.8,
      columns: ['单位', '不含税单价'],
      values: ['米', '25.00'],
      itemData: { 项目名称: '铝芯电缆', 单位: '米', 不含税单价: '25.00' }
    }
  ]

  it('应该能正常渲染匹配状态图标与按钮', () => {
    render(
      <MatchPopover
        status="matched"
        itemName="铜芯电缆"
        baseName="铜芯电缆"
        candidates={mockCandidates}
        onOpen={vi.fn()}
        onSelect={vi.fn()}
        onMarkCustom={vi.fn()}
      />
    )
    expect(screen.getByTitle('点击打开物料匹配与搜索')).toBeDefined()
  })

  it('点击触发按钮时应该调用 onOpen 展开弹窗', () => {
    const onOpen = vi.fn()
    render(
      <MatchPopover
        status="pending"
        itemName="铜芯电缆"
        baseName="铜芯电缆"
        candidates={mockCandidates}
        onOpen={onOpen}
        onSelect={vi.fn()}
        onMarkCustom={vi.fn()}
      />
    )
    const button = screen.getByTitle('点击打开物料匹配与搜索')
    fireEvent.click(button)
    expect(onOpen).toHaveBeenCalled()
  })

  it('在弹窗中搜索输入框变动时应该触发 onSearch', async () => {
    vi.useFakeTimers()
    const onSearch = vi.fn()
    render(
      <MatchPopover
        status="pending"
        itemName="电缆"
        baseName="电缆"
        candidates={mockCandidates}
        onOpen={vi.fn()}
        onSearch={onSearch}
        onSelect={vi.fn()}
        onMarkCustom={vi.fn()}
      />
    )
    const button = screen.getByTitle('点击打开物料匹配与搜索')
    fireEvent.click(button)

    const input = screen.getByPlaceholderText(/检索/i)
    fireEvent.change(input, { target: { value: '铝芯' } })

    // 快进防抖定时器 (250ms)
    vi.advanceTimersByTime(300)
    expect(onSearch).toHaveBeenCalledWith('铝芯')
    vi.useRealTimers()
  })

  it('点击候选物料列表条目时应该触发 onSelect 并选定数据', () => {
    const onSelect = vi.fn()
    render(
      <MatchPopover
        status="pending"
        itemName="电缆"
        baseName="电缆"
        candidates={mockCandidates}
        onOpen={vi.fn()}
        onSelect={onSelect}
        onMarkCustom={vi.fn()}
      />
    )
    const button = screen.getByTitle('点击打开物料匹配与搜索')
    fireEvent.click(button)

    const candidateItem = screen.getByText('国标铜芯电缆')
    fireEvent.click(candidateItem)
    expect(onSelect).toHaveBeenCalledWith(mockCandidates[0])
  })

  it('点击标记为自定义按钮时应该触发 onMarkCustom', () => {
    const onMarkCustom = vi.fn()
    render(
      <MatchPopover
        status="pending"
        itemName="非标定制件"
        baseName="非标定制件"
        candidates={[]}
        onOpen={vi.fn()}
        onSelect={vi.fn()}
        onMarkCustom={onMarkCustom}
      />
    )
    const button = screen.getByTitle('点击打开物料匹配与搜索')
    fireEvent.click(button)

    const customButton = screen.getByRole('button', { name: /不匹配/i })
    fireEvent.click(customButton)
    expect(onMarkCustom).toHaveBeenCalled()
  })
})
