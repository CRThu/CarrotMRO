import { render, screen } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { RateCardWorkspace } from './RateCardWorkspace'

describe('RateCardWorkspace Component', () => {
  const defaultProps = {
    currentRateCard: '2026协议定价表.json',
    ratecardTableData: {
      columns: ['项目名称', '单位', '不含税单价', '说明'],
      items: [
        { '项目名称': '铜芯电缆', '单位': '米', '不含税单价': '45.00', '说明': '国标' }
      ]
    },
    onRefreshData: vi.fn(),
  }

  it('1. 正常渲染协议定价表列表与标题', () => {
    render(<RateCardWorkspace {...defaultProps} />)
    expect(screen.getByText('协议定价表: 2026协议定价表.json')).toBeInTheDocument()
    expect(screen.getByText('铜芯电缆')).toBeInTheDocument()
    expect(screen.getByText('45.00')).toBeInTheDocument()
  })

  it('2. 包含 Excel / CSV 导入按钮', () => {
    render(<RateCardWorkspace {...defaultProps} />)
    expect(screen.getByText('导入 Excel / CSV')).toBeInTheDocument()
  })
})
