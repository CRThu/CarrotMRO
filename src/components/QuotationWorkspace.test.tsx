import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { QuotationWorkspace, calculateRowFormulas } from './QuotationWorkspace'

describe('QuotationWorkspace & Formulas', () => {
  it('1. 自动计算公式测试：含百分号税率正确计算三联动', () => {
    const raw = {
      '项目名称': '螺栓',
      '数量': '100',
      '不含税单价': '5.00',
      '税率': '13%', // 含 % 号格式
    }
    const calculated = calculateRowFormulas(raw)
    // 不含税总价 = 100 * 5.00 = 500.00
    expect(calculated['不含税总价']).toBe('500.00')
    // 含税单价 = 5.00 * 1.13 = 5.65
    expect(calculated['含税单价']).toBe('5.65')
    // 含税总价 = 100 * 5.65 = 565.00
    expect(calculated['含税总价']).toBe('565.00')
  })

  it('2. calculateRowFormulas：税率为数字（如 13 而非 13%）时也应正确换算', () => {
    // 边界：部分导入数据税率格式为纯数字 13（不带%）
    const raw = { '数量': '10', '不含税单价': '100', '税率': '13' }
    const calculated = calculateRowFormulas(raw)
    // 税率 13 → /100 = 0.13
    expect(calculated['含税单价']).toBe('113.00')
    expect(calculated['不含税总价']).toBe('1000.00')
    expect(calculated['含税总价']).toBe('1130.00')
  })

  it('3. calculateRowFormulas：缺少税率时不计算含税字段，仍计算不含税总价', () => {
    // 回归：无税率时不能用 NaN 覆盖含税字段
    const raw = { '数量': '5', '不含税单价': '200' }
    const calculated = calculateRowFormulas(raw)
    expect(calculated['不含税总价']).toBe('1000.00')
    // 无税率 → 含税字段不被覆盖（仍为原值或 undefined）
    expect(calculated['含税单价']).toBeUndefined()
    expect(calculated['含税总价']).toBeUndefined()
  })

  it('4. 渲染报价单表格与数据项（按 quotationColumns 过滤列）', () => {
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [
        { '项目名称': '阀门', '单位': '个', '数量': '10', '不含税单价': '100', '不含税总价': '1000' }
      ],
      projectRateCard: '协议表A.json',
      projectTemplate: '模板A.xlsx',
      quotationColumns: ['项目名称', '单位', '数量', '不含税单价', '不含税总价'],
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationDataChange: vi.fn(),
    }
    render(<QuotationWorkspace {...mockProps} />)
    expect(screen.getByText('quotation-1.json')).toBeInTheDocument()
    expect(screen.getByDisplayValue('阀门')).toBeInTheDocument()
    expect(screen.getByDisplayValue('1000')).toBeInTheDocument()
    // 未配置在 quotationColumns 中的列表头不应出现
    expect(screen.queryByText('税率')).not.toBeInTheDocument()
  })

  it('5. 点击新增数据行按钮调用 onAddRow', () => {
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [],
      projectRateCard: null,
      projectTemplate: null,
      quotationColumns: ['项目名称', '数量'],
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationDataChange: vi.fn(),
    }
    render(<QuotationWorkspace {...mockProps} />)
    const addBtn = screen.getByText('新增数据行')
    fireEvent.click(addBtn)
    expect(mockProps.onAddRow).toHaveBeenCalled()
  })
})
