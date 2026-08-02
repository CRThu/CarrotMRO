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

  it('3.1 calculateRowFormulas：数量或单价清空时，自动重置总价为空字符串', () => {
    const itemWithTotal = { '数量': '', '不含税单价': '200', '不含税总价': '1000.00' }
    const calculated = calculateRowFormulas(itemWithTotal)
    expect(calculated['不含税总价']).toBe('')
  })

  it('3.2 calculateRowFormulas：【有列/无列】限制，当 activeColumns 未勾选总价列时过滤不生成该字段', () => {
    const raw = { '数量': '10', '不含税单价': '50', '不含税总价': '500' }
    // activeColumns 只选择了 ['项目名称', '单位', '数量', '不含税单价'] (没有选择 '不含税总价')
    const activeColumns = ['项目名称', '单位', '数量', '不含税单价']
    const calculated = calculateRowFormulas(raw, undefined, activeColumns)
    // 证明：未在 activeColumns 中的不含税总价被过滤清除
    expect(calculated['不含税总价']).toBeUndefined()
  })

  it('3.3 calculateRowFormulas：以不含税单价为单一主数据源，结合 0% 税率推导含税单价与含税总价', () => {
    // 0% 税率测试：以不含税单价为主源推导含税单价
    const zeroTax = { '数量': '10', '不含税单价': '100', '税率': '0%' }
    const resZero = calculateRowFormulas(zeroTax)
    expect(resZero['不含税总价']).toBe('1000.00')
    expect(resZero['含税单价']).toBe('100.00')
    expect(resZero['含税总价']).toBe('1000.00')
  })

  it('3.4 calculateRowFormulas：单一主数据源原则，以不含税单价 + 税率 自动正推含税单价与含税总价', () => {
    const raw = { '数量': '10', '不含税单价': '200.00', '税率': '13%' }
    const calculated = calculateRowFormulas(raw)
    // 统一以 200.00 * 1.13 正推出 含税单价 = 226.00，含税总价 = 2260.00
    expect(calculated['不含税总价']).toBe('2000.00')
    expect(calculated['含税单价']).toBe('226.00')
    expect(calculated['含税总价']).toBe('2260.00')
  })

  it('3.5 calculateRowFormulas：纯含税模式降级保底（无有效税率但有含税单价时，按含税单价 × 数量计算含税总价）', () => {
    const pureInclItem = { '数量': '5', '含税单价': '50.00' }
    const calculated = calculateRowFormulas(pureInclItem)
    expect(calculated['含税总价']).toBe('250.00')
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

  it('6. 渲染实时自动保存状态，且修改单元格内容能够触发 onQuotationDataChange', () => {
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [{ '项目名称': '旧物料', '数量': '1' }],
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
    expect(screen.getByText('已自动保存')).toBeInTheDocument()

    const input = screen.getByDisplayValue('旧物料')
    fireEvent.change(input, { target: { value: '新物料' } })

    expect(mockProps.onQuotationDataChange).toHaveBeenCalledWith([
      expect.objectContaining({ '项目名称': '新物料' }),
    ])
  })

  it('7. 点击【一键校验】按钮能正确比对全表条目（包含未匹配行警示与已匹配物料误改动列）并写入复核备注', () => {
    const onRemarksChange = vi.fn()
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [
        {
          _matchStatus: 'matched' as const,
          '项目名称': '铜芯电缆',
          '单位': '个', // 被用户误改为了 个
          _matchedRateCardItem: { '项目名称': '铜芯电缆', '单位': '米' }, // 原定价单为 米
        },
        {
          _matchStatus: 'pending' as const,
          '项目名称': '地面打磨',
          '单位': '平方米',
        },
      ],
      quotationRemarks: [],
      projectRateCard: '协议表A.json',
      projectTemplate: '模板A.xlsx',
      quotationColumns: ['项目名称', '单位'],
      matchValidationRules: {
        strict_name_match: true,
        check_columns: ['项目名称', '单位'],
        fill_columns: ['单位'],
      },
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationDataChange: vi.fn(),
      onQuotationRemarksChange: onRemarksChange,
    }

    render(<QuotationWorkspace {...mockProps} />)

    const validateBtn = screen.getByRole('button', { name: /一键校验/ })
    expect(validateBtn).toBeInTheDocument()

    fireEvent.click(validateBtn)

    expect(onRemarksChange).toHaveBeenCalledWith(
      expect.arrayContaining([
        expect.stringContaining('单位误修改：当前为 "个"，原定价单为 "米"'),
        expect.stringContaining('第 2 行 [地面打磨] ⚠️ 未匹配：尚未关联/匹配协议定价单物料'),
      ])
    )
  })

  it('8. 匹配弹窗支持搜索框交互与自定义关键字检索', () => {
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [{ '项目名称': '拆除地毯', '单位': '平方米' }],
      projectRateCard: '协议表A.json',
      projectTemplate: null,
      quotationColumns: ['项目名称', '单位'],
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationDataChange: vi.fn(),
    }

    render(<QuotationWorkspace {...mockProps} />)

    // 点击匹配按钮触发弹窗
    const matchBtn = screen.getByTitle('点击打开物料匹配与搜索')
    fireEvent.click(matchBtn)

    // 确认弹窗渲染搜索框且默认显示 '拆除地毯'
    const searchInput = screen.getByPlaceholderText('输入关键字检索协议定价库物料...') as HTMLInputElement
    expect(searchInput).toBeInTheDocument()
    expect(searchInput.value).toBe('拆除地毯')

    // 修改搜索框内容为 '地毯' 并回车触发搜索
    fireEvent.change(searchInput, { target: { value: '地毯' } })
    expect(searchInput.value).toBe('地毯')
    fireEvent.keyDown(searchInput, { key: 'Enter', code: 'Enter' })
  })

  it('9. 打开工程/报价单文件时，自动扫描并触发公式计算补全缺失的总价', () => {
    const onQuotationDataChangeMock = vi.fn()
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [
        // 模拟打开旧报价单文件：只有数量与单价，缺失不含税总价
        { '项目名称': '铜芯电缆', '数量': '10', '不含税单价': '45.00' }
      ],
      projectRateCard: null,
      projectTemplate: null,
      quotationColumns: ['项目名称', '数量', '不含税单价', '不含税总价'],
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationDataChange: onQuotationDataChangeMock,
    }

    render(<QuotationWorkspace {...mockProps} />)

    // 验证初始化扫描触发了 onQuotationDataChange，且不含税总价被补全计算为 450.00
    expect(onQuotationDataChangeMock).toHaveBeenCalled()
    const updatedItems = onQuotationDataChangeMock.mock.calls[0][0]
    expect(updatedItems[0]['不含税总价']).toBe('450.00')
  })

  it('10. 物料匹配带入单价时，带入行自动触发公式计算更新总价', () => {
    // 测试 calculateRowFormulas 在物料带入时的行为
    const rawMatchItem = {
      '项目名称': '铜芯电缆',
      '数量': '20',
      '不含税单价': '50.00', // 协议单价带入
    }
    const finalItem = calculateRowFormulas(rawMatchItem)
    // 带入行自动算出不含税总价 20 * 50.00 = 1000.00
    expect(finalItem['不含税总价']).toBe('1000.00')
  })

  it('11. 点击一键校验时，检测算术计算偏差并给出错误警告提醒', () => {
    const onQuotationRemarksChangeMock = vi.fn()
    const mockProps = {
      currentProject: '项目A',
      activeQuotationFilename: 'quotation-1.json',
      quotationItems: [
        // 数量 5 * 单价 10.00 = 50.00，但错填了 999.00
        { '项目名称': 'PVC管', '数量': '5', '不含税单价': '10.00', '不含税总价': '999.00', _matchStatus: 'pending' as const }
      ],
      projectRateCard: '协议表A.json',
      matchValidationRules: { check_columns: ['单位'] },
      quotationColumns: ['项目名称', '数量', '不含税单价', '不含税总价'],
      onEdit: vi.fn(),
      onAddRow: vi.fn(),
      onDeleteRow: vi.fn(),
      onSave: vi.fn(),
      onQuotationRemarksChange: onQuotationRemarksChangeMock,
    }

    render(<QuotationWorkspace {...mockProps} />)
    const validateBtn = screen.getByText('一键校验')
    fireEvent.click(validateBtn)

    // 验证一键校验成功捕获算术计算偏差并写入提示备注
    expect(onQuotationRemarksChangeMock).toHaveBeenCalled()
    const remarks = onQuotationRemarksChangeMock.mock.calls[0][0]
    expect(remarks.some((r: string) => r.includes('不含税总价计算偏差'))).toBe(true)
  })
})
