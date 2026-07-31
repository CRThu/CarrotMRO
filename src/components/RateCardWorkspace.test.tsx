import { render, screen, fireEvent, waitFor } from '@testing-library/react'
import { describe, it, expect, vi, beforeEach } from 'vitest'
import { RateCardWorkspace } from './RateCardWorkspace'
import * as api from '../api'

vi.mock('../api', () => ({
  previewRateCardImport: vi.fn(),
  importRateCardFile: vi.fn(),
}))

describe('RateCardWorkspace 协议定价表导入端到端 UI 流程测试', () => {
  const defaultProps = {
    currentRateCard: '2026协议定价表.json',
    ratecardTableData: {
      columns: ['项目组', '项目名称', '单位', '不含税单价', '说明'],
      items: [
        { '项目组': '电缆类', '项目名称': '铜芯电缆', '单位': '米', '不含税单价': '45.00', '说明': '国标' }
      ]
    },
    onRefreshData: vi.fn(),
  }

  beforeEach(() => {
    vi.clearAllMocks()
    vi.stubGlobal('alert', vi.fn())
  })

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

  it('3. 端到端完整流程：上传文件 -> 解析映射弹窗 -> 精确选择标准列 -> 提交导入 -> 刷新数据', async () => {
    // 模拟后端返回解析好的 Excel 数据 (包含全等匹配列和自定义列)
    vi.mocked(api.previewRateCardImport).mockResolvedValue({
      data: {
        headers: ['项目组', '物料名称', '计量单位', '协议单价'],
        sampleRows: [
          { '项目组': '电缆类', '物料名称': '铝芯电缆', '计量单位': '米', '协议单价': '25.00' }
        ],
        allRows: [
          { '项目组': '电缆类', '物料名称': '铝芯电缆', '计量单位': '米', '协议单价': '25.00' }
        ]
      }
    } as any)

    vi.mocked(api.importRateCardFile).mockResolvedValue({
      data: { success: true, count: 1, columns: ['项目组', '项目名称', '单位', '不含税单价'] }
    } as any)

    render(<RateCardWorkspace {...defaultProps} />)

    // 找到隐藏的 file input 并触发文件上传
    const fileInput = screen.getByTestId('file-input') as HTMLInputElement
    expect(fileInput).not.toBeNull()

    const dummyFile = new File(['fake excel content'], 'test_ratecard.xlsx', { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    fireEvent.change(fileInput!, { target: { files: [dummyFile] } })

    // 验证调用了 API previewRateCardImport
    await waitFor(() => {
      expect(api.previewRateCardImport).toHaveBeenCalledWith('2026协议定价表.json', expect.any(FormData))
    })

    // 验证弹出列映射配置对话框
    expect(await screen.findByText('导入定价表 - 列映射配置')).toBeInTheDocument()
    expect(screen.getAllByText('项目组').length).toBeGreaterThan(0)
    expect(screen.getByText('物料名称')).toBeInTheDocument()

    // 模拟用户在下拉列表中将 '物料名称' 显式映射选为内置标准列 '项目名称'
    const selects = screen.getAllByRole('combobox')
    expect(selects.length).toBe(4)

    // '项目组' 为 100% 精确全等匹配，下拉框自动选定为 '项目组'
    expect((selects[0] as HTMLSelectElement).value).toBe('项目组')
    // '物料名称' 未精确全等匹配，初始应为 ''
    expect((selects[1] as HTMLSelectElement).value).toBe('')

    // 用户在界面手动下拉将 '物料名称' 选为 '项目名称'，将 '协议单价' 选为 '不含税单价'
    fireEvent.change(selects[1], { target: { value: '项目名称' } })
    fireEvent.change(selects[3], { target: { value: '不含税单价' } })

    // 点击确认映射导入按钮
    const confirmBtn = screen.getByText('确认映射导入')
    fireEvent.click(confirmBtn)

    // 验证调用了 importRateCardFile 提交选中的标准列映射
    await waitFor(() => {
      expect(api.importRateCardFile).toHaveBeenCalledWith('2026协议定价表.json', {
        headers: ['项目组', '物料名称', '计量单位', '协议单价'],
        items: [
          { '项目组': '电缆类', '物料名称': '铝芯电缆', '计量单位': '米', '协议单价': '25.00' }
        ],
        mapping: {
          '项目组': '项目组',
          '物料名称': '项目名称',
          '计量单位': '',
          '协议单价': '不含税单价'
        }
      })
    })

    // 验证成功后触发了数据刷新 onRefreshData
    expect(defaultProps.onRefreshData).toHaveBeenCalled()
  })
})
