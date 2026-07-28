import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { ProjectConfigWorkspace } from './ProjectConfigWorkspace'
import { ProjectSettings, PRESET_COLUMNS } from '@/types'

describe('ProjectConfigWorkspace Component', () => {
  const mockSettings: ProjectSettings = {
    name: '测试项目A',
    created_at: '2026-07-28T12:00:00.000Z',
    ratecard_name: '协议表1.json',
    template_name: '标准模板.xlsx',
    ocr_columns: ['项目名称', '数量', '不含税单价'],
    quotation_columns: ['项目组', '项目名称', '单位', '数量', '不含税单价', '不含税总价', '含税单价', '含税总价'],
  }

  const defaultProps = {
    currentProject: '测试项目A',
    settings: mockSettings,
    rateCards: ['协议表1.json', '协议表2.json'],
    templates: ['标准模板.xlsx', '模板2.xlsx'],
    onUpdateSettings: vi.fn(),
    onUploadTemplate: vi.fn(),
    onDeleteTemplate: vi.fn(),
  }

  it('1. 渲染项目设置并正确展现已关稳定价单与模板', () => {
    render(<ProjectConfigWorkspace {...defaultProps} />)
    expect(screen.getByText('项目设置: 测试项目A')).toBeInTheDocument()
    expect(screen.getByText('关联协议定价表')).toBeInTheDocument()
    expect(screen.getByText('关联报价单 Excel 模板')).toBeInTheDocument()
    expect(screen.getByText('OCR 图像识别提取列规范')).toBeInTheDocument()
    expect(screen.getByText('报价单所需展示与编辑列规范')).toBeInTheDocument()
  })

  it('2. 点击 OCR 提取列复选框可以触发 onUpdateSettings 更新列', () => {
    render(<ProjectConfigWorkspace {...defaultProps} />)
    // 寻找 "单位" 提取列按钮
    const unitButtons = screen.getAllByRole('button', { name: /单位/ })
    expect(unitButtons.length).toBeGreaterThan(0)

    // 点击该列触发勾选
    fireEvent.click(unitButtons[0])
    expect(defaultProps.onUpdateSettings).toHaveBeenCalledWith({
      ocr_columns: ['项目名称', '数量', '不含税单价', '单位'],
    })
  })

  it('3. 呈现系统全局 10 项标准预制列选项', () => {
    render(<ProjectConfigWorkspace {...defaultProps} />)
    PRESET_COLUMNS.forEach(col => {
      expect(screen.getAllByText(col).length).toBeGreaterThan(0)
    })
  })
})
