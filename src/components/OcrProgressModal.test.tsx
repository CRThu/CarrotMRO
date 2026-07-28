import { render, screen, fireEvent } from '@testing-library/react'
import { describe, it, expect, vi } from 'vitest'
import { OcrProgressModal } from './OcrProgressModal'

describe('OcrProgressModal Component', () => {
  it('1. 当 isOpen 为 false 时不渲染任何模态框 DOM', () => {
    const { container } = render(
      <OcrProgressModal
        isOpen={false}
        status="processing"
        imageCount={2}
        currentStep="处理中..."
        logs={[]}
        onClose={vi.fn()}
      />
    )
    expect(container.innerHTML).toBe('')
  })

  it('2. 渲染正在处理状态与流式日志内容', () => {
    const logs = ['[12:00:00] 开始任务', '[12:00:01] 已发送请求至大模型']
    render(
      <OcrProgressModal
        isOpen={true}
        status="processing"
        imageCount={3}
        currentStep="大模型推理中..."
        logs={logs}
        onClose={vi.fn()}
      />
    )
    expect(screen.getByText(/AI 多图智能 OCR 提取中 \(共 3 张图片\)/)).toBeInTheDocument()
    expect(screen.getByText('大模型推理中...')).toBeInTheDocument()
    expect(screen.getByText('[12:00:00] 开始任务')).toBeInTheDocument()
    expect(screen.getByText('[12:00:01] 已发送请求至大模型')).toBeInTheDocument()
  })

  it('3. 渲染错误提示状态与关闭事件响应', () => {
    const onCloseMock = vi.fn()
    render(
      <OcrProgressModal
        isOpen={true}
        status="error"
        imageCount={1}
        currentStep="网络失败"
        logs={['[12:00:00] 400 Bad Request']}
        errorMessage="API Key 无效或请求超限"
        onClose={onCloseMock}
      />
    )
    expect(screen.getByText('识别中断或产生错误')).toBeInTheDocument()
    expect(screen.getByText('API Key 无效或请求超限')).toBeInTheDocument()

    const closeBtn = screen.getByText('关闭窗口')
    fireEvent.click(closeBtn)
    expect(onCloseMock).toHaveBeenCalled()
  })

  it('4. 【回归】done 状态下正确显示提取行数，不显示 0（OCR 0 行 Bug 回归）', () => {
    // 模拟修复后：task.result.items 是数组，itemCount 应为真实行数
    render(
      <OcrProgressModal
        isOpen={true}
        status="done"
        imageCount={2}
        currentStep="识别成功"
        logs={['[12:00:00] 识别成功！合并提取 24 行表格数据。']}
        itemCount={24}  // 正确行数，不应显示 0
        onClose={vi.fn()}
      />
    )
    expect(screen.getByText('识别完成！表格数据提取成功')).toBeInTheDocument()
    // 提取行数必须显示正确值 24，而不是 0
    expect(screen.getByText(/共提取/)).toHaveTextContent('24')
    expect(screen.queryByText(/共提取 0 行/)).not.toBeInTheDocument()
  })

  it('5. done 状态下 itemCount 为 0 时正确显示 0（合法边界，不崩溃）', () => {
    render(
      <OcrProgressModal
        isOpen={true}
        status="done"
        imageCount={1}
        currentStep="识别完成"
        logs={[]}
        itemCount={0}
        onClose={vi.fn()}
      />
    )
    // 应仍然展示成功状态，但行数为 0
    expect(screen.getByText('识别完成！表格数据提取成功')).toBeInTheDocument()
    expect(screen.getByText(/共提取/)).toHaveTextContent('0')
  })
})

