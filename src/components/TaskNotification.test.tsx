import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import { TaskNotification } from './TaskNotification'

describe('TaskNotification', () => {
  it('renders nothing when status is null', () => {
    const { container } = render(
      <TaskNotification status={null} onDismiss={vi.fn()} />
    )
    expect(container.innerHTML).toBe('')
  })

  it('renders processing status', () => {
    render(
      <TaskNotification
        status={{ status: 'processing' }}
        labels={{ processing: '处理中...' }}
        onDismiss={vi.fn()}
      />
    )
    expect(screen.getByText('处理中...')).toBeInTheDocument()
  })

  it('renders done status', () => {
    render(
      <TaskNotification
        status={{ status: 'done' }}
        labels={{ done: '完成' }}
        onDismiss={vi.fn()}
      />
    )
    expect(screen.getByText('完成')).toBeInTheDocument()
  })

  it('renders error status with message', () => {
    render(
      <TaskNotification
        status={{ status: 'error', message: '出错了' }}
        labels={{ error: '失败' }}
        onDismiss={vi.fn()}
      />
    )
    expect(screen.getByText(/失败/)).toBeInTheDocument()
    expect(screen.getByText(/出错了/)).toBeInTheDocument()
  })

  it('renders with default labels', () => {
    render(
      <TaskNotification
        status={{ status: 'processing' }}
        onDismiss={vi.fn()}
      />
    )
    expect(screen.getByText('处理中...')).toBeInTheDocument()
  })
})
