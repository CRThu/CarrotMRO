import { describe, it, expect, vi } from 'vitest'
import { render, screen } from '@testing-library/react'
import { DataTable } from './DataTable'
import type { RateCardColumn, TableItem } from '@/types'

// Mock UI components
vi.mock('@/components/ui/table', () => ({
  Table: ({ children }: { children: React.ReactNode }) => <table data-testid="table">{children}</table>,
  TableBody: ({ children }: { children: React.ReactNode }) => <tbody>{children}</tbody>,
  TableCell: ({ children }: { children: React.ReactNode }) => <td>{children}</td>,
  TableHead: ({ children }: { children: React.ReactNode }) => <th>{children}</th>,
  TableHeader: ({ children }: { children: React.ReactNode }) => <thead>{children}</thead>,
  TableRow: ({ children }: { children: React.ReactNode }) => <tr>{children}</tr>,
}))

vi.mock('@/components/ui/input', () => ({
  Input: ({ value, onChange, ...props }: any) => (
    <input data-testid="input" value={value} onChange={onChange} {...props} />
  ),
}))

vi.mock('@/components/ui/button', () => ({
  Button: ({ children, onClick, ...props }: any) => <button onClick={onClick} {...props}>{children}</button>,
}))

describe('DataTable', () => {
  const columns: RateCardColumn[] = [
    { name: '项目名称', strict: true, alias: null },
    { name: '数量', strict: true, alias: null },
  ]

  const items: TableItem[] = [
    { '项目名称': '螺栓M10', '数量': '100' },
    { '项目名称': '垫片M10', '数量': '200' },
  ]

  it('renders table with correct headers', () => {
    render(<DataTable columns={columns} items={items} />)
    expect(screen.getByText('项目名称')).toBeInTheDocument()
    expect(screen.getByText('数量')).toBeInTheDocument()
  })

  it('renders correct number of rows', () => {
    render(<DataTable columns={columns} items={items} />)
    expect(screen.getByText('螺栓M10')).toBeInTheDocument()
    expect(screen.getByText('垫片M10')).toBeInTheDocument()
  })

  it('shows add row button when onAddRow provided', () => {
    const onAddRow = vi.fn()
    render(<DataTable columns={columns} items={items} onAddRow={onAddRow} />)
    expect(screen.getByText('新增数据行')).toBeInTheDocument()
  })

  it('does not show add row button when onAddRow not provided', () => {
    render(<DataTable columns={columns} items={items} />)
    expect(screen.queryByText('新增数据行')).not.toBeInTheDocument()
  })

  it('renders empty table when items is empty', () => {
    render(<DataTable columns={columns} items={[]} />)
    expect(screen.getByText('项目名称')).toBeInTheDocument()
    expect(screen.getByText('数量')).toBeInTheDocument()
  })

  it('calls onAddRow when add button clicked', () => {
    const onAddRow = vi.fn()
    render(<DataTable columns={columns} items={items} onAddRow={onAddRow} />)
    screen.getByText('新增数据行').click()
    expect(onAddRow).toHaveBeenCalled()
  })

  it('renders computed & derived columns (含税单价, 不含税总价, 含税总价) with readOnly attribute', () => {
    const computedColumns = ['项目名称', '数量', '不含税单价', '含税单价', '不含税总价', '含税总价']
    const testItems = [{ '项目名称': '角钢', '数量': '10', '不含税单价': '50.00', '含税单价': '56.50', '不含税总价': '500.00', '含税总价': '565.00' }]
    render(<DataTable columns={computedColumns} items={testItems} onEdit={vi.fn()} />)
    
    const incPriceInput = screen.getByDisplayValue('56.50') as HTMLInputElement
    const exTotalInput = screen.getByDisplayValue('500.00') as HTMLInputElement
    const incTotalInput = screen.getByDisplayValue('565.00') as HTMLInputElement

    expect(incPriceInput.readOnly).toBe(true)
    expect(exTotalInput.readOnly).toBe(true)
    expect(incTotalInput.readOnly).toBe(true)
  })
})
