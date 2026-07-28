import { useState, useRef } from 'react';
import { ChevronRight, ChevronDown, Plus, Trash2, Folder, FolderOpen, FileText, Upload, ScanLine, Receipt, Settings } from 'lucide-react';
import { Input } from '@/components/ui/input';

interface SidebarSectionProps {
  title: string;
  icon?: React.ReactNode;
  expanded: boolean;
  onToggle: () => void;
  onCreate?: (value: string) => Promise<void>;
  createPlaceholder?: string;
  children: React.ReactNode;
}

export function SidebarSection({ title, icon, expanded, onToggle, onCreate, createPlaceholder, children }: SidebarSectionProps) {
  const [creating, setCreating] = useState(false);
  const [newValue, setNewValue] = useState('');
  const [submitting, setSubmitting] = useState(false);

  const handleSubmit = async () => {
    const trimmed = newValue.trim();
    if (!trimmed || submitting || !onCreate) return;
    setSubmitting(true);
    try {
      await onCreate(trimmed);
      setNewValue('');
      setCreating(false);
    } finally {
      setSubmitting(false);
    }
  };

  const handleCancel = () => {
    setCreating(false);
    setNewValue('');
  };

  return (
    <div className="mb-4">
      <div className="flex items-center gap-2 mb-2">
        <button
          onClick={onToggle}
          className="flex items-center gap-2 flex-1 text-left text-sm font-semibold opacity-70 hover:opacity-100 transition-opacity"
        >
          {expanded ? <ChevronDown size={14} /> : <ChevronRight size={14} />}
          {icon}
          {title}
        </button>
        {onCreate && (
          <button
            onClick={() => setCreating(!creating)}
            className="text-slate-400 hover:text-white transition-colors flex-shrink-0"
            title="新建"
          >
            <Plus size={14} />
          </button>
        )}
      </div>
      {expanded && (
        <div className="ml-2">
          {creating && (
            <div className="flex gap-1 mb-2">
              <Input
                autoFocus
                value={newValue}
                onChange={(e) => setNewValue(e.target.value)}
                placeholder={createPlaceholder || '输入名称...'}
                className="bg-white text-gray-800 placeholder-gray-400 text-sm h-8"
                onKeyDown={(e) => {
                  if (e.key === 'Enter') handleSubmit();
                  if (e.key === 'Escape') handleCancel();
                }}
                onBlur={handleCancel}
              />
            </div>
          )}
          {children}
        </div>
      )}
    </div>
  );
}

interface TreeItemProps {
  label: string;
  active?: boolean;
  onClick?: () => void;
  onDelete?: () => void;
  expandable?: boolean;
  expanded?: boolean;
  onToggleExpand?: () => void;
  onUpload?: (files: FileList) => void;
  onCreate?: () => void;
  onSettings?: () => void;
  children?: React.ReactNode;
  icon?: React.ReactNode;
}

export function TreeItem({
  label,
  active,
  onClick,
  onDelete,
  expandable,
  expanded,
  onToggleExpand,
  onUpload,
  onCreate,
  onSettings,
  children,
  icon,
}: TreeItemProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  return (
    <div className="mb-1">
      <div
        className={`group flex items-center gap-2 p-2 cursor-pointer rounded transition text-sm ${
          active ? 'bg-slate-700 text-white' : 'text-slate-300 hover:bg-slate-700/50 hover:text-white'
        }`}
        onClick={expandable ? onToggleExpand : onClick}
      >
        {expandable && (
          <span className="flex-shrink-0">
            {expanded ? <ChevronDown size={12} /> : <ChevronRight size={12} />}
          </span>
        )}
        {icon && <span className="flex-shrink-0">{icon}</span>}
        <span className="truncate flex-1">{label}</span>
        {onUpload && (
          <>
            <input
              ref={fileInputRef}
              type="file"
              multiple
              className="hidden"
              onChange={(e) => {
                if (e.target.files && e.target.files.length > 0) {
                  onUpload(e.target.files);
                }
                e.target.value = '';
              }}
            />
            <span
              className="text-slate-400 opacity-0 group-hover:opacity-100 hover:text-white transition-opacity select-none flex-shrink-0"
              title="上传图片"
              onClick={(e) => {
                e.stopPropagation();
                fileInputRef.current?.click();
              }}
            >
              <Upload size={14} />
            </span>
          </>
        )}
        {onCreate && (
          <span
            className="text-slate-400 opacity-0 group-hover:opacity-100 hover:text-white transition-opacity select-none flex-shrink-0"
            title="新建"
            onClick={(e) => {
              e.stopPropagation();
              onCreate();
            }}
          >
            <Plus size={14} />
          </span>
        )}
        {onDelete && (
          <span
            className="text-red-400 opacity-0 group-hover:opacity-100 hover:text-red-300 transition-opacity select-none flex-shrink-0"
            title="删除"
            onClick={(e) => {
              e.stopPropagation();
              onDelete();
            }}
          >
            <Trash2 size={14} />
          </span>
        )}
        {onSettings && (
          <span
            className="text-slate-400 opacity-0 group-hover:opacity-100 hover:text-white transition-opacity select-none flex-shrink-0"
            title="项目设置"
            onClick={(e) => {
              e.stopPropagation();
              onSettings();
            }}
          >
            <Settings size={14} />
          </span>
        )}
      </div>
      {expandable && expanded && children && <div className="ml-4">{children}</div>}
    </div>
  );
}

export { Folder, FolderOpen, FileText, ScanLine, Receipt, Settings };
