import React, { useState, useMemo, useEffect, Component, ErrorInfo } from 'react';
import { useNavigate, Routes, Route, useLocation, useParams, Navigate } from 'react-router-dom';
import * as XLSX from 'xlsx';

const handleExcelExport = (data: any[], fileName: string) => {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
};

import { HOLIDAYS_2026, CUTI_BERSAMA_2026 } from './constants';
import { 
  Plus, 
  History,
  LayoutGrid,
  AlertTriangle,
  Cpu,
  LayoutDashboard,
  Users,
  ShieldAlert,
  FolderKanban,
  RefreshCw,
  AlertCircle,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  LogOut,
  Settings as SettingsIcon,
  Search,
  Trash2,
  ShieldCheck,
  ExternalLink,
  Edit3,
  Save,
  UserPlus,
  X,
  ArrowRight,
  User as UserIcon,
  Clock,
  ArrowLeft,
  ArrowDown,
  Rocket,
  Activity,
  Calendar,
  Layers,
  PlusSquare,
  Upload,
  Download,
  Filter,
  Copy
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  format, 
  addHours, 
  startOfDay, 
  endOfDay, 
  eachHourOfInterval, 
  isSameHour,
  differenceInHours,
  addDays,
  addWeeks,
  startOfWeek,
  endOfWeek,
  eachDayOfInterval,
  eachMonthOfInterval,
  isSameDay,
  differenceInDays,
  startOfMonth,
  endOfMonth,
  eachWeekOfInterval,
  isWithinInterval,
  addMonths,
  subDays,
  isToday,
  isWeekend
} from 'date-fns';
import { calculateTimelineRange, calculateBarCoordinates, sanitizeDate, normalizeDate, calculateParentRange } from './utils/timelineEngine';
import { Task, ViewScale, TaskStatus, ProjectStatus, Project, AppUser, AppView, AuditLog, Schedule, ProjectRescheduleLog, RescheduleRequest } from './types';
import { cn } from './lib/utils';
import { getSafeKey } from './utils/keyHelper';

/**
 * ROBUST PERSISTENT ID GENERATOR
 * ONLY call this when creating NEW entities (Phase, Task, Project).
 * NEVER call this inside a render function or .map() logic.
 */
export const createPersistentId = () => {
    try {
      if (typeof crypto !== 'undefined' && crypto.randomUUID) {
        return crypto.randomUUID();
      }
      return `pers-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    } catch (e) {
      return `pers-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    }
};

import { useTasks } from './hooks/useTasks';
import { taskService } from './services/taskService';
import { supabase } from './lib/supabase';

import { useAuth } from './context/AuthContext';
import { LoginPage } from './components/LoginPage';

const formatWorkday = (totalHours: number) => {
  if (!totalHours || totalHours <= 0) return '-';
  const days = Math.floor(totalHours / 9);
  const remainingHours = Math.floor(totalHours % 9);
  const remainingMinutes = Math.round((totalHours % 1) * 60);
  
  const parts = [];
  if (days > 0) parts.push(`${days} Hari Kerja`);
  if (remainingHours > 0) parts.push(`${remainingHours} Jam`);
  if (remainingMinutes > 0) parts.push(`${remainingMinutes} Menit`);
  
  return parts.length > 0 ? parts.join(' ') : '-';
};

// Utils moved to ganttDateUtils.ts
const formatDateForDB = (dateInput: any) => {
    return sanitizeDate(dateInput);
};

// Use getSafeKey from utils/keyHelper.ts instead.

// --- Collision Engine ---
const getCollision = (currentTask: Task, allTasks: Task[], projects: Project[]) => {
  const overlap = allTasks.filter(task => {
    if (task.id === currentTask.id) return false;
    
    // Ignore parent-child relationship as they naturally overlap in time
    if (task.id === currentTask.parent_id || currentTask.id === task.parent_id) return false;
    
    if (!currentTask.start_time || !task.start_time) return false;
    
    const start1 = new Date(currentTask.start_time).getTime();
    const end1 = new Date(currentTask.end_time).getTime();
    const start2 = new Date(task.start_time).getTime();
    const end2 = new Date(task.end_time).getTime();

    const timeOverlap = start1 < end2 && end1 > start2;
    if (!timeOverlap) return false;

    const samePIC = currentTask.assignee && task.assignee && currentTask.assignee === task.assignee;
    const sameDev = currentTask.developer_name && task.developer_name && currentTask.developer_name === task.developer_name;
    const sameQA = currentTask.qa_name && task.qa_name && currentTask.qa_name === task.qa_name;

    return samePIC || sameDev || sameQA;
  });

  if (overlap.length === 0) return null;

  return overlap.map(collisionTask => {
    const project = projects.find(p => p.id === collisionTask.project_id);
    return {
      ...collisionTask,
      projectName: project?.name || 'Unknown Project'
    };
  });
};

const isWBSConflict = (parent: any, child: any) => {
  if (!parent || !child) return false;
  if (!parent.start_time || !parent.end_time || !child.start_time || !child.end_time) return false;
  
  const pStart = new Date(parent.start_time).getTime();
  const pEnd = new Date(parent.end_time).getTime();
  const cStart = new Date(child.start_time).getTime();
  const cEnd = new Date(child.end_time).getTime();

  // It is ONLY a conflict if child starts before parent OR ends after parent
  return (cStart < pStart || cEnd > pEnd);
};

// --- Components ---

// --- Error Boundary ---
interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
}

class ErrorBoundary extends Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(_: Error): ErrorBoundaryState {
    return { hasError: true };
  }

  componentDidCatch(error: Error, errorInfo: ErrorInfo) {
    console.error("ErrorBoundary caught an error:", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div className="fixed inset-0 z-[100] flex items-center justify-center bg-slate-950/80 backdrop-blur-sm p-4">
          <div className="bg-slate-900 border border-slate-800 p-8 rounded-2xl max-w-md w-full shadow-2xl text-center space-y-6">
            <div className="w-16 h-16 bg-rose-500/10 rounded-full flex items-center justify-center mx-auto">
              <AlertTriangle className="w-8 h-8 text-rose-500" />
            </div>
            <div className="space-y-2">
              <h2 className="text-xl font-bold text-white tracking-tight">Wizard Critical Failure</h2>
              <p className="text-slate-400 text-sm leading-relaxed">
                The interface encountered a runtime error or missing asset. The component has been isolated for safety.
              </p>
            </div>
            <button 
              onClick={() => window.location.reload()}
              className="w-full py-3 bg-slate-800 hover:bg-slate-700 text-white font-bold uppercase text-xs tracking-widest rounded-xl transition-all"
            >
              Relational Reset (Reload)
            </button>
          </div>
        </div>
      );
    }

    return this.props.children; 
  }
}

const StatusBadge = ({ status, type = 'task' }: { status: string, type?: 'project' | 'task' }) => {
  const styles: Record<string, string> = {
    // Task Status
    'On Hold': 'bg-slate-500/10 text-slate-400 border-slate-500/20',
    'On Progress': 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20',
    'Done': 'bg-emerald-500/10 text-emerald-400 border-emerald-500/20',
    // Project Status
    'FSD on Progress': 'bg-blue-500/10 text-blue-400 border-blue-500/20',
    'FSD on Review': 'bg-purple-500/10 text-purple-400 border-purple-500/20',
    'SIT on Progress': 'bg-amber-500/10 text-amber-400 border-amber-500/20',
    'UAT on Progress': 'bg-indigo-500/10 text-indigo-400 border-indigo-500/20',
    'Project Late': 'bg-rose-500/10 text-rose-400 border-rose-500/20',
  };

  return (
    <span className={cn(
      "px-2 py-0.5 rounded-full text-[9px] font-black border uppercase tracking-wider whitespace-nowrap",
      styles[status] || 'bg-slate-500/10 text-slate-400 border-slate-500/20'
    )}>
      {status || 'Unknown'}
    </span>
  );
};

const TaskStatusSelector = ({ status, onUpdate, disabled }: { status: any, onUpdate: (s: any) => void, disabled?: boolean }) => {
  const options = [TaskStatus.ON_HOLD, TaskStatus.ON_PROGRESS, TaskStatus.DONE];
  if (disabled) return <StatusBadge status={status} />;
  return (
    <select 
      value={status || TaskStatus.ON_PROGRESS}
      onChange={(e) => onUpdate(e.target.value)}
      disabled={disabled}
      className={cn(
        "bg-slate-900 border border-slate-800 rounded px-2 py-1 text-[9px] font-black uppercase outline-none transition-all cursor-pointer",
        status === TaskStatus.DONE ? "text-emerald-400" : status === TaskStatus.ON_HOLD ? "text-rose-400" : "text-indigo-400"
      )}
    >
      {options.map((opt, i) => <option key={getSafeKey({id: opt}, i, 'task-status-opt')} value={opt}>{opt}</option>)}
    </select>
  );
};

const ProjectStatusSelector = ({ status, onUpdate, disabled }: { status: any, onUpdate: (s: any) => void, disabled?: boolean }) => {
  const options = [
    ProjectStatus.FSD_PROGRESS, 
    ProjectStatus.FSD_REVIEW, 
    ProjectStatus.SIT_PROGRESS, 
    ProjectStatus.ON_HOLD, 
    ProjectStatus.UAT_PROGRESS
  ];
  if (disabled) return <StatusBadge status={status} />;
  return (
    <select 
      value={status || ProjectStatus.FSD_PROGRESS}
      onChange={(e) => onUpdate(e.target.value)}
      disabled={disabled}
      className="bg-slate-900 border border-slate-800 rounded px-2 py-1 text-[9px] font-black uppercase text-indigo-400 outline-none transition-all cursor-pointer"
    >
      {options.map((opt, i) => <option key={getSafeKey({id: opt}, i, 'project-status-opt')} value={opt}>{opt}</option>)}
    </select>
  );
};

const ApprovalBadge = ({ value, label, onUpdate, disabled }: { value: string | null, label: string, onUpdate: (val: string) => void, disabled?: boolean }) => {
  const options = ['Pending', 'Revise', 'OK'];
  const safeValue = options.includes(value || '') ? (value || 'Pending') : 'Pending';

  if (disabled) return <span className="text-[9px] font-black uppercase text-slate-500">{safeValue}</span>;

  return (
    <div className="flex flex-col gap-1 w-full max-w-[80px]">
      <select 
        value={safeValue}
        onChange={(e) => onUpdate(e.target.value)}
        disabled={disabled}
        className={cn(
          "bg-slate-900 border border-slate-800 rounded px-1.5 py-1 text-[9px] font-black uppercase tracking-tighter outline-none focus:ring-1 focus:ring-indigo-500/50 transition-all cursor-pointer w-full",
          safeValue === 'OK' ? "text-emerald-400 border-emerald-500/30" :
          safeValue === 'Revise' ? "text-rose-400 border-rose-500/30" :
          "text-slate-500"
        )}
      >
        {options.map((opt, i) => (
          <option 
            key={getSafeKey({id: opt}, i, 'approval-opt')} 
            value={opt} 
            className="bg-slate-900"
          >
            {opt}
          </option>
        ))}
      </select>
    </div>
  );
};


// --- Main App ---

// --- Health Engine ---

const getTaskHealth = (task: Task) => {
  if (task.status === TaskStatus.DONE) return 'Stable';
  
  const now = new Date();
  const end = new Date(task.end_time);
  
  if (now > end) return 'OVERDUE';
  
  if (task.target_sla_date) {
    const sla = new Date(task.target_sla_date);
    if (now > sla) return 'OVER SLA';
  }
  
  return 'Healthy';
};

const HealthBadge = ({ health }: { health: string }) => {
  if (health === 'Stable') return null;
  if (health === 'Healthy') return null;
  
  return (
    <span className={cn(
      "px-2 py-0.5 rounded-full text-[9px] font-black uppercase tracking-widest animate-pulse shadow-lg",
      health === 'OVERDUE' ? "bg-rose-500 text-white shadow-rose-500/20" : "bg-amber-500 text-slate-900 shadow-amber-500/20"
    )}>
      {health}
    </span>
  );
};

// --- Custom UI Overlays ---

const ConfirmModal = ({ 
  isOpen, 
  onClose, 
  onConfirm, 
  title, 
  description, 
  confirmText = 'Hapus', 
  cancelText = 'Batal',
  variant = 'danger'
}: { 
  isOpen: boolean, 
  onClose: () => void, 
  onConfirm: () => void, 
  title: string, 
  description: string,
  confirmText?: string,
  cancelText?: string,
  variant?: 'danger' | 'primary'
}) => {
  if (!isOpen) return null;

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-[150] flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm"
        />
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          className="relative w-full max-w-md bg-slate-900 border border-slate-700/50 rounded-xl shadow-2xl overflow-hidden p-6"
        >
          <div className="flex flex-col items-center text-center">
            <div className={cn(
              "w-12 h-12 rounded-full flex items-center justify-center mb-4",
              variant === 'primary' ? "bg-indigo-500/10" : "bg-red-500/10"
            )}>
              {variant === 'primary' ? (
                <ShieldCheck className="w-6 h-6 text-indigo-500" />
              ) : (
                <AlertTriangle className="w-6 h-6 text-red-500" />
              )}
            </div>
            <h3 className="text-xl font-bold text-slate-100 mb-2">{title}</h3>
            <p className="text-slate-400 mb-6 text-sm leading-relaxed">{description}</p>
          </div>
          
          <div className="flex justify-end gap-3">
            <button 
              onClick={onClose}
              className="bg-slate-800 hover:bg-slate-700 text-slate-200 px-4 py-2 rounded-lg transition-colors text-sm font-semibold"
            >
              {cancelText}
            </button>
            <button 
              onClick={() => { onConfirm(); onClose(); }}
              className={cn(
                "px-4 py-2 rounded-lg transition-colors flex items-center gap-2 text-sm font-semibold",
                variant === 'primary' 
                  ? "bg-indigo-600 hover:bg-indigo-500 text-white" 
                  : "bg-red-500/10 hover:bg-red-600 border border-red-500/20 text-red-500 hover:text-white"
              )}
            >
              {variant === 'danger' && <Trash2 className="w-4 h-4" />}
              {confirmText}
            </button>
          </div>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

const SuccessNotification = ({ show, message, onClose }: { show: boolean, message: string, onClose: () => void }) => {
  useEffect(() => {
    if (show) {
      const timer = setTimeout(onClose, 3000);
      return () => clearTimeout(timer);
    }
  }, [show, onClose]);

  return (
    <AnimatePresence>
      {show && (
        <motion.div 
          initial={{ opacity: 0, y: -50, x: '-50%' }}
          animate={{ opacity: 1, y: 0, x: '-50%' }}
          exit={{ opacity: 0, y: -50, x: '-50%' }}
          className="fixed top-8 left-1/2 z-[150] bg-emerald-500 text-white px-6 py-3 rounded-full shadow-[0_10px_40px_rgba(16,185,129,0.4)] flex items-center gap-3 border border-emerald-400"
        >
          <div className="w-5 h-5 bg-white/20 rounded-full flex items-center justify-center">
            <Plus className="w-3 h-3 rotate-45" />
          </div>
          <span className="text-xs font-black uppercase tracking-widest">{message}</span>
        </motion.div>
      )}
    </AnimatePresence>
  );
};

// --- Components ---

function EditableInput({ value, onSave, className, placeholder, type = 'text', min, max, disabled }: { 
  value: string | number, 
  onSave: (val: any) => void, 
  className?: string, 
  placeholder?: string,
  type?: 'text' | 'number',
  min?: string | number,
  max?: string | number,
  disabled?: boolean
}) {
  const [localValue, setLocalValue] = useState(value);

  useEffect(() => {
    setLocalValue(value);
  }, [value]);

  const handleBlur = () => {
    if (disabled) return;
    if (localValue !== value) {
      if (type === 'number') {
        let val = parseFloat(String(localValue)) || 0;
        if (min !== undefined) val = Math.max(Number(min), val);
        if (max !== undefined) val = Math.min(Number(max), val);
        onSave(val);
      } else {
        onSave(localValue);
      }
    }
  };

  const handleKeyDown = (e: React.KeyboardEvent) => {
    if (disabled) return;
    if (e.key === 'Enter') {
      (e.target as HTMLInputElement).blur();
    }
    if (e.key === 'Escape') {
      setLocalValue(value);
      (e.target as HTMLInputElement).blur();
    }
  };

  if (disabled) {
    return <span className={cn(className, "cursor-default opacity-80")}>{value || placeholder || '-'}</span>;
  }

  return (
    <input 
      type={type}
      min={min}
      value={localValue}
      onChange={(e) => {
        const val = e.target.value;
        if (type === 'number' && min !== undefined) {
           const n = parseInt(val) || 0;
           setLocalValue(Math.max(Number(min), n));
        } else {
           setLocalValue(val);
        }
      }}
      onBlur={handleBlur}
      onKeyDown={handleKeyDown}
      className={className}
      placeholder={placeholder}
      disabled={disabled}
    />
  );
}

const CollisionWarning = ({ collisions }: { collisions: any[] }) => {
  return (
    <div className="relative group inline-flex items-center justify-center ml-2">
      <button 
        className="text-rose-500 animate-pulse hover:scale-110 transition-transform"
      >
        <AlertTriangle className="w-4 h-4" />
      </button>
      
      <div className="absolute bottom-full mb-2 hidden group-hover:block w-max max-w-xs bg-slate-800 border border-slate-700 text-slate-100 text-xs rounded-md p-3 z-[100] shadow-xl pointer-events-none">
        <p className="text-[10px] font-bold text-rose-400 uppercase tracking-widest mb-2 flex items-center gap-1">
          <ShieldAlert className="w-3 h-3" /> External Conflict Detected
        </p>
        <div className="space-y-3">
          {(collisions || []).map((c, i) => {
            if (!c) return null;
            return (
              <div 
                key={getSafeKey(c, i, 'collision')} 
                className="text-[11px] leading-relaxed border-l-2 border-slate-800 pl-2"
              >
                <span className="text-slate-300 font-bold block mb-0.5">Project: {c.projectName || 'Unknown'}</span>
                <span className="text-slate-500 block">Task: <span className="text-indigo-400">"{c.title || 'Untitled'}"</span></span>
                <span className="text-slate-500 block">Conflict: {c.start_time ? format(new Date(c.start_time), 'MMM dd') : '??'} - {c.end_time ? format(new Date(c.end_time), 'MMM dd') : '??'}</span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

const ProfileSettingsModal = ({ 
  isOpen, 
  onClose, 
  user,
  onUpdate
}: { 
  isOpen: boolean, 
  onClose: () => void, 
  user: AppUser | null,
  onUpdate: (updated: AppUser) => void
}) => {
  const [name, setName] = useState(user?.name || '');
  const [role, setRole] = useState(user?.role || '');
  const [accessLevel, setAccessLevel] = useState(user?.access_level || 'PIC');
  const [isSaving, setIsSaving] = useState(false);

  if (!isOpen || !user) return null;

  const handleSave = async (e: React.FormEvent) => {
    e.preventDefault();
    setIsSaving(true);
    try {
      const updates = {
        name,
        role,
        access_level: accessLevel,
        updated_at: new Date().toISOString()
      };

      const { data, error } = await supabase
        .from('users')
        .update(updates)
        .eq('id', user.id)
        .select()
        .single();

      if (error) throw error;
      if (data) {
        onUpdate(data);
        onClose();
      }
    } catch (err: any) {
      console.error('Failed to update profile:', err);
      alert('Error updating profile: ' + err.message);
    } finally {
      setIsSaving(false);
    }
  };

  const isPIC = user?.access_level === 'PIC';

  return (
    <AnimatePresence>
      <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-slate-950/80 backdrop-blur-sm"
        />
        <motion.div 
          initial={{ opacity: 0, scale: 0.95, y: 10 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 10 }}
          className="relative w-full max-w-md bg-slate-900 border border-slate-800 rounded-2xl shadow-2xl overflow-hidden"
        >
          <div className="p-6 border-b border-slate-800 bg-indigo-600/5">
            <h3 className="text-xl font-bold text-white flex items-center gap-2">
              <UserIcon className="w-5 h-5 text-indigo-500" />
              Profile Configuration
            </h3>
            <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Personnel Detail Node</p>
          </div>

          <form onSubmit={handleSave} className="p-6 space-y-5">
            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Full Identity Name</label>
              <input 
                type="text"
                required
                value={name}
                onChange={(e) => setName(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors"
                placeholder="Ex: Fachrul Wisnu"
              />
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Organizational Role</label>
              <input 
                type="text"
                required
                value={role}
                onChange={(e) => setRole(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors"
                placeholder="Ex: Senior Developer"
              />
            </div>

            <div className="space-y-2">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Access Level Privilege</label>
              <select 
                value={accessLevel}
                onChange={(e) => setAccessLevel(e.target.value)}
                disabled={isPIC}
                className={cn(
                  "w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors appearance-none",
                  isPIC && "opacity-50 cursor-not-allowed"
                )}
              >
                <option value="PIC">PIC / Standard Personnel</option>
                <option value="Admin">Administrator</option>
                <option value="Superadmin">Super Administrator</option>
              </select>
              {isPIC && (
                <p className="text-[9px] text-rose-500/70 font-bold uppercase italic mt-1 ml-1 flex items-center gap-1">
                  <ShieldAlert className="w-3 h-3" /> Promotion Restricted for PIC role
                </p>
              )}
            </div>

            <div className="pt-4 flex gap-3">
              <button 
                type="button"
                onClick={onClose}
                className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 font-bold uppercase text-[10px] tracking-widest rounded-xl transition-all"
              >
                Abort
              </button>
              <button 
                type="submit"
                disabled={isSaving}
                className="flex-[2] py-3 bg-indigo-600 hover:bg-indigo-500 disabled:opacity-50 text-white font-black uppercase text-[10px] tracking-widest rounded-xl shadow-lg shadow-indigo-600/20 transition-all flex items-center justify-center gap-2"
              >
                {isSaving ? "Syncing..." : "Commit Changes"}
              </button>
            </div>
          </form>
        </motion.div>
      </div>
    </AnimatePresence>
  );
};

export default function App() {
  const { user: authUser, currentUser, loading: authLoading, signOut: realSignOut, setCurrentUser } = useAuth();
  
  // No longer use mock user as primary, use currentUser from context
  const user = useMemo(() => currentUser || (authUser ? ({
    id: authUser.id,
    name: authUser.email?.split('@')[0] || 'Unknown',
    email: authUser.email || '',
    access_level: 'PIC',
    role: 'Viewer'
  } as AppUser) : null), [authUser, currentUser]);

  const [isProfileModalOpen, setIsProfileModalOpen] = useState(false);
  const isSyncingRef = React.useRef(false);

  const handleOpenReschedule = (p: Project) => {
    setReschedulingProject(p);
  };

  const isAdmin = useMemo(() => {
    if (!user) return false;
    return user.access_level === 'Superadmin' || user.access_level === 'Admin';
  }, [user]);

  const isSuperadmin = useMemo(() => {
    if (!user) return false;
    return user.access_level === 'Superadmin' || user.email?.toLowerCase().includes('wisnu');
  }, [user]);

  const canAccessReschedule = useMemo(() => {
    if (!user) return false;
    return isAdmin || user.access_level === 'Developer' || user.access_level === 'QA';
  }, [user, isAdmin]);

  const signOut = () => {
    realSignOut();
    navigate('/login');
  };

  const navigate = useNavigate();
  const location = useLocation();
  const [selectedProjectId, setSelectedProjectId] = useState<string | null>(null);
  const [pendingRescheduleCount, setPendingRescheduleCount] = useState(0);

  const [viewState, setViewState] = useState<AppView>('PROJECTS');
  
  // Explicitly map URL to View for legacy support and cleaner routing
  const activeView = useMemo(() => {
    const path = location.pathname;
    if (path === '/') return 'PROJECTS';
    if (path === '/portfolio') return 'PROJECTS';
    if (path === '/timeline') return 'GANTT_DETAIL';
    if (path === '/kanban') return 'KANBAN';
    if (path === '/omdedy') return 'SCHEDULE';
    if (path === '/audit') return 'AUDIT';
    if (path === '/personnel') return 'PERSONEL';
    if (path === '/reschedule') return 'RESCHEDULE';
    if (path === '/login') return 'LOGIN';
    if (path.startsWith('/project/')) return 'GANTT_DETAIL';
    return viewState;
  }, [location.pathname, viewState]);

  const setActiveView = (view: AppView) => {
    setViewState(view);
    switch(view) {
      case 'PROJECTS': navigate('/portfolio'); break;
      case 'KANBAN': navigate('/kanban'); break;
      case 'SCHEDULE': navigate('/omdedy'); break;
      case 'AUDIT': navigate('/audit'); break;
      case 'PERSONEL': navigate('/personnel'); break;
      case 'RESCHEDULE': navigate('/reschedule'); break;
      case 'GANTT_DETAIL': 
        if (selectedProjectId) navigate(`/project/${selectedProjectId}`);
        else navigate('/timeline');
        break;
    }
  };

  // RBAC Redirect Protection
  useEffect(() => {
    if (!authLoading && user) {
      const path = location.pathname;
      const isRestrictedPath = path === '/personnel' || path === '/reschedule' || path === '/audit';
      if (isRestrictedPath && !isAdmin) {
        navigate('/portfolio');
      }
    }
  }, [location.pathname, isAdmin, user, authLoading]);

  // Sync selectedProjectId from URL
  useEffect(() => {
    const match = location.pathname.match(/\/project\/(.+)/);
    if (match) {
      if (match[1] === 'global') {
        setSelectedProjectId(null);
      } else {
        setSelectedProjectId(match[1]);
      }
    } else {
      setSelectedProjectId(null);
    }
  }, [location.pathname]);

  const [notif, setNotif] = useState<string | null>(null);
  
  const [projects, setProjects] = useState<Project[]>([]);
  const [users, setUsers] = useState<AppUser[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [allAuditLogs, setAllAuditLogs] = useState<AuditLog[]>([]);

  // Automatic "Project Late" Detection
  useEffect(() => {
    const checkLateProjects = async () => {
      const lateProjects = projects.filter(p => {
        if (p.status === 'Project Late') return false;
        const projectTasks = tasks.filter(t => t.project_id === p.id);
        if (projectTasks.length === 0) return false;
        
        // Find the latest task end date
        const latestEnd = Math.max(...projectTasks.map(t => new Date(t.end_time).getTime()));
        return latestEnd < new Date().getTime();
      });

      for (const p of lateProjects) {
        await taskService.updateProject(p.id, { status: ProjectStatus.PROJECT_LATE }, 'System Monitor');
      }
      if (lateProjects.length > 0) setRefreshKey(prev => prev + 1);
    };

    if (projects.length > 0 && tasks.length > 0) {
      checkLateProjects();
    }
  }, [projects, tasks]);

  
  const [showAuditLog, setShowAuditLog] = useState(false);
  const [selectedTask, setSelectedTask] = useState<Task | null>(null);
  const [auditLogs, setAuditLogs] = useState<AuditLog[]>([]);
  
  const [loading, setLoading] = useState(true);
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);
  
  const [scale, setScale] = useState<ViewScale>('DAY');
  const [expandedRows, setExpandedRows] = useState<Set<string>>(new Set());
  const [refreshKey, setRefreshKey] = useState(0);
  const [rescheduleRequests, setRescheduleRequests] = useState<any[]>([]);
  const [requestsLoading, setRequestsLoading] = useState(false);

  const fetchRescheduleRequests = async () => {
    if (!canAccessReschedule) return;
    setRequestsLoading(true);
    try {
      const data = await taskService.getRescheduleRequests();
      setRescheduleRequests(data || []);
      // BUG FIX: Count only pending requests for the badge
      const pendingCount = (data || []).filter(r => r.status === 'Pending').length;
      setPendingRescheduleCount(pendingCount);
    } catch (err) {
      console.error("Failed to fetch reschedule requests:", err);
    } finally {
      setRequestsLoading(false);
    }
  };
  const [isCreateProjectModalOpen, setIsCreateProjectModalOpen] = useState(false);
  const [reschedulingProject, setReschedulingProject] = useState<Project | null>(null);
  const [itemToDelete, setItemToDelete] = useState<{id: string, type: 'task' | 'project' | 'user' | 'phase' | 'subtask', phaseIdx?: number, subIdx?: number} | null>(null);

  const requestDeleteTask = (id: string) => {
    setItemToDelete({ id, type: 'task' });
  };

  const confirmDelete = () => {
    if (!itemToDelete) return;
    const { id, type, phaseIdx, subIdx } = itemToDelete;
    
    if (type === 'task') handleDeleteTask(id);
    else if (type === 'project') handleDeleteProject(id);
    else if (type === 'user') handleDeleteUser(id);
    
    setItemToDelete(null);
  };

  // Auth guard temporarily disabled for setup purposes
  // if (authLoading) return (
  //   <div className="min-h-screen bg-[#020617] flex items-center justify-center">
  //     <div className="w-8 h-8 border-4 border-indigo-500/20 border-t-indigo-500 rounded-full animate-spin" />
  //   </div>
  // );

  // if (!user) return <LoginPage />;

  const fetchData = async () => {
    try {
      const [p, u, t, a] = await Promise.all([
        taskService.getProjects(),
        taskService.getUsers(),
        taskService.getAllTasks(),
        taskService.getAuditLogs()
      ]);
      setProjects(p || []);
      setUsers(u || []);
      setTasks(t || []);
      setAllAuditLogs(a || []);
    } catch (err) {
      console.error("Fetch failed:", err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchData();

    const channel = supabase
      .channel('portfolio_realtime')
      .on('postgres_changes', { event: '*', schema: 'public' }, () => setRefreshKey(prev => prev + 1))
      .subscribe();

    const requestChannel = supabase.channel('reschedule_realtime')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'reschedule_requests' }, () => fetchRescheduleRequests())
      .subscribe();

    fetchRescheduleRequests();

    return () => {
      supabase.removeChannel(channel);
      supabase.removeChannel(requestChannel);
    };
  }, [refreshKey]);

  // Bootstrap Superadmin "fachrul wisnu"
  useEffect(() => {
    const bootstrap = async () => {
      if (loading || users.length === 0) return;
      const targetEmail = 'fachrulwisnunovianto@gmail.com';
      const exists = users.some(u => u.email?.toLowerCase() === targetEmail.toLowerCase());
      
      if (!exists) {
        console.log("Bootstrapping Superadmin: fachrul wisnu");
        try {
          await taskService.createUser({
            name: 'fachrul wisnu',
            email: targetEmail,
            access_level: 'Superadmin',
            role: 'Product Manager',
            password: 'bosskubabi'
          }, 'System Bootstrap');
          setRefreshKey(prev => prev + 1);
        } catch (err) {
          console.error("Bootstrap failed:", err);
        }
      }
    };
    bootstrap();
  }, [loading, users.length]);

  const currentProject = useMemo(() => 
    (projects || []).find(p => p.id === selectedProjectId), 
  [projects, selectedProjectId]);

  const currentUserProfile = useMemo(() => {
    if (!user || !user.email || (users || []).length === 0) return null;
    return (users || []).find(u => u.email?.toLowerCase() === user.email?.toLowerCase());
  }, [user, users]);

  const filteredTasks = useMemo(() => 
    (tasks || []).filter(t => t.project_id === selectedProjectId),
  [tasks, selectedProjectId]);

  const hierarchicalTasks = useMemo(() => {
    const map = new Map<string, Task[]>();
    const roots: Task[] = [];
    
    (filteredTasks || []).forEach(t => {
      if (t.parent_id) {
        const children = map.get(t.parent_id) || [];
        children.push(t);
        map.set(t.parent_id, children);
      } else {
        roots.push(t);
      }
    });

    return { roots, map };
  }, [filteredTasks]);

  const [ownershipModal, setOwnershipModal] = useState<{
    isOpen: boolean;
    picName: string;
    onConfirm: () => void;
    onCancel: () => void;
    itemName: string;
    type: 'task' | 'schedule';
  }>({
    isOpen: false,
    picName: '',
    onConfirm: () => {},
    onCancel: () => {},
    itemName: '',
    type: 'task'
  });

  const handleUpdateTask = async (id: string, field: keyof Task, value: any) => {
    // Get existing task for audit/date check
    const task = tasks.find(t => t.id === id);
    if (!task) return;
    const oldValue = task[field];

    // Ownership check: If not owner, show warning modal
    const isOwner = user?.name === task.assignee || user?.email === task.assignee;
    if (user && !isOwner && !isAdmin && !isSuperadmin) {
      setOwnershipModal({
        isOpen: true,
        picName: task.assignee || 'PIC',
        itemName: task.title,
        type: 'task',
        onCancel: () => {
          setOwnershipModal(prev => ({ ...prev, isOpen: false }));
          fetchData(); // Revert local state
        },
        onConfirm: async () => {
          setOwnershipModal(prev => ({ ...prev, isOpen: false }));
          await executeTaskUpdate(id, field, value, task, oldValue);
        }
      });
      return;
    }

    await executeTaskUpdate(id, field, value, task, oldValue);
  };

  const executeTaskUpdate = async (id: string, field: keyof Task, value: any, task: any, oldValue: any) => {
    // --- UNIFIED BI-DIRECTIONAL SYNC ENGINE (Mathematical Precision) ---
    if (field === 'start_time' || field === 'end_time') {
      const newValueStr = sanitizeDate(value); 
      if (value !== null && newValueStr === null && value !== '') return;

      const proj = projects.find(p => p.id === task.project_id);
      
      // 1. PROJECT LEVEL AUTO-EXPANSION (Bottom-Up)
      if (proj && newValueStr) {
        const pStartStr = proj.start_date || '';
        const pEndStr = proj.end_date || '';
        let needsProjectUpdate = false;
        let projectUpdatePayload: any = {};

        if (field === 'start_time' && (!pStartStr || newValueStr < pStartStr)) {
          projectUpdatePayload.start_date = newValueStr;
          needsProjectUpdate = true;
        } else if (field === 'end_time' && (!pEndStr || newValueStr > pEndStr)) {
          projectUpdatePayload.end_date = newValueStr;
          needsProjectUpdate = true;
        }

        if (needsProjectUpdate) {
          try {
            const { data: updatedProj, error: prjErr } = await supabase
              .from('projects')
              .update(projectUpdatePayload)
              .eq('id', proj.id)
              .select().single();
            
            if (prjErr) throw prjErr;
            setProjects(prev => prev.map(p => p.id === proj.id ? updatedProj : p));
          } catch (err) {
            console.error('Project expansion failed:', err);
          }
        }
      }

      // 2. BI-DIRECTIONAL BIAS: PARENT L1 AUTO-SYNC
      if (task.parent_id) {
        // L2 -> L1 Bottom-Up Expansion
        const parent = tasks.find(t => t.id === task.parent_id);
        if (parent) {
          const l1Start = parent.start_time;
          const l1End = parent.end_time;
          let needsL1Update = false;
          let l1UpdatePayload: any = {};

          if (field === 'start_time' && (!l1Start || newValueStr < l1Start)) {
            l1UpdatePayload.start_time = newValueStr;
            needsL1Update = true;
          } else if (field === 'end_time' && (!l1End || newValueStr > l1End)) {
            l1UpdatePayload.end_time = newValueStr;
            needsL1Update = true;
          }

          if (needsL1Update) {
            try {
              const { data: updatedL1, error: l1Err } = await supabase
                .from('tasks')
                .update(l1UpdatePayload)
                .eq('id', parent.id)
                .select().single();
              if (l1Err) throw l1Err;
              setTasks(prev => prev.map(t => t.id === parent.id ? updatedL1 : t));
            } catch (err) {
              console.error('L1 expansion failed:', err);
            }
          }
        }
      } else {
        // L1 -> L2 Top-Down Push (Enforce boundaries on children)
        const children = tasks.filter(t => t.parent_id === id);
        if (children.length > 0 && newValueStr) {
          const isStart = field === 'start_time';
          const outOfBoundsChildren = children.filter(c => {
            const cDate = isStart ? c.start_time : c.end_time;
            if (!cDate) return false;
            return isStart ? cDate < newValueStr : cDate > newValueStr;
          });

          for (const child of outOfBoundsChildren) {
             const childField = isStart ? 'start_time' : 'end_time';
             try {
                const { data: updatedChild } = await supabase
                  .from('tasks')
                  .update({ [childField]: newValueStr, updated_at: new Date().toISOString() })
                  .eq('id', child.id)
                  .select().single();
                if (updatedChild) {
                  setTasks(prev => prev.map(t => t.id === child.id ? updatedChild : t));
                }
             } catch (e) {
                console.error("Top-down sync failed for child:", child.id);
             }
          }
        }
      }
    }

    // --- EXECUTE PRIMARY UPDATE ---
    try {
       const actorName = currentUser?.name || user?.name || user?.email || 'User';
       const sanitizedVal = (field === 'start_time' || field === 'end_time' || field === 'target_sla_date') 
         ? sanitizeDate(value) 
         : value;

       const { data: updatedTask, error: updateError } = await supabase
         .from('tasks')
         .update({ [field]: sanitizedVal, updated_at: new Date().toISOString() })
         .eq('id', id)
         .select()
         .single();
          
       if (updateError) throw updateError;
       
       if (updatedTask) {
         setTasks((prev: Task[]) => prev.map((t: any) => t.id === id ? updatedTask : t));
       }
       
       if (field === 'start_time' || field === 'end_time' || field === 'status') {
         setNotif(`Task "${task.title}" Synced Successfully.`);
       }
    } catch (err: any) {
       console.error('Update failed:', err);
       fetchData(); 
       alert('CRUD Failed (Update Task): ' + err.message);
    }
  };

  const handleOpenAudit = async (task: Task) => {
    setSelectedTask(task);
    setShowAuditLog(true);
    try {
      const logs = await taskService.getAuditLogs({ taskId: task.id });
      setAuditLogs(logs);
    } catch (err) {
      console.error('Audit fetch failed:', err);
    }
  };

  const handleDeleteTask = async (id: string) => {
    // Optimistic Delete - disappear instantly from UI
    // Cascade removal: filter out the task itself AND any tasks that have it as parent_id
    setTasks(prevTasks => prevTasks.filter(t => t.id !== id && t.parent_id !== id));
    
    try {
      const { error } = await supabase.from('tasks').delete().eq('id', id);
      if (error) throw error;

      // Also log audit via service for traceability
      await taskService.logAudit({ 
        task_id: id, 
        actor: user?.email || 'Administrator', 
        action: 'DELETED' 
      });
      
    } catch (err: any) {
      console.error("Delete failed:", err);
      // Revert if failed to ensure data integrity
      fetchData(); 
      setNotif("Gagal menghapus task: " + err.message);
    }
  };

  const handleAddUser = async () => {
    try {
      await taskService.createUser({ 
        name: 'New Personnel...', 
        access_level: 'PIC', 
        role: 'Staff' 
      }, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to add user:", err);
    }
  };

  const handleDeleteUser = async (id: string) => {
    try {
      await taskService.deleteUser(id, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to delete user:", err);
    }
  };

  const handleUpdateUser = async (id: string, field: keyof AppUser, value: any) => {
    try {
      await taskService.updateUser(id, { [field]: value }, 'User');
      setRefreshKey(prev => prev + 1);
    } catch (err: any) {
      console.error("Failed to update user:", err);
    }
  };

  // --- Auto-Sync Level 1 Tasks to Project Header ---
  useEffect(() => {
    if (loading || projects.length === 0 || tasks.length === 0 || isSyncingRef.current) return;

    const syncProjects = async () => {
      isSyncingRef.current = true;
      try {
        for (const project of projects) {
          const l1Tasks = tasks.filter(t => t.project_id === project.id && !t.parent_id);
          if (l1Tasks.length === 0) continue;

          const startDates = l1Tasks.filter(t => t.start_time).map(t => new Date(t.start_time).getTime());
          const endDates = l1Tasks.filter(t => t.end_time).map(t => new Date(t.end_time).getTime());
          
          if (startDates.length === 0 || endDates.length === 0) continue;

          const minStart = formatDateForDB(new Date(Math.min(...startDates))) || '';
          const maxEnd = formatDateForDB(new Date(Math.max(...endDates))) || '';

          const currentStart = project.start_date || '';
          const currentEnd = project.end_date || '';

          // ONE-WAY EXPANSION ONLY:
          const needsStartUpdate = currentStart && minStart < currentStart;
          const needsEndUpdate = currentEnd && maxEnd > currentEnd;
          const isInitial = !currentStart || !currentEnd;

          if (needsStartUpdate || needsEndUpdate || isInitial) {
             const updatePayload: any = {};
             let logOldStart = currentStart;
             let logOldEnd = currentEnd;
             
             if (needsStartUpdate || (!currentStart && minStart)) updatePayload.start_date = minStart;
             if (needsEndUpdate || (!currentEnd && maxEnd)) updatePayload.end_date = maxEnd;
             
             if (Object.keys(updatePayload).length > 0) {
               console.log(`Syncing project ${project.name} timeline (Expansion Only):`, updatePayload);
               await handleUpdateProject(project.id, updatePayload);
               
               await taskService.createProjectRescheduleLog({
                 project_id: project.id,
                 changed_by: 'System Logic (Sync)',
                 old_start_date: logOldStart,
                 old_end_date: logOldEnd,
                 new_start_date: updatePayload.start_date || logOldStart,
                 new_end_date: updatePayload.end_date || logOldEnd,
                 reason: 'Auto-Expansion: One-way sync from WBS boundaries.'
               });
             }
          }
        }
      } finally {
        isSyncingRef.current = false;
      }
    };

    const timer = setTimeout(() => {
      syncProjects();
    }, 1500); // Increased debounce to prevent collision with manual header edits

    return () => clearTimeout(timer);
  }, [tasks, projects.length, loading]);
  
  const handleUpdateProject = async (id: string, updates: Partial<Project>) => {
    const oldProject = projects.find(p => p.id === id);
    // Optimistic Update
    setProjects(prev => prev.map(p => p.id === id ? { ...p, ...updates } : p));
    
    try {
      const { data, error } = await supabase
        .from('projects')
        .update(updates)
        .eq('id', id)
        .select()
        .single();
        
      if (error) throw error;
      
      // Immediate state sync with server data
      if (data) {
        setProjects(prev => prev.map(p => p.id === id ? data : p));
      }
      
      // Audit Logging for Timeline Changes
      if (oldProject && (updates.start_date || updates.end_date)) {
        await taskService.createProjectRescheduleLog({
          project_id: id,
          changed_by: user?.name || user?.email || 'Anonymous',
          old_start_date: oldProject.start_date || '',
          new_start_date: updates.start_date || oldProject.start_date || '',
          old_end_date: oldProject.end_date || '',
          new_end_date: updates.end_date || oldProject.end_date || '',
          reason: 'Auto-sync from WBS Inline Edit'
        });
      }
    } catch (err: any) {
      console.error("Failed to update project:", err);
      fetchData(); // Rollback
    }
  };

  const handleDeleteProject = async (id: string) => {
    // Optimistic Delete - disappear instantly
    setProjects(prev => prev.filter(p => p.id !== id));
    if (selectedProjectId === id) setSelectedProjectId(null);

    try {
      const { error } = await supabase.from('projects').delete().eq('id', id);
      if (error) throw error;
      setNotif('Project berhasil dihapus');
    } catch (err: any) {
      console.error("Failed to delete project:", err);
      setNotif('Gagal menghapus project');
      fetchData(); // Rollback
    }
  };

  const handleToggleExpand = (id: string) => {
    const next = new Set(expandedRows);
    if (next.has(id)) next.delete(id);
    else next.add(id);
    setExpandedRows(next);
  };

  const menuItems = useMemo(() => {
    const base = [
      { id: 'PROJECTS', label: 'Project List', icon: LayoutDashboard },
      { id: 'GANTT_DETAIL', label: 'Om Dedy Timeline', icon: Activity },
      { id: 'KANBAN', label: 'Status Monitoring', icon: LayoutGrid },
      { id: 'SCHEDULE', label: 'Om Dedy Schedule', icon: Calendar },
      { id: 'RESCHEDULE', label: 'APPROVALS', icon: History },
      { id: 'PERSONEL', label: 'Personel OM DEDY', icon: Users },
      { id: 'AUDIT', label: 'System Audit Logs', icon: ShieldAlert },
    ];

    const filtered = base.filter(item => {
      // Only Admin/Superadmin can see Personnel and Reschedule menus
      if (item.id === 'PERSONEL' || item.id === 'RESCHEDULE' || item.id === 'AUDIT') return isAdmin;
      return true;
    });

    return filtered;
  }, [isAdmin]);

  if (loading && projects.length === 0) {
    return (
      <div className="fixed inset-0 bg-[#0B1120] flex flex-col items-center justify-center z-[9999]">
        <motion.div 
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          className="flex flex-col items-center"
        >
          {/* Logo with Glow */}
          <div className="w-20 h-20 bg-indigo-600 rounded-2xl flex items-center justify-center mb-6 shadow-[0_0_40px_rgba(79,70,229,0.4)] relative overflow-hidden group">
             <span className="text-3xl font-black text-white relative z-10">OD</span>
             <div className="absolute inset-0 bg-gradient-to-tr from-white/20 to-transparent" />
          </div>

          {/* Branded Text */}
          <h1 className="text-3xl font-black tracking-tighter text-white mb-2">
            OM <span className="text-indigo-500">DEDY</span>
          </h1>
          <p className="text-slate-400 text-sm font-medium tracking-widest uppercase">
            Operational Monitoring Dashboard
          </p>
          <p className="text-slate-500 text-[10px] mt-1 font-bold tracking-[0.2em] uppercase">
            FOR EFFICIENT DELIVERY
          </p>

          {/* Modern Progress Bar */}
          <div className="w-48 h-1 bg-slate-800 rounded-full mt-10 overflow-hidden relative">
             <div 
               className="h-full bg-indigo-500 absolute top-0 left-0 animate-[loading_1.5s_infinite] w-1/3 shadow-[0_0_10px_#6366f1]" 
             />
          </div>
          
          <p className="text-[9px] text-slate-600 mt-6 font-mono uppercase tracking-widest animate-pulse">
            Synchronizing Critical Systems...
          </p>
        </motion.div>
      </div>
    );
  }

  if (activeView === 'LOGIN') {
    return <LoginPage />;
  }

  return (
    <div className="min-h-screen bg-[#020617] text-slate-200 flex overflow-hidden">
      {/* Sidebar */}
      <aside className={cn(
        "bg-slate-900 border-r border-slate-800 transition-all duration-300 flex flex-col z-40",
        isSidebarOpen ? "w-72" : "w-16"
      )}>
        <div className="h-20 flex items-center px-4 border-b border-slate-800 shrink-0">
          <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center shrink-0 shadow-lg shadow-indigo-500/20 group cursor-pointer hover:scale-105 transition-transform">
            <span className="text-lg font-black text-white">OD</span>
          </div>
          {isSidebarOpen && (
            <div className="ml-3 overflow-hidden">
              <h1 className="font-black text-white uppercase italic tracking-tighter text-xl leading-none">OM <span className="text-indigo-500">DEDY</span></h1>
              <p className="text-[7px] text-slate-500 font-black tracking-[0.05em] uppercase leading-tight mt-1">Operational Monitoring Dashboard<br/>for Efficient Delivery</p>
            </div>
          )}
        </div>

        <nav className="flex-1 py-6 space-y-1">
            {menuItems.map((item, i) => (
              <React.Fragment key={getSafeKey({id: item.id}, i, 'nav-item')}>
                <button
                onClick={() => { setActiveView(item.id as AppView); setSelectedProjectId(null); }}
                className={cn(
                  "w-full flex items-center py-3 px-4 transition-all relative group",
                  (activeView === item.id && !selectedProjectId)
                    ? "text-white bg-gradient-to-r from-indigo-600/20 to-transparent" 
                    : "text-slate-400 hover:text-slate-200 hover:bg-white/5"
                )}
              >
                <item.icon className={cn("w-5 h-5 shrink-0 transition-colors", (activeView === item.id && !selectedProjectId) ? "text-indigo-400" : "text-slate-500")} />
                {isSidebarOpen && <span className="ml-4 font-bold text-xs uppercase tracking-widest">{item.label}</span>}
                {item.id === 'RESCHEDULE' && pendingRescheduleCount > 0 && (
                  <span className={cn(
                    "ml-auto bg-rose-500 text-white text-[8px] font-black rounded-full flex items-center justify-center animate-bounce shadow-lg ring-2 ring-slate-900 transition-all",
                    isSidebarOpen ? "w-5 h-5" : "w-4 h-4 absolute top-1 right-1"
                  )}>
                    {pendingRescheduleCount}
                  </span>
                )}
                {!isSidebarOpen && (
                  <div className="absolute left-full ml-2 px-3 py-1 bg-slate-800 text-white text-[10px] rounded opacity-0 group-hover:opacity-100 whitespace-nowrap shadow-xl pointer-events-none z-50">
                    {item.label}
                  </div>
                )}
                {((activeView === item.id && !selectedProjectId)) && <div className="absolute left-0 top-0 bottom-0 w-1 bg-indigo-500" />}
              </button>

              {/* Dynamic Sub-menu for Project Detail */}
              {item.id === 'PROJECTS' && selectedProjectId && (
                <div className={cn("mt-1 mb-2", isSidebarOpen ? "pl-8" : "pl-4")}>
                  <div className="flex items-center gap-2 py-2 text-[10px] font-black text-emerald-400 uppercase tracking-widest bg-emerald-500/5 border-l-2 border-emerald-500 pl-3">
                    <div className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
                    {isSidebarOpen ? 'Om Dedy - Detail Timeline' : ''}
                  </div>
                </div>
              )}
            </React.Fragment>
          ))}
        </nav>

        <div className="p-4 border-t border-slate-800 space-y-4">
           {isSidebarOpen && user && (
             <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800">
               <div className="flex items-center gap-3 mb-3">
                 <div className="w-8 h-8 rounded-full bg-slate-800 border border-indigo-500/30 flex items-center justify-center overflow-hidden">
                   <span className="text-[10px] font-black text-indigo-400 capitalize">{user.email?.charAt(0) || '?'}</span>
                 </div>
                 <div className="flex-1 min-w-0">
                   <p className="text-[10px] font-bold text-slate-200 truncate">{user?.email}</p>
                   <p className="text-[9px] text-slate-500 uppercase tracking-widest font-bold">{user?.access_level || 'Authenticated PIC'}</p>
                 </div>
               </div>
               <button 
                 onClick={() => signOut()}
                 className="w-full flex items-center justify-center gap-2 py-2 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 rounded-lg text-[10px] font-bold uppercase transition-all"
               >
                 <LogOut className="w-3.5 h-3.5" /> Log Out
               </button>
             </div>
           )}

           {isSidebarOpen && !user && (
             <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800">
               <p className="text-[9px] text-slate-500 font-bold uppercase tracking-widest text-center mb-3">Viewing as External</p>
               <button 
                 onClick={() => navigate('/login')}
                 className="w-full flex items-center justify-center gap-2 py-2 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-[10px] font-black uppercase tracking-widest shadow-lg shadow-indigo-600/20 transition-all"
               >
                 <UserIcon className="w-3.5 h-3.5" /> Personnel Login
               </button>
             </div>
           )}
           <button 
             onClick={() => setIsSidebarOpen(!isSidebarOpen)}
             className="w-full py-2 flex items-center justify-center rounded-lg hover:bg-slate-800 transition-colors"
           >
             {isSidebarOpen ? <ChevronLeft className="w-5 h-5 text-slate-600" /> : <ChevronRight className="w-5 h-5 text-slate-600" />}
           </button>
        </div>
      </aside>

      {/* Content Area */}
      <main className="flex-1 flex flex-col overflow-hidden relative">
        {loading && (
          <div className="absolute inset-0 bg-slate-950/60 backdrop-blur-[4px] z-[100] flex items-center justify-center">
            <div className="flex flex-col items-center gap-4">
              <div className="w-10 h-10 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin shadow-[0_0_20px_rgba(99,102,241,0.3)]" />
              <div className="text-center">
                <span className="text-[10px] font-black uppercase tracking-[0.3em] text-indigo-500">Syncing Node</span>
                <p className="text-[9px] text-slate-500 font-bold uppercase tracking-widest mt-1">Operational Monitoring Active</p>
              </div>
            </div>
          </div>
        )}
        <header className="h-20 border-b border-slate-800/60 flex items-center justify-between px-8 bg-slate-950/50 backdrop-blur-md z-30 shrink-0">
          <div className="flex items-center gap-4">
             {activeView === 'GANTT_DETAIL' && selectedProjectId && (
               <button 
                 onClick={() => { setActiveView('PROJECTS'); setSelectedProjectId(null); }}
                 className="p-2 hover:bg-slate-800 rounded-lg transition-colors text-slate-500 hover:text-white"
               >
                 <ArrowLeft className="w-5 h-5" />
               </button>
             )}
             <div>
               <h2 className="text-xl font-black text-white tracking-tighter flex items-center gap-3 uppercase italic">
                 {activeView === 'PROJECTS' && <span>OM <span className="text-indigo-500">DEDY</span></span>}
                 {activeView === 'KANBAN' && 'Project Status Monitoring'}
                 {activeView === 'PERSONEL' && 'Personel OM DEDY'}
                 {activeView === 'RESCHEDULE' && 'Reschedule Om Dedy'}
                 {activeView === 'SCHEDULE' && 'Om Dedy Schedule'}
                 {activeView === 'AUDIT' && 'System Audit Rails'}
                 {activeView === 'GANTT_DETAIL' && (
                    selectedProjectId ? (
                      <div className="flex items-center gap-3">
                        <span>{projects.find(p => p.id === selectedProjectId)?.name}</span>
                        <div className="flex items-center gap-2 not-italic">
                          {(() => {
                            const p = projects.find(prj => prj.id === selectedProjectId);
                            const pTasks = tasks.filter(t => t.project_id === selectedProjectId);
                            const leafTasks = pTasks.filter(t => !pTasks.some(other => other.parent_id === t.id));
                            const totalHours = leafTasks.reduce((sum, t) => sum + (Number(t.man_hours) || 0), 0);
                            return (
                              <>
                                <div className="bg-slate-800/80 px-2 py-1 rounded-md border border-slate-700 flex items-center gap-1.5 shadow-sm">
                                  <Clock className="w-3 h-3 text-indigo-400" />
                                  <span className="text-[10px] font-black text-slate-300 tracking-wider uppercase">
                                    ⏱️ {totalHours.toFixed(1)} HOURS
                                  </span>
                                </div>
                                <div className={cn(
                                  "px-2 py-1 rounded-md border text-[10px] font-black tracking-widest uppercase",
                                  p?.status === ProjectStatus.DONE ? "bg-emerald-500/10 border-emerald-500/30 text-emerald-400" :
                                  p?.status === ProjectStatus.ACTIVE ? "bg-amber-500/10 border-amber-500/30 text-amber-400" :
                                  "bg-slate-800/50 border-slate-700 text-slate-500"
                                )}>
                                  {p?.status === ProjectStatus.ACTIVE ? 'FSD ON PROGRESS' : p?.status || 'UNKNOWN'}
                                </div>
                              </>
                            );
                          })()}
                        </div>
                      </div>
                    ) : 'Om Dedy Timeline'
                  )}
               </h2>
               <p className="text-[9px] text-slate-600 font-black uppercase tracking-[0.3em]">WITH OM DEDY EVERYTHING WILL BE DELIVERED</p>
             </div>
          </div>

             <div className="flex items-center gap-4">
              {activeView === 'GANTT_DETAIL' && (
                <div className="flex bg-slate-900 p-1 rounded-lg border border-slate-800">
                  {(['MONTH', 'WEEK', 'DAY', 'HOUR'] as ViewScale[]).map((s, si) => (
                    <button
                      key={`header-scale-${s}-${si}`}
                      onClick={() => setScale(s)}
                      className={cn(
                        "px-3 py-1 rounded-md text-[10px] font-bold tracking-wider transition-all",
                        scale === s ? "bg-slate-800 text-indigo-400 shadow-sm" : "text-slate-500 hover:text-slate-300"
                      )}
                    >
                      {s}
                    </button>
                  ))}
                </div>
              )}

              {(activeView === 'GANTT_DETAIL' || activeView === 'PROJECTS') && (
                <button 
                  onClick={() => {
                    navigator.clipboard.writeText(window.location.href);
                    setNotif("🔗 Public Link Copied to Clipboard");
                  }}
                  className="flex items-center gap-2 bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest border border-slate-700 transition-all group"
                >
                  <ExternalLink className="w-3.5 h-3.5 text-indigo-400 group-hover:scale-110 transition-transform" />
                  Copy Public Link
                </button>
              )}

            {user && (
              <button 
                onClick={() => setIsProfileModalOpen(true)}
                className="p-2 text-slate-400 hover:text-white transition-colors"
              >
                <SettingsIcon className="w-5 h-5" />
              </button>
            )}
          </div>
        </header>

        <section className={cn("flex-1 overflow-auto scrollbar-hide", activeView === 'GANTT_DETAIL' ? 'p-4' : 'p-8')}>
          <ErrorBoundary>
           {activeView === 'PROJECTS' && (
             <PortfolioDashboard 
               user={user}
               projects={projects} 
               tasks={tasks}
               loading={loading}
               onOpenProject={(id) => { setSelectedProjectId(id); navigate(`/project/${id}`); }} 
               onDeleteProject={handleDeleteProject}
               onUpdateProject={handleUpdateProject}
               onCreateRequested={() => setIsCreateProjectModalOpen(true)}
                onReschedule={(p) => setReschedulingProject(p)}
             />
           )}
           {activeView === 'SCHEDULE' && (
              <OmDedySchedule 
                user={user} 
                users={users} 
                setActiveView={setActiveView} 
                isAdmin={isAdmin}
                isSuperadmin={isSuperadmin}
                setOwnershipModal={setOwnershipModal}
                onNotif={setNotif}
              />
            )}
           {activeView === 'KANBAN' && (
             <KanbanView 
               projects={projects} 
               tasks={tasks}
               onOpenGantt={(id) => { setSelectedProjectId(id); navigate(`/project/${id}`); }}
               onUpdateProject={handleUpdateProject}
             />
           )}
           {activeView === 'PERSONEL' && (
             <PersonelManagement isAdmin={isAdmin} 
               users={users} 
               projects={projects}
               currentUser={user}
               onRefresh={() => setRefreshKey(prev => prev + 1)}
             />
           )}
           {activeView === 'RESCHEDULE' && (
             <RescheduleRequestsView 
               requests={rescheduleRequests}
               isLoading={requestsLoading}
               onRefresh={() => {
                 fetchRescheduleRequests();
                 setRefreshKey(prev => prev + 1);
               }}
               user={user}
             />
           )}
           {activeView === 'AUDIT' && <AuditLogView logs={allAuditLogs} projects={projects} users={users} />}
           {activeView === 'GANTT_DETAIL' && (
             <GanttDetailView
               user={user}
               setProjects={setProjects}
               projectId={selectedProjectId}
               setFocusedProjectId={(id: string) => navigate(`/project/${id}`)}
               projects={projects}
               tasks={tasks}
               hierarchicalTasks={hierarchicalTasks}
               expandedRows={expandedRows}
               scale={scale}
               setRefreshKey={setRefreshKey}
               handleToggleExpand={handleToggleExpand}
               handleUpdateTask={handleUpdateTask}
               handleOpenAudit={handleOpenAudit}
               handleDeleteTask={requestDeleteTask}
               setScale={setScale}
               onReschedule={handleOpenReschedule}
               onNotif={setNotif}
               users={users}
               setTasks={setTasks}
             />
           )}
          </ErrorBoundary>
        </section>
      </main>

      <AnimatePresence>
        {isCreateProjectModalOpen && user && (
          <ErrorBoundary key="create-project-wizard-error-boundary">
            <CreateProjectModal 
              key="create-project-wizard-modal"
              user={user}
              users={users}
              onClose={() => setIsCreateProjectModalOpen(false)} 
              onSuccess={() => {
                setRefreshKey(prev => prev + 1);
                fetchData();
                setNotif("Project Created Successfully!");
              }} 
            />
          </ErrorBoundary>
        )}
        {showAuditLog && selectedTask && (
          <React.Fragment key="audit-trail-fragment">
            <motion.div 
              key="audit-trail-overlay"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowAuditLog(false)}
              className="fixed inset-0 bg-black/60 backdrop-blur-sm z-40"
            />
            <motion.div 
              key="audit-trail-panel"
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              className="fixed top-0 right-0 bottom-0 w-[450px] bg-slate-900 border-l border-slate-800 shadow-2xl z-50 overflow-y-auto"
            >
              <div className="p-6">
                <div className="flex items-center justify-between mb-8">
                  <div>
                    <h2 className="text-xl font-bold text-white flex items-center gap-2">
                      <History className="w-5 h-5 text-indigo-500" />
                      Immutable Audit Trail
                    </h2>
                    <p className="text-xs text-slate-500 mt-1">Governance Layer for {selectedTask.title}</p>
                  </div>
                  <button onClick={() => setShowAuditLog(false)} className="text-slate-400 hover:text-white">
                    <Plus className="w-6 h-6 rotate-45" />
                  </button>
                </div>

                <div className="space-y-6">
                    {auditLogs.map((log, i) => {
                      if (!log) return null;
                      const logId = getSafeKey(log, i, 'audit-log');
                      return (
                        <div key={logId} className="relative pl-8 pb-6 group">
                        {i !== (auditLogs || []).length - 1 && (
                          <div className="absolute left-[11px] top-6 bottom-0 w-[2px] bg-slate-800 group-last:hidden" />
                        )}
                        <div className="absolute left-0 top-1 w-6 h-6 rounded-full bg-slate-800 border-2 border-slate-700 flex items-center justify-center z-10">
                          <div className="w-2 h-2 rounded-full bg-indigo-500" />
                        </div>
                        
                        <div className="bg-slate-800/40 border border-slate-800 p-4 rounded-xl">
                          <div className="flex items-center justify-between mb-2">
                            <span className="text-xs font-bold text-indigo-400">{log.actor || 'System'}</span>
                            <span className="text-[10px] text-slate-500 font-mono">{log.created_at ? format(new Date(log.created_at), 'MM/dd HH:mm:ss') : 'N/A'}</span>
                          </div>
                          <p className="text-sm font-medium text-slate-200 mb-2">{log.action || 'Unknown Action'}</p>
                          
                          {log.old_payload && log.new_payload && (
                            <div className="mt-4 space-y-2 border-t border-slate-700/50 pt-4">
                              <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-2">Payload Comparison</p>
                              {(() => {
                                const oldP = log.old_payload as any;
                                const newP = log.new_payload as any;
                                const allKeys = Array.from(new Set([...Object.keys(oldP), ...Object.keys(newP)]));
                                
                                return allKeys.map((key, ki) => {
                                  if (key === 'updated_at' || key === 'id' || key === 'created_at') return null;
                                  if (JSON.stringify(oldP[key]) === JSON.stringify(newP[key])) return null;
                                  
                                  const diffKey = getSafeKey({id: key}, ki, 'payload-diff');
                                  return (
                                    <div key={diffKey} className="flex flex-col gap-1 pb-2 border-b border-white/5 last:border-0">
                                      <span className="text-[10px] font-mono text-indigo-400/70">{key}</span>
                                      <div className="grid grid-cols-[1fr,auto,1fr] items-center gap-2">
                                        <div className="bg-rose-500/10 text-rose-400 text-[10px] p-1 rounded border border-rose-500/20 line-through opacity-60 truncate">
                                          {String(oldP[key] ?? 'null')}
                                        </div>
                                        <ArrowRight className="w-3 h-3 text-slate-600" />
                                        <div className="bg-emerald-500/10 text-emerald-400 text-[10px] p-1 rounded border border-emerald-500/20 font-bold italic truncate">
                                          {String(newP[key] ?? 'null')}
                                        </div>
                                      </div>
                                    </div>
                                  );
                                });
                              })()}
                            </div>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          </React.Fragment>
        )}
        {reschedulingProject && (
          <ProjectRescheduleModal 
            project={reschedulingProject}
            user={user}
            onClose={() => setReschedulingProject(null)}
            onSuccess={() => {
              setRefreshKey(prev => prev + 1);
              fetchData();
              setNotif("Timeline berhasil digeser dan disinkronkan!");
            }}
          />
        )}
        <ConfirmModal
          isOpen={!!itemToDelete}
          onClose={() => setItemToDelete(null)}
          onConfirm={confirmDelete}
          title="Konfirmasi Hapus"
          description="Apakah Anda yakin ingin menghapus data ini? Semua rincian di dalamnya juga akan ikut terhapus secara permanen. Aksi ini tidak dapat dibatalkan."
        />
        {isProfileModalOpen && user && (
          <ProfileSettingsModal 
            key="profile-settings-modal"
            isOpen={isProfileModalOpen}
            onClose={() => setIsProfileModalOpen(false)}
            user={user}
            onUpdate={() => fetchData()}
          />
        )}
        <AnimatePresence>
          {ownershipModal.isOpen && (
            <OwnershipWarningModal 
              key="ownership-warning-modal"
              picName={ownershipModal.picName}
              itemName={ownershipModal.itemName}
              type={ownershipModal.type}
              onCancel={ownershipModal.onCancel}
              onConfirm={ownershipModal.onConfirm}
            />
          )}
        </AnimatePresence>
      </AnimatePresence>
    </div>
  );
}

function OwnershipWarningModal({ picName, itemName, type, onConfirm, onCancel }: {
  picName: string;
  itemName: string;
  type: 'task' | 'schedule',
  onConfirm: () => void;
  onCancel: () => void;
}) {
  return (
    <div className="fixed inset-0 z-[1000] bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.95, y: 20 }}
        className="bg-slate-900 border border-amber-500/30 rounded-[2.5rem] w-full max-w-md shadow-2xl p-8 relative overflow-hidden"
      >
        <div className="absolute top-0 right-0 w-32 h-32 bg-amber-500/5 rounded-full blur-3xl -translate-y-1/2 translate-x-1/2" />
        
        <div className="relative text-center">
          <div className="w-16 h-16 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20 mx-auto mb-6 shadow-xl">
            <AlertCircle className="w-8 h-8 text-amber-500" />
          </div>
          
          <h2 className="text-2xl font-black text-white italic tracking-tighter uppercase mb-4">⚠️ Peringatan Kepemilikan</h2>
          
          <p className="text-slate-300 text-sm leading-relaxed mb-8 text-center px-4">
            Anda sedang mengedit {type === 'task' ? 'tugas' : 'jadwal'} milik <span className="text-amber-400 font-bold">{picName}</span>. 
            Tindakan ini akan dicatat dalam <span className="text-indigo-400 font-bold underline decoration-indigo-500/30">Audit Log</span> sebagai bukti (evidence) untuk akuntabilitas.
            <br/><br/>
            Apakah Anda yakin ingin melanjutkan update pada <span className="italic">"{itemName}"</span>?
          </p>
          
          <div className="flex gap-4">
            <button 
              onClick={onCancel}
              className="flex-1 bg-slate-800 hover:bg-slate-700 text-slate-400 py-4 rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-slate-700"
            >
              Batal
            </button>
            <button 
              onClick={onConfirm}
              className="flex-1 bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-500 hover:to-orange-500 text-white py-4 rounded-2xl text-[10px] font-black uppercase tracking-widest shadow-xl shadow-amber-900/20 transition-all active:scale-95"
            >
              Lanjutkan Update
            </button>
          </div>
        </div>
      </motion.div>
    </div>
  );
}

// --- Modals ---

function ProjectRescheduleModal({ project, user, onClose, onSuccess }: { project: Project, user: any, onClose: () => void, onSuccess: () => void }) {
  const [newStart, setNewStart] = useState(formatDateForDB(project.start_date) || formatDateForDB(new Date()) || '');
  const [newEnd, setNewEnd] = useState(formatDateForDB(project.end_date) || formatDateForDB(addDays(new Date(), 30)) || '');
  const [reason, setReason] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!reason.trim()) return;
    setLoading(true);
    try {
      // 1. Log the change
      await taskService.createProjectRescheduleLog({
        project_id: project.id,
        changed_by: user?.name || user?.email || 'Admin',
        old_start_date: project.start_date || 'N/A',
        old_end_date: project.end_date || 'N/A',
        new_start_date: newStart,
        new_end_date: newEnd,
        reason: reason
      });

      // 2. Update project
      await taskService.updateProject(project.id, {
        start_date: newStart,
        end_date: newEnd
      }, user?.email || 'Admin');

      onSuccess();
      onClose();
    } catch (err: any) {
      console.error("Failed to reschedule project:", err);
      alert("Error: " + err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[200] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/80 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-full max-w-lg w-[95%] mx-auto bg-slate-900 border border-white/10 rounded-3xl shadow-2xl overflow-hidden max-h-[90vh] overflow-y-auto"
      >
        <div className="p-8 border-b border-white/5 bg-gradient-to-br from-slate-900 to-slate-950">
          <div className="flex justify-between items-start mb-2">
            <div>
              <h3 className="text-2xl font-black text-white italic uppercase tracking-tighter">Reschedule Project</h3>
              <p className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.3em] mt-1">Timeline Governance Layer</p>
            </div>
            <button onClick={onClose} className="text-slate-500 hover:text-white transition-colors">
              <X className="w-6 h-6" />
            </button>
          </div>
          <div className="p-4 bg-indigo-500/5 rounded-2xl border border-indigo-500/10 mt-6">
            <h4 className="text-[10px] font-black text-indigo-400 uppercase tracking-widest mb-1">Target Infrastructure</h4>
            <p className="text-sm font-bold text-slate-200">{project.name}</p>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-6">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-6">
            <div className="space-y-4 opacity-50 grayscale pointer-events-none hidden sm:block">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest block">Current Timeline</label>
              <div className="space-y-3">
                <div className="flex flex-col gap-1.5 p-3 bg-slate-950/50 rounded-xl border border-white/5">
                  <span className="text-[8px] font-bold text-slate-600 uppercase">Start Date</span>
                  <span className="text-xs text-slate-400 font-mono italic">{project.start_date || 'N/A'}</span>
                </div>
                <div className="flex flex-col gap-1.5 p-3 bg-slate-950/50 rounded-xl border border-white/5">
                  <span className="text-[8px] font-bold text-slate-600 uppercase">End Date</span>
                  <span className="text-xs text-slate-400 font-mono italic">{project.end_date || 'N/A'}</span>
                </div>
              </div>
            </div>

            <div className="space-y-4">
              <label className="text-[10px] font-black text-indigo-500 uppercase tracking-widest block">New Timeline Configuration</label>
              <div className="space-y-4">
                <div className="flex flex-col gap-1.5">
                   <label className="text-[8px] font-bold text-slate-500 uppercase ml-1">Start Date</label>
                   <div className="relative group">
                     <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600 group-focus-within:text-indigo-500 transition-colors" />
                     <input 
                       type="date"
                       value={newStart}
                       onChange={(e) => setNewStart(e.target.value)}
                       required
                       className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl pl-10 pr-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
                     />
                   </div>
                </div>
                <div className="flex flex-col gap-1.5">
                   <label className="text-[8px] font-bold text-slate-500 uppercase ml-1">End Date</label>
                   <div className="relative group">
                     <Calendar className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-600 group-focus-within:text-indigo-500 transition-colors" />
                     <input 
                       type="date"
                       value={newEnd}
                       onChange={(e) => setNewEnd(e.target.value)}
                       required
                       className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl pl-10 pr-4 py-3 text-sm focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all"
                     />
                   </div>
                </div>
              </div>
            </div>
          </div>

          <div className="space-y-3 pt-4 border-t border-white/5">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest ml-2">Justification / Reason</label>
            <textarea 
              required
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              className="w-full bg-slate-950/50 border border-slate-800 rounded-2xl p-4 text-sm text-white focus:ring-2 focus:ring-indigo-500/20 focus:border-indigo-500/50 outline-none transition-all placeholder:text-slate-700 min-h-[100px]"
              placeholder="Provide context for this schedule shift..."
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-4 bg-slate-800/50 hover:bg-slate-800 text-slate-400 rounded-2xl text-[10px] font-black uppercase tracking-[0.2em] transition-all border border-white/5"
            >
              Abrupty Cancel
            </button>
            <button 
              type="submit"
              disabled={loading}
              className="flex-[2] py-4 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-[0.2em] rounded-2xl shadow-xl shadow-indigo-600/20 transition-all active:scale-[0.98] disabled:opacity-50"
            >
              {loading ? 'Committing Changes...' : 'Execute Timeline Shift'}
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}




function CreateProjectModal({ onClose, onSuccess, user, users }: { onClose: () => void, onSuccess: () => void, user: any, users: AppUser[] }) {
  const [title, setTitle] = useState('');
  const [pic, setPic] = useState(user?.name || user?.email || '');
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  
  const [phases, setPhases] = useState([{ 
    id: createPersistentId(),
    custom_id: `#PH-${Math.random().toString(36).substring(2, 6).toUpperCase()}`,
    title: '', 
    assignee: user?.name || user?.email || '',
    start_date: '',
    end_date: '',
    duration_hours: 0,
    man_hours: 0,
    subtasks: [{ 
      id: createPersistentId(), 
      custom_id: `#TS-${Math.random().toString(36).substring(2, 6).toUpperCase()}`,
      title: '', 
      assignee: user?.name || user?.email || '',
      start_date: '',
      end_date: '',
      duration_hours: 0,
      man_hours: 0
    }]
  }]);
  const [loading, setLoading] = useState(false);

  // Auto-Sync Level 1 to Project Header (Truly Automatic Master Dates)
  useEffect(() => {
    const validFromDates = phases.map(p => {
      const dates = [p.start_date, ...(p.subtasks?.map(s => s.start_date) || [])].filter(Boolean);
      return dates.length > 0 ? dates.reduce((a, b) => a < b ? a : b) : '';
    }).filter(Boolean);

    const validToDates = phases.map(p => {
      const dates = [p.end_date, ...(p.subtasks?.map(s => s.end_date) || [])].filter(Boolean);
      return dates.length > 0 ? dates.reduce((a, b) => a > b ? a : b) : '';
    }).filter(Boolean);

    if (validFromDates.length > 0) {
       const minDate = validFromDates.reduce((a, b) => a < b ? a : b);
       setStartDate(minDate);
    }
    if (validToDates.length > 0) {
       const maxDate = validToDates.reduce((a, b) => a > b ? a : b);
       setEndDate(maxDate);
    }
  }, [phases]);

  // Allocation Validation
  const allocationStats = useMemo(() => {
    return phases.map(phase => {
      const capacity = phase.man_hours || 0;
      const used = phase.subtasks.reduce((sum, sub) => sum + (sub.man_hours || 0), 0);
      return { 
        capacity, 
        used, 
        isOver: used > capacity, 
        isUnder: used < capacity,
        isPerfect: Math.abs(used - capacity) < 0.01 && capacity > 0,
        remaining: capacity - used
      };
    });
  }, [phases]);

  const totalManHours = useMemo(() => {
    return phases.reduce((acc, phase) => acc + (phase.man_hours || 0), 0);
  }, [phases]);

  const isAnyOverAllocated = allocationStats.some(s => s.isOver);

  const addPhase = () => {
    setPhases([...phases, { 
      id: createPersistentId(),
      custom_id: `#PH-${Math.random().toString(36).substr(2, 4).toUpperCase()}`,
      title: 'ENTER TASK NAME...', 
      assignee: pic || user?.name || user?.email || '',
      start_date: '',
      end_date: '',
      duration_hours: 0,
      man_hours: 0,
      subtasks: [{ 
        id: createPersistentId(), 
        custom_id: `#TS-${Math.random().toString(36).substr(2, 4).toUpperCase()}`,
        title: 'ENTER TASK NAME...', 
        assignee: pic || user?.name || user?.email || '',
        start_date: '',
        end_date: '',
        duration_hours: 0,
        man_hours: 0
      }]
    }]);
  };

  const addSubtask = (phaseIndex: number) => {
    setPhases(prev => {
      const next = [...prev];
      const phase = { ...next[phaseIndex] };
      phase.subtasks = [...phase.subtasks, { 
        id: createPersistentId(), 
        custom_id: `#TS-${Math.random().toString(36).substr(2, 4).toUpperCase()}`,
        title: 'ENTER TASK NAME...', 
        assignee: pic || user?.name || user?.email || '',
        start_date: '',
        end_date: '',
        duration_hours: 0,
        man_hours: 0
      }];
      next[phaseIndex] = phase;
      return next;
    });
  };

  const handleCreate = async () => {
    if (!title || isAnyOverAllocated) return;
    setLoading(true);
    try {
      const finalPic = pic || user?.name || user?.email || 'Administrator';
      const actorName = user?.name || 'Fachrul Wisnu Novianto';
      const actorEmail = user?.email || 'Administrator';

      // Calculate dynamic project start and end dates from phases and subtasks
      const allDates: number[] = [];
      phases.forEach(p => {
        if (p.start_date) allDates.push(new Date(`${p.start_date}T08:00:00`).getTime());
        if (p.end_date) allDates.push(new Date(`${p.end_date}T17:00:00`).getTime());
        p.subtasks?.forEach(s => {
          if (s.start_date) allDates.push(new Date(`${s.start_date}T08:00:00`).getTime());
          if (s.end_date) allDates.push(new Date(`${s.end_date}T17:00:00`).getTime());
        });
      });

      const pStart = allDates.length > 0 ? format(new Date(Math.min(...allDates)), 'yyyy-MM-dd') : null;
      const pEnd = allDates.length > 0 ? format(new Date(Math.max(...allDates)), 'yyyy-MM-dd') : null;
      
      const prj = await taskService.createProject({ 
        name: title, 
        status: ProjectStatus.ACTIVE,
        leader_email: user?.email || null,
        pic_name: finalPic,
        start_date: pStart || undefined,
        end_date: pEnd || undefined
      }, actorEmail);
      
      for (const phase of phases) {
        const h = 8;
        const m = 0;
        const mH = parseFloat(String(phase.man_hours ?? 0)) || 0;

        const phaseStartStr = phase.start_date || startDate || format(new Date(), 'yyyy-MM-dd');
        const phaseEndStr = phase.end_date || endDate || format(addDays(new Date(phaseStartStr), 7), 'yyyy-MM-dd');

        const phaseStart = new Date(`${phaseStartStr}T08:00:00`);
        const phaseEnd = new Date(`${phaseEndStr}T17:00:00`);
        
        const l1 = await taskService.createTask({
          title: phase.title || 'Untitled Phase',
          custom_id: (phase as any).custom_id,
          project_id: prj.id,
          assignee: (phase as any).assignee || finalPic,
          start_time: phaseStart.toISOString(),
          end_time: phaseEnd.toISOString(),
          start_hour: h,
          start_minute: m,
          duration_hours: mH,
          man_hours: mH,
          duration_minutes: 0,
          created_by_name: user?.name || actorName
        }, actorEmail);

        for (const sub of phase.subtasks) {
          const subH = 8;
          const subM = 0;
          const subMH = parseFloat(String(sub.man_hours ?? 0)) || 0;

          const subStartStr = sub.start_date || phase.start_date || startDate || format(new Date(), 'yyyy-MM-dd');
          const subEndStr = sub.end_date || sub.start_date || phase.end_date || endDate || format(addDays(new Date(subStartStr), 7), 'yyyy-MM-dd');

          const start = new Date(`${subStartStr}T08:00:00`);
          const end = new Date(`${subEndStr}T17:00:00`);
          
          await taskService.createTask({
            title: sub.title || 'Untitled Sub-task',
            custom_id: (sub as any).custom_id,
            project_id: prj.id,
            parent_id: l1.id,
            assignee: (sub as any).assignee || finalPic,
            start_time: start.toISOString(),
            end_time: end.toISOString(),
            start_hour: subH,
            start_minute: subM,
            duration_hours: subMH,
            man_hours: subMH,
            duration_minutes: 0,
            created_by_name: user?.name || actorName
          }, actorEmail);
        }
      }

      onSuccess();
      onClose();
    } catch (err) {
      console.error("Project Creation Wizard failed:", err);
      alert("Failed to create project infrastructure.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-slate-950/80 backdrop-blur-md z-[100] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="bg-slate-900 border border-slate-800 rounded-3xl w-full max-w-7xl max-h-[95vh] overflow-hidden shadow-[0_0_50px_rgba(79,70,229,0.2)] flex flex-col"
      >
        <div className="p-6 border-b border-slate-800 bg-indigo-600/5 flex justify-between items-center">
          <div>
            <h2 className="text-xl font-black text-white tracking-tighter uppercase italic">
              Project <span className="text-indigo-500">Wizard</span>
            </h2>
            <p className="text-[10px] text-slate-500 mt-1 font-bold uppercase tracking-widest">Multi-Level Batch Provisioning</p>
          </div>
          <button onClick={onClose} className="p-2 hover:bg-slate-800 rounded-full transition-colors">
            <Plus className="w-5 h-5 text-slate-500 rotate-45" />
          </button>
        </div>

        <div className="p-8 flex-1 overflow-x-auto overflow-y-auto space-y-8 scrollbar-hide">
          <div className="grid grid-cols-4 gap-8 max-w-6xl mx-auto">
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Project Meta</label>
              <LocalInput 
                autoFocus
                value={title}
                onChange={v => setTitle(v)}
                placeholder="Project Name..."
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Lead PIC</label>
              <LocalInput 
                value={pic}
                onChange={v => setPic(v)}
                placeholder="Lead PIC Name..."
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-3 text-slate-200 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Start Date</label>
              <input 
                type="date"
                value={startDate}
                onChange={e => setStartDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-3 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
            <div className="space-y-4">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">End Date</label>
              <input 
                type="date"
                value={endDate}
                onChange={e => setEndDate(e.target.value)}
                className="w-full bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-3 outline-none focus:border-indigo-500 transition-colors font-bold"
              />
            </div>
          </div>

          <div className="space-y-6 min-w-[1100px]">
            <div className="flex justify-between items-center px-1">
              <label className="text-[11px] font-black text-indigo-500 uppercase tracking-widest">Work Breakdown Structure (WBS)</label>
              <button 
                onClick={addPhase}
                className="text-[9px] font-black uppercase text-indigo-400 hover:text-indigo-300 transition-colors bg-indigo-500/10 px-4 py-2 rounded-lg border border-indigo-500/20"
              >
                + Add Phase (L1)
              </button>
            </div>

            {/* Header Labels */}
            <div className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] gap-4 px-6 mb-[-16px]">
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest pl-2">Phase / Task Name</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">From Date</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">To Date</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Man-Hours</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">PIC</label>
              <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Actions</label>
            </div>

            {phases.map((phase, pIdx) => {
              const stats = allocationStats[pIdx];
              return (
                <div key={getSafeKey(phase, pIdx, 'wizard-phase')} className="bg-slate-950/40 border border-slate-800/60 rounded-2xl p-6 space-y-4 relative overflow-hidden group">
                  <div className="absolute top-0 left-0 w-1.5 h-full bg-indigo-600/30 group-hover:bg-indigo-500 transition-colors" />
                  
                  {/* L1 Header & Inputs */}
                  <div className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] items-center gap-4 bg-slate-900/50 p-4 rounded-xl border border-slate-800/50 shadow-sm transition-all hover:bg-slate-900/70">
                    <div className="flex items-center gap-2">
                      <span className="shrink-0 px-2 py-0.5 rounded-full bg-slate-800 border border-indigo-500/30 text-[8px] font-black text-indigo-400 font-mono tracking-wider shadow-inner">
                        {phase.custom_id}
                      </span>
                      <LocalInput 
                        value={phase.title}
                        onChange={(v) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].title = v;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-950 border border-slate-800 rounded-lg px-4 py-2.5 text-sm text-white focus:border-indigo-500 outline-none transition-all font-bold"
                        placeholder="Phase Title"
                      />
                    </div>

                    <div className="flex flex-col gap-1">
                      <input 
                        type="date"
                        value={phase.start_date}
                        min={startDate}
                        max={endDate || undefined}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].start_date = e.target.value;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-[10px] text-white focus:border-indigo-500 outline-none transition-all"
                      />
                    </div>

                    <div className="flex flex-col gap-1">
                      <input 
                        type="date"
                        value={phase.end_date}
                        min={phase.start_date || startDate}
                        max={endDate || undefined}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].end_date = e.target.value;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-[10px] text-white focus:border-indigo-500 outline-none transition-all"
                      />
                    </div>

                    <div className="flex flex-col items-center gap-1">
                      <input 
                        type="number" min="0" step="0.5"
                        value={phase.man_hours}
                        onChange={(e) => {
                          const newPhases = [...phases];
                          newPhases[pIdx].man_hours = Math.max(0, parseFloat(e.target.value) || 0);
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-900 border border-slate-700 rounded-lg px-2 py-2 text-xs text-indigo-400 text-center font-black outline-none focus:border-indigo-500"
                      />
                      <span className="text-[8px] font-black text-slate-600 uppercase tracking-tight h-3">
                         {phase.man_hours > 0 ? `≈ ${formatWorkday(phase.man_hours)}` : ''}
                      </span>
                    </div>

                    <div className="flex flex-col">
                      <LocalInput 
                        value={(phase as any).assignee}
                        onChange={(v) => {
                          const newPhases = [...phases];
                          (newPhases[pIdx] as any).assignee = v;
                          setPhases(newPhases);
                        }}
                        className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-[10px] text-slate-300 text-center outline-none focus:border-indigo-500"
                        placeholder="PIC Name"
                      />
                    </div>

                    <div className="flex justify-center gap-3">
                      <button 
                        onClick={() => addSubtask(pIdx)}
                        className="p-2 hover:bg-indigo-500/10 text-indigo-500/40 hover:text-indigo-400 rounded-lg transition-all"
                        title="Add Breakdown"
                      >
                        <Plus className="w-5 h-5" />
                      </button>
                      <button 
                        onClick={() => {
                          const newPhases = [...phases];
                          newPhases.splice(pIdx, 1);
                          setPhases(newPhases);
                        }}
                        className="p-2 hover:bg-rose-500/10 text-slate-700 hover:text-rose-500 rounded-lg transition-all"
                        title="Delete Phase"
                      >
                        <Trash2 className="w-5 h-5" />
                      </button>
                    </div>
                  </div>

                  {/* Allocation Validation Message */}
                  <div className="px-4">
                    {stats.isOver && (
                      <p key={`over-${phase.id}`} className="text-[10px] font-black text-rose-500 uppercase tracking-widest flex items-center gap-2 bg-rose-500/10 py-2 px-3 rounded-lg border border-rose-500/20">
                        <AlertTriangle className="w-3 h-3" />
                        ⚠️ Over-Allocated: Breakdown melebihi durasi Task Utama! (Kekurangan {formatWorkday(Math.abs(stats.remaining))})
                      </p>
                    )}
                    {stats.isUnder && (
                      <p key={`under-${phase.id}`} className="text-[10px] font-black text-amber-500 uppercase tracking-widest flex items-center gap-2 bg-amber-500/10 py-2 px-3 rounded-lg border border-amber-500/20">
                        <AlertTriangle className="w-3 h-3" />
                        ⚠️ Under-Allocated: Sisa waktu {formatWorkday(stats.remaining)} belum dialokasikan
                      </p>
                    )}
                    {stats.isPerfect && (
                      <p key={`perfect-${phase.id}`} className="text-[10px] font-black text-emerald-500 uppercase tracking-widest flex items-center gap-2 bg-emerald-500/5 py-2 px-3 rounded-lg border border-emerald-500/20">
                        <Plus className="w-3 h-3 text-emerald-500" />
                        ✅ Alokasi Waktu Pas
                      </p>
                    )}
                    {stats.capacity === 0 && (
                       <p key={`no-duration-${phase.id}`} className="text-[9px] font-bold text-slate-600 italic uppercase">Tentukan durasi untuk menghitung validasi alokasi.</p>
                    )}
                  </div>

                  <div className="space-y-3">
                    <div className="flex justify-between items-center px-4">
                      <p className="text-[9px] font-black text-slate-400 uppercase tracking-[0.2em] flex items-center gap-2">
                        <span className="w-1.5 h-1.5 rounded-full bg-indigo-500/50" />
                        Execution Breakdown
                      </p>
                    </div>

                    <div className="space-y-1.5">
                      {phase.subtasks.map((sub, sIdx) => {
                        return (
                          <div key={getSafeKey(sub, sIdx, 'wizard-subtask')} className="grid grid-cols-[1fr_130px_130px_80px_150px_80px] items-center gap-4 group/sub hover:bg-slate-900/30 p-2 rounded-lg transition-colors mx-2">
                          {/* Col 1 with Indentation Built-in */}
                          <div className="flex items-center gap-3 pl-8">
                            <span className="text-indigo-500/20 font-black text-xl select-none group-hover/sub:text-indigo-500/40">↳</span>
                            <span className="shrink-0 px-2 py-0.5 rounded-full bg-slate-950 border border-indigo-500/20 text-[8px] font-bold text-indigo-500/60 font-mono tracking-wider">
                              {(sub as any).custom_id}
                            </span>
                            <LocalInput 
                              value={sub.title}
                              placeholder="Activity description..."
                              onChange={(v) => {
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].title = v;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-transparent border-b border-white/5 text-xs text-slate-400 py-1.5 outline-none focus:border-indigo-500/50 transition-all font-medium"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="date"
                              value={sub.start_date}
                              min={phase.start_date}
                              max={phase.end_date || undefined}
                              onChange={(e) => {
                                const val = e.target.value;
                                if (phase.start_date && val < phase.start_date) {
                                  alert("Tanggal mulai child task tidak boleh mendahului tanggal pada task!");
                                  return;
                                }
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].start_date = val;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-lg px-1 py-1.5 text-[9px] text-white focus:border-indigo-500/50 outline-none transition-all"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="date"
                              value={sub.end_date}
                              min={sub.start_date || phase.start_date}
                              max={phase.end_date || undefined}
                              onChange={(e) => {
                                const val = e.target.value;
                                if (phase.end_date && val > phase.end_date) {
                                  alert("Tanggal selesai child task tidak boleh melebihi task utama!");
                                  return;
                                }
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].end_date = val;
                                setPhases(newPhases);
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-lg px-1 py-1.5 text-[9px] text-white focus:border-indigo-500/50 outline-none transition-all"
                            />
                          </div>

                          <div className="flex justify-center">
                            <input 
                              type="number" min="0" step="0.5"
                              value={sub.man_hours}
                              onChange={(e) => {
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks[sIdx].man_hours = Math.max(0, parseFloat(e.target.value) || 0);
                                setPhases(newPhases);
                              }}
                              className="w-16 bg-slate-900 border border-slate-700 rounded-md px-1 py-1 text-[10px] text-indigo-400/80 text-center outline-none focus:border-indigo-500/50 font-black"
                            />
                          </div>

                          <div className="flex justify-center">
                            <LocalInput 
                              value={(sub as any).assignee}
                              onChange={(v) => {
                                const newPhases = [...phases];
                                (newPhases[pIdx].subtasks[sIdx] as any).assignee = v;
                                setPhases(newPhases);
                              }}
                              placeholder="Assignee"
                              className="w-full bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-500 text-center outline-none focus:border-indigo-500/50"
                            />
                          </div>

                          <div className="flex justify-center">
                            <button 
                              onClick={() => {
                                const newPhases = [...phases];
                                newPhases[pIdx].subtasks.splice(sIdx, 1);
                                setPhases(newPhases);
                              }}
                              className="p-1.5 opacity-0 group-hover/sub:opacity-100 hover:bg-rose-500/10 text-slate-700 hover:text-rose-500 rounded transition-all"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </div>
                        </div>
                      );
                    })}
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        <div className="p-8 border-t border-slate-800 bg-slate-950/50 flex justify-between items-center px-10">
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-[0.2em] flex items-center gap-6">
            <span key="wizard-stat-phases" className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 rounded-full bg-indigo-500 animate-pulse outline outline-offset-2 outline-indigo-500/20" />
              {phases.length} Phases
            </span>
            <span key="wizard-stat-hours" className="flex items-center gap-2">
               <span className="w-2.5 h-2.5 rounded-full bg-slate-700" />
               <span className="text-indigo-400 font-black">Total Man-Hours: {(totalManHours || 0).toFixed(1)} Jam</span>
               <span className="text-slate-600 font-medium ml-1">{totalManHours > 0 ? `≈ ${formatWorkday(totalManHours)}` : ''}</span>
            </span>
            {isAnyOverAllocated && (
               <span key="wizard-stat-error" className="text-rose-500 font-black animate-pulse">⚠️ Resolusi Over-Alokasi diperlukan</span>
            )}
          </p>
          <div className="flex gap-4">
            <button 
              disabled={loading}
              onClick={onClose}
              className="px-8 py-3 bg-slate-800 hover:bg-slate-700 text-slate-500 font-black uppercase text-[10px] tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              disabled={loading || !title || isAnyOverAllocated}
              onClick={handleCreate}
              className="bg-indigo-600 hover:bg-indigo-500 disabled:opacity-30 disabled:grayscale text-white px-10 py-4 rounded-xl text-xs font-black uppercase tracking-widest transition-all shadow-[0_10px_30px_rgba(79,70,229,0.3)] hover:shadow-[0_15px_40px_rgba(79,70,229,0.4)] flex items-center gap-3 disabled:cursor-not-allowed"
            >
              {loading ? (
                <div className="w-5 h-5 border-2 border-white/20 border-t-white rounded-full animate-spin" />
              ) : (
                <>
                  <Rocket className="w-5 h-5" />
                  Spawn Infrastructure
                </>
              )}
            </button>
          </div>
        </div>
      </motion.div>
    </div>
  );
}

function BatchManualEntryModal({ onClose, onSuccess, users }: { onClose: () => void, onSuccess: () => void, users: AppUser[] }) {
  const [rows, setRows] = useState([{ id: crypto.randomUUID(), pic_name: '', status: 'WFO', schedule_date: format(new Date(), 'yyyy-MM-dd') }]);
  const [pasteData, setPasteData] = useState('');
  const [loading, setLoading] = useState(false);

  const addRow = () => {
    const lastRow = rows[rows.length - 1];
    setRows([...rows, { 
      id: crypto.randomUUID(),
      pic_name: lastRow?.pic_name || '', 
      status: 'WFO', 
      schedule_date: lastRow?.schedule_date || format(new Date(), 'yyyy-MM-dd') 
    }]);
  };

  const handlePaste = () => {
    if (!pasteData.trim()) return;
    const lines = pasteData.trim().split('\n');
    const newRows = lines.map(line => {
      const parts = line.split('\t');
      if (parts.length >= 2) {
        const pic_name = parts[0].trim();
        const status = (parts[1]?.trim().toUpperCase() || 'WFO');
        let schedule_date = parts[2]?.trim() || format(new Date(), 'yyyy-MM-dd');
        
        // Basic date normalization if it looks like DD/MM/YYYY
        if (schedule_date.includes('/')) {
          const dateParts = schedule_date.split('/');
          if (dateParts.length === 3) {
            const [d, m, y] = dateParts;
            schedule_date = `${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`;
          }
        }
        
        return { id: crypto.randomUUID(), pic_name, status, schedule_date };
      }
      return null;
    }).filter(Boolean) as any[];

    if (newRows.length > 0) {
      setRows([...rows, ...newRows]);
      setPasteData('');
    }
  };

  const handleSave = async () => {
    setLoading(true);
    try {
      const payload = rows.filter(r => r.pic_name && r.schedule_date);
      if (payload.length === 0) return;
      
      // Chunk processing for safety
      for (let i = 0; i < payload.length; i += 50) {
        await taskService.upsertSchedules(payload.slice(i, i + 50));
      }
      
      onSuccess();
      onClose();
    } catch (err) {
      console.error('Batch upsert failed:', err);
      alert('Gagal menyimpan data massal.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 z-[110] flex items-center justify-center p-4">
      <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} onClick={onClose} className="absolute inset-0 bg-slate-950/90 backdrop-blur-md" />
      <motion.div 
        initial={{ opacity: 0, scale: 0.95, y: 30 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        className="relative w-full max-w-4xl bg-slate-900 border border-slate-800 rounded-[2.5rem] shadow-2xl overflow-hidden max-h-[90vh] flex flex-col"
      >
        <div className="p-8 border-b border-slate-800 flex items-center justify-between bg-slate-900/50 backdrop-blur-md">
          <div>
            <h2 className="text-3xl font-black text-white italic tracking-tighter uppercase leading-none">Batch Manual Entry</h2>
            <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em] mt-3 flex items-center gap-2">
              <Copy className="w-3 h-3 text-indigo-500" /> Massive Schedule Update
            </p>
          </div>
          <button onClick={onClose} className="p-3 hover:bg-slate-800 rounded-2xl text-slate-400 transition-all">
            <X className="w-6 h-6" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto p-8 space-y-8 custom-scrollbar">
          {/* Paste Area */}
          <div className="space-y-3">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest px-1">Paste from Excel (Tab-Separated: Name \t Status \t Date)</label>
            <div className="flex gap-4">
              <textarea 
                value={pasteData}
                onChange={(e) => setPasteData(e.target.value)}
                placeholder="Paste multi-line data here..."
                className="flex-1 bg-slate-950/50 border border-slate-800 rounded-2xl p-4 text-xs font-mono text-slate-300 min-h-[100px] outline-none focus:border-indigo-500/50 transition-all"
              />
              <button 
                onClick={handlePaste}
                className="px-6 bg-slate-800 hover:bg-slate-700 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all"
              >
                Parse Paste
              </button>
            </div>
          </div>

          {/* Manual Table */}
          <div className="space-y-4">
            <div className="flex items-center justify-between px-1">
              <h3 className="text-sm font-black text-white uppercase tracking-widest">Entry Items</h3>
              <button onClick={addRow} className="flex items-center gap-2 px-3 py-1.5 bg-indigo-500/10 hover:bg-indigo-500/20 text-indigo-400 text-[10px] font-black uppercase tracking-widest rounded-lg border border-indigo-500/20 transition-all">
                <Plus className="w-3 h-3" /> Tambah Baris
              </button>
            </div>

            <div className="space-y-2 max-h-[400px] overflow-y-auto pr-2 custom-scrollbar">
              {rows.map((row, idx) => (
                <div key={getSafeKey(row, idx, 'entry-row')} className="grid grid-cols-[1fr_150px_180px_50px] gap-3 items-center bg-slate-950/30 p-2 rounded-xl border border-white/[0.02]">
                  <div className="relative group">
                    <select 
                      value={row.pic_name}
                      onChange={(e) => {
                        const newRows = [...rows];
                        newRows[idx].pic_name = e.target.value;
                        setRows(newRows);
                      }}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none"
                    >
                      <option value="">Select PIC</option>
                      {users.map((u, i) => <option key={getSafeKey(u, i, 'user-opt')} value={u.name}>{u.name}</option>)}
                    </select>
                  </div>
                  <select 
                    value={row.status}
                    onChange={(e) => {
                      const newRows = [...rows];
                      newRows[idx].status = e.target.value;
                      setRows(newRows);
                    }}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none font-bold"
                  >
                    <option value="WFO">WFO</option>
                    <option value="WFH">WFH</option>
                    <option value="WFC">WFC</option>
                    <option value="LIBUR">LIBUR</option>
                  </select>
                  <input 
                    type="date"
                    value={row.schedule_date}
                    onChange={(e) => {
                      const newRows = [...rows];
                      newRows[idx].schedule_date = e.target.value;
                      setRows(newRows);
                    }}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-xs text-white focus:border-indigo-500/50 outline-none font-bold"
                  />
                  <button 
                    onClick={() => setRows(rows.filter((_, i) => i !== idx))}
                    className="text-slate-600 hover:text-rose-500 p-2"
                  >
                    <X className="w-4 h-4" />
                  </button>
                </div>
              ))}
            </div>
          </div>
        </div>

        <div className="p-8 border-t border-slate-800 bg-slate-900 flex items-center justify-end gap-4">
          <button onClick={onClose} className="px-6 py-3 bg-slate-800 text-slate-400 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all">Cancel</button>
          <button 
            onClick={handleSave}
            disabled={loading || rows.length === 0}
            className="px-10 py-3 bg-emerald-600 hover:bg-emerald-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-emerald-500/20 transition-all active:scale-95"
          >
            {loading ? 'Processing...' : 'Simpan Semua'}
          </button>
        </div>
      </motion.div>
    </div>
  );
}

// --- View Components ---

function DashboardStats({ tasks, projects }: { tasks: Task[], projects: Project[] }) {
  const stats = useMemo(() => {
    const safeProjects = projects || [];
    const safeTasks = tasks || [];
    const totalProjects = safeProjects.length;
    const totalLevel1 = safeTasks.filter(t => !t.parent_id);
    const totalTasks = totalLevel1.length;
    const totalChildTasks = safeTasks.filter(t => !!t.parent_id).length;
    const totalManHours = safeTasks
      .reduce((sum, t) => sum + (Number(t.man_hours) || 0), 0);
    
    return { totalProjects, totalTasks, totalChildTasks, totalManHours };
  }, [tasks, projects]);

  return (
    <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-8">
      {[
        { label: 'Total Project', value: stats.totalProjects, icon: FolderKanban, color: 'emerald' },
        { label: 'Total Task', value: stats.totalTasks, icon: Layers, color: 'sky' },
        { label: 'Total Child Task', value: stats.totalChildTasks, icon: Activity, color: 'indigo' },
        { label: 'Total Man Hour', value: Number(stats.totalManHours).toFixed(1), icon: Clock, color: 'amber' },
      ].map((stat, i) => (
        <div key={getSafeKey({id: stat.label}, i, 'stat')} className="bg-slate-900/60 backdrop-blur-xl border border-white/5 p-5 rounded-2xl relative overflow-hidden group hover:border-indigo-500/30 transition-all shadow-2xl">
          <div className="flex items-center justify-between relative z-10">
            <div>
              <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest mb-1">{stat.label}</p>
              <h3 className="text-2xl font-black text-white italic tracking-tighter">{stat.value}</h3>
            </div>
            <div className={cn(
              "w-10 h-10 rounded-xl flex items-center justify-center transition-transform group-hover:scale-110",
              stat.color === 'emerald' ? "bg-emerald-500/10 text-emerald-400" :
              stat.color === 'sky' ? "bg-sky-500/10 text-sky-400" :
              stat.color === 'indigo' ? "bg-indigo-500/10 text-indigo-400" :
              "bg-amber-500/10 text-amber-400"
            )}>
              <stat.icon className="w-5 h-5" />
            </div>
          </div>
          <div className={cn(
            "absolute bottom-0 left-0 h-1 transition-all",
            stat.color === 'emerald' ? "bg-emerald-500/20 group-hover:bg-emerald-500" :
            stat.color === 'sky' ? "bg-sky-500/20 group-hover:bg-sky-500" :
            stat.color === 'indigo' ? "bg-indigo-500/20 group-hover:bg-indigo-500" :
            "bg-amber-500/20 group-hover:bg-amber-500"
          )} style={{ width: '100%' }} />
        </div>
      ))}
    </div>
  );
}

function PortfolioDashboard({ user, projects, tasks, loading, onOpenProject, onDeleteProject, onUpdateProject, onCreateRequested, onReschedule }: { 
  user: any,
  projects: Project[], 
  tasks: Task[],
  loading: boolean,
  onOpenProject: (id: string) => void,
  onDeleteProject: (id: string) => void,
  onUpdateProject: (id: string, updates: Partial<Project>) => void,
  onCreateRequested: () => void,
  onReschedule: (p: Project) => void
}) {
  const [deleteId, setDeleteId] = useState<string | null>(null);
  const [notif, setNotif] = useState<string | null>(null);
  const safeProjects = projects || [];

  const handleExportProjects = () => {
    const exportData = safeProjects.map(p => ({
      'Project Name': p.name || 'Untitled',
      'Lead PIC': p.pic_name || p.leader_email || 'Unassigned',
      'Start Date': p.start_date || 'N/A',
      'End Date': p.end_date || 'N/A',
      'Status': p.status || 'Unknown'
    }));
    handleExcelExport(exportData, 'Projects_Summary');
  };

  return (
    <div className="space-y-6">
      <SuccessNotification show={!!notif} message={notif || ''} onClose={() => setNotif(null)} />
      
      <div className="flex items-center justify-between gap-4 mb-2">
        <div>
          <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">Om Dedy Portfolio</h2>
          <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] mt-1">Global Project Tracking System</p>
        </div>
        <div className="flex items-center gap-3">
          {user && (
            <button 
              onClick={handleExportProjects}
              className="bg-slate-800 hover:bg-slate-700 text-slate-300 px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2 border border-slate-700 transition-all active:scale-95"
            >
              <Download className="w-4 h-4" /> Export to Excel
            </button>
          )}
          {user && (
            <button 
              onClick={onCreateRequested}
              className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white px-4 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center gap-2 transition-all active:scale-95"
            >
              <Plus className="w-4 h-4" /> New Project
            </button>
          )}
        </div>
      </div>
      
      {!loading && <DashboardStats tasks={tasks} projects={projects} />}
      
      {loading ? (
        <div className="h-64 flex flex-col items-center justify-center gap-4 bg-slate-900/20 rounded-3xl border border-slate-800/50">
           <div className="w-8 h-8 border-2 border-indigo-500 border-t-transparent rounded-full animate-spin" />
           <p className="text-[10px] font-black uppercase tracking-[0.2em] text-slate-500">Retrieving Portfolio Data</p>
        </div>
      ) : (
        <>
          <ConfirmModal 
            isOpen={!!deleteId}
            onClose={() => setDeleteId(null)}
            onConfirm={() => {
              if (deleteId) {
                onDeleteProject(deleteId);
                setNotif('Project successfully decommissioned');
              }
            }}
            title="Archive Project?"
            description="Are you sure you want to remove this project? All associated tasks and audit trails will be archived."
          />

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {(() => {
              const safeProjects = projects || [];
              const uniqueProjects = Array.from(new Map(safeProjects.filter(p => !!p).map(p => [p.id, p])).values());
              if (uniqueProjects.length === 0) {
                return (
                  <div className="col-span-full flex flex-col items-center justify-center py-20 px-8 bg-slate-900/50 border border-slate-800 border-dashed rounded-3xl text-center space-y-4">
                    <div className="w-16 h-16 bg-slate-800/50 rounded-2xl flex items-center justify-center">
                      <FolderKanban className="w-8 h-8 text-slate-600" />
                    </div>
                    <div className="space-y-1">
                      <h3 className="text-white font-bold tracking-tight">No Projects Found</h3>
                      <p className="text-slate-500 text-xs uppercase tracking-widest font-black">Portfolio is currently empty</p>
                    </div>
                  </div>
                );
              }
              return uniqueProjects.map((p, i) => {
                const combinedKey = getSafeKey(p, i, 'portfolio-project');
                return (
                  <motion.div
                    key={combinedKey}
                    whileHover={{ y: -4 }}
                    className="bg-slate-900 border border-slate-800 rounded-2xl p-6 cursor-pointer hover:border-indigo-500/50 transition-all group relative overflow-hidden"
                  >
              <div className="absolute top-4 right-4 flex gap-2 opacity-0 group-hover:opacity-100 transition-opacity z-10">
                {user && (
                  <button 
                    onClick={(e) => {
                      e.stopPropagation();
                      onReschedule(p);
                    }}
                    className="p-2 bg-slate-800 hover:bg-indigo-500/20 text-slate-400 hover:text-indigo-400 rounded-lg transition-colors border border-slate-700 hover:border-indigo-500/50"
                    title="Reschedule Project"
                  >
                    <Calendar className="w-3.5 h-3.5" />
                  </button>
                )}
                {user && (
                  <button 
                    onClick={(e) => {
                      e.stopPropagation();
                      setDeleteId(p.id);
                    }}
                    className="p-2 bg-slate-800 hover:bg-rose-500/20 text-slate-400 hover:text-rose-400 rounded-lg transition-colors border border-slate-700 hover:border-rose-500/50"
                    title="Archive Project"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                )}
              </div>

                  <div onClick={() => onOpenProject(p.id)}>
                <div className="flex justify-between items-start mb-6">
                  <div className="w-12 h-12 bg-slate-800 rounded-xl flex items-center justify-center group-hover:bg-indigo-500/10 transition-colors">
                    <FolderKanban className="w-6 h-6 text-slate-400 group-hover:text-indigo-400" />
                  </div>
                  <div onClick={(e) => e.stopPropagation()}>
                    <ProjectStatusSelector status={p.status} onUpdate={(v) => onUpdateProject(p.id, { status: v })} disabled={!user} />
                  </div>
                </div>
                <input 
                  defaultValue={p.name || ''}
                  onClick={(e) => e.stopPropagation()}
                  onBlur={(e) => {
                    if (e.target.value !== p.name) {
                      onUpdateProject(p.id, { name: e.target.value });
                      setNotif('Project title synchronized');
                    }
                  }}
                  className="text-lg font-bold text-white mb-2 bg-transparent border-none outline-none focus:text-indigo-400 w-full"
                  placeholder="Untitled Project"
                  disabled={!user}
                />
                
                <div className="flex flex-col gap-1.5 mb-2">
                  <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                    <UserIcon className="w-3 h-3 text-indigo-500" />
                    <span>PIC: <span className="text-slate-300">{p.pic_name || p.leader_email || 'Unassigned'}</span></span>
                  </div>
                  <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                    <Clock className="w-3 h-3 text-slate-600" />
                    <span>Timeline: <span className="text-slate-300 font-mono tracking-tighter">
                      {p.start_date ? format(new Date(p.start_date), 'MMM dd') : 'Belum Set'} - {p.end_date ? format(new Date(p.end_date), 'MMM dd, yyyy') : 'Belum Set'}
                    </span></span>
                  </div>
                </div>

                <div className="flex items-center justify-between text-[10px] font-black uppercase tracking-widest text-slate-600 mt-4 pt-4 border-t border-slate-800 group-hover:text-indigo-400 transition-colors">
                  <span>View Full Timeline</span>
                  <ExternalLink className="w-4 h-4 opacity-0 group-hover:opacity-100 transition-all transform translate-x-2 group-hover:translate-x-0 text-indigo-400" />
                </div>
              </div>
            </motion.div>
                  );
                });
              })()}
        
        {user && (
          <button 
            onClick={onCreateRequested}
          className="border-2 border-dashed border-slate-800 rounded-2xl flex flex-col items-center justify-center p-6 text-slate-500 hover:text-indigo-400 hover:border-indigo-500/50 transition-all gap-3 bg-slate-900/10 group"
        >
          <div className="w-12 h-12 rounded-2xl bg-slate-800 group-hover:bg-gradient-to-br group-hover:from-indigo-600 group-hover:to-violet-600 flex items-center justify-center transition-all group-hover:shadow-[0_0_20px_rgba(79,70,229,0.3)]">
            <Plus className="w-8 h-8 group-hover:rotate-90 group-hover:text-white transition-all" />
          </div>
          <span className="text-xs font-black uppercase tracking-widest group-hover:text-indigo-400 transition-colors">Create New Project</span>
        </button>
        )}
      </div>
    </>
  )}
</div>
  );
}

const SCHEDULE_STATUS_COLORS: Record<string, string> = {
  'WFO': 'text-emerald-400 bg-emerald-500/10 border-emerald-500/20 shadow-[0_0_8px_rgba(16,185,129,0.1)]',
  'WFH': 'text-indigo-400 bg-indigo-500/10 border-indigo-500/20 shadow-[0_0_8px_rgba(99,102,241,0.1)]',
  'WFC': 'text-amber-400 bg-amber-500/10 border-amber-500/20 shadow-[0_0_8px_rgba(245,158,11,0.1)]',
  'A2': 'text-amber-400 bg-amber-500/10 border-amber-500/20 shadow-[0_0_8px_rgba(245,158,11,0.1)]',
  'LIBUR': 'text-rose-400 bg-rose-500/10 border-rose-500/20 shadow-[0_0_8px_rgba(244,63,94,0.1)]',
  'LATE': 'text-rose-400 bg-rose-500/10 border-rose-500/20 shadow-[0_0_8px_rgba(244,63,94,0.1)]',
};

function OmDedySchedule({ user, users, setActiveView, isAdmin, isSuperadmin, setOwnershipModal, onNotif }: { 
  user: any, 
  users: AppUser[], 
  setActiveView: (view: AppView) => void,
  isAdmin: boolean,
  isSuperadmin: boolean,
  setOwnershipModal: React.Dispatch<React.SetStateAction<any>>,
  onNotif?: (msg: string | null) => void
}) {
  const [currentMonth, setCurrentMonth] = useState(new Date());
  const [schedules, setSchedules] = useState<Schedule[]>([]);
  const [loading, setLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);
  const [showBatchModal, setShowBatchModal] = useState(false);
  const [editingCell, setEditingCell] = useState<{ pic: string, date: string } | null>(null);
  const [rescheduleRequests, setRescheduleRequests] = useState<any[]>([]);
  const [showRescheduleRequests, setShowRescheduleRequests] = useState(false);
  const [requestModal, setRequestModal] = useState<{ 
    pic: string, 
    date: string, 
    currentStatus: string, 
    newStatus: string,
    isSwap?: boolean,
    swapDate?: string,
    swapStatus?: string,
    swapCurrentStatus?: string
  } | null>(null);

  const currentUserProfile = useMemo(() => {
    return users.find(u => u.email?.toLowerCase() === user?.email?.toLowerCase());
  }, [users, user]);

  const userAccess = (currentUserProfile?.access_level || '').toLowerCase().trim();
  const fullName = currentUserProfile?.name || '';

  const monthStart = startOfMonth(currentMonth);
  const monthEnd = endOfMonth(currentMonth);
  const days = eachDayOfInterval({ start: monthStart, end: monthEnd });

  const stats = useMemo(() => {
    const counts = { WFO: 0, WFH: 0, CUTI_LIBUR: 0, WFC: 0 };
    schedules.forEach(s => {
      // Only count if it's in the current month's days
      const dateKey = s.schedule_date; // already YYYY-MM-DD
      const date = new Date(dateKey);
      if (date >= monthStart && date <= monthEnd) {
        const st = s.status.toUpperCase();
        if (st === 'WFO') counts.WFO++;
        else if (st === 'WFH') counts.WFH++;
        else if (st === 'WFC' || st === 'A2') counts.WFC++;
        else if (st === 'CUTI' || st === 'LIBUR') counts.CUTI_LIBUR++;
      }
    });
    return counts;
  }, [schedules, monthStart, monthEnd]);

  const picStats = useMemo(() => {
    const statsMap: Record<string, { WFO: number, WFH: number, TOTAL: number }> = {};
    const list = Array.from(new Set(schedules.map(s => s.pic_name))).sort();
    
    list.forEach(p => statsMap[p] = { WFO: 0, WFH: 0, TOTAL: 0 });
    
    schedules.forEach(s => {
      const d = new Date(s.schedule_date);
      if (d >= monthStart && d <= monthEnd && statsMap[s.pic_name]) {
        const st = (s.status || '').toUpperCase();
        if (st === 'WFO') statsMap[s.pic_name].WFO++;
        else if (st === 'WFH') statsMap[s.pic_name].WFH++;
        statsMap[s.pic_name].TOTAL = statsMap[s.pic_name].WFO + statsMap[s.pic_name].WFH;
      }
    });
    return statsMap;
  }, [schedules, monthStart, monthEnd]);

  const fetchSchedules = async () => {
    try {
      setLoading(true);
      const data = await taskService.getSchedules(currentMonth);
      setSchedules(data);
      if (user?.access_level && user.access_level !== 'PIC') {
        const reqs = await taskService.getRescheduleRequests();
        setRescheduleRequests(reqs);
      }
    } catch (err) {
      console.error('Failed to fetch schedules:', err);
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchSchedules();
    
    // Subscribe to schedule changes
    const scheduleChannel = supabase.channel('schedules-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'schedules' }, () => {
        fetchSchedules();
      })
      .subscribe();

    const requestChannel = supabase.channel('requests-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'reschedule_requests' }, () => {
        fetchSchedules();
      })
      .subscribe();
      
    return () => { 
      supabase.removeChannel(scheduleChannel); 
      supabase.removeChannel(requestChannel);
    };
  }, [currentMonth]);

  const pics = useMemo(() => {
    // Collect all unique PICs from schedules
    const list = Array.from(new Set(schedules.map(s => s.pic_name))).sort();
    return list;
  }, [schedules]);

  const scheduleGrid = useMemo(() => {
    const grid: Record<string, Record<string, string>> = {};
    schedules.forEach(s => {
      const dateKey = format(new Date(s.schedule_date), 'yyyy-MM-dd');
      if (!grid[s.pic_name]) grid[s.pic_name] = {};
      grid[s.pic_name][dateKey] = s.status.toUpperCase();
    });
    return grid;
  }, [schedules]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImporting(true);
    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        
        const rawJsonAoA: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1 });
        const jsonData: any[] = XLSX.utils.sheet_to_json(ws);

        let upsertPayload: Partial<Schedule>[] = [];

        // 1. Detect if it's a FLAT format (name, status, date)
        const firstRow = jsonData[0] || {};
        const keys = Object.keys(firstRow).map(k => k.toLowerCase());
        const flatMapping = {
          name: Object.keys(firstRow).find(k => ['name', 'pic'].includes(k.toLowerCase()) || k.toLowerCase().includes('nama')),
          status: Object.keys(firstRow).find(k => k.toLowerCase() === 'status'),
          date: Object.keys(firstRow).find(k => k.toLowerCase() === 'date' || k.toLowerCase() === 'tanggal')
        };

        if (flatMapping.name && flatMapping.status && flatMapping.date) {
           console.log('Flat Format Detected');
           upsertPayload = jsonData.map(row => {
             const pic_name = String(row[flatMapping.name!] || '').trim();
             const status = String(row[flatMapping.status!] || '').trim().toUpperCase();
             let schedule_date = '';
             
             const rawDate = row[flatMapping.date!];
             if (typeof rawDate === 'number') {
               const d = XLSX.SSF.parse_date_code(rawDate);
               schedule_date = format(new Date(d.y, d.m - 1, d.d), 'yyyy-MM-dd');
             } else {
               const d = new Date(rawDate);
               if (!isNaN(d.getTime())) schedule_date = format(d, 'yyyy-MM-dd');
             }
             
             return { pic_name, status, schedule_date };
           }).filter(r => r.pic_name && r.status && r.schedule_date);

        } else {
          // 2. MATRIX HEURISTIC (existing robust logic)
          console.log('Matrix Format Fallback');
          let headerRowIndex = -1;
          let nameColumnIndex = -1;
          let dateHeaders: any[] = [];

          for (let i = 0; i < Math.min(rawJsonAoA.length, 100); i++) {
            const row = rawJsonAoA[i];
            if (!row || !Array.isArray(row)) continue;
            
            const picIdx = row.findIndex(c => {
              const s = String(c || '').toLowerCase().trim();
              return s.includes('nama') || s.includes('pic') || s === 'name' || s === 'karyawan';
            });
            
            if (picIdx !== -1) {
              headerRowIndex = i;
              nameColumnIndex = picIdx;
              dateHeaders = row.slice(nameColumnIndex + 1);
              break;
            }
          }

          if (headerRowIndex !== -1) {
            const normalizeDate = (raw: any): string | null => {
              if (!raw) return null;
              const str = String(raw).trim().toLowerCase();
              if (!isNaN(Number(str)) && Number(str) > 10000) {
                const d = XLSX.SSF.parse_date_code(Number(str));
                return format(new Date(d.y, d.m - 1, d.d), 'yyyy-MM-dd');
              }
              const monthMap: Record<string, string> = { 
                'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'mei': '05', 
                'jun': '06', 'jul': '07', 'aug': '08', 'agu': '08', 'ags': '08', 'sep': '09', 
                'oct': '10', 'okt': '10', 'nov': '11', 'dec': '12', 'des': '12' 
              };
              const match = str.match(/(\d{1,2})[\s\-\/]([a-z]{3}|\d{1,2})/);
              if (match) {
                const day = match[1].padStart(2, '0');
                const monthPart = match[2];
                let m = /^\d+$/.test(monthPart) ? monthPart.padStart(2, '0') : (monthMap[monthPart.substring(0, 3)] || '');
                if (m && day) return `${new Date().getFullYear()}-${m}-${day}`;
              }
              try {
                const d = new Date(str);
                if (!isNaN(d.getTime())) return format(d, 'yyyy-MM-dd');
              } catch (e) {}
              return null;
            };

            const parsedDates = dateHeaders.map(d => normalizeDate(d));
            const dataRows = rawJsonAoA.slice(headerRowIndex + 1);

            dataRows.forEach(row => {
              const picName = String(row[nameColumnIndex] || '').trim();
              if (!picName || picName.length < 2 || /^\d+$/.test(picName)) return;

              parsedDates.forEach((scheduleDate, dateIdx) => {
                if (!scheduleDate) return;
                const rawValue = String(row[nameColumnIndex + 1 + dateIdx] || '').trim().toUpperCase();
                let status = 'WFH';
                if (rawValue === 'WFO') status = 'WFO';
                else if (['LIBUR', 'OFF', 'CUTI'].includes(rawValue)) status = 'LIBUR';
                else if (['WFC', 'A2'].includes(rawValue)) status = 'WFC';
                else if (rawValue === 'WFH') status = 'WFH';
                
                upsertPayload.push({ pic_name: picName, status, schedule_date: scheduleDate });
              });
            });
          }
        }

        if (upsertPayload.length > 0) {
          const chunkSize = 100;
          for (let i = 0; i < upsertPayload.length; i += chunkSize) {
            await taskService.upsertSchedules(upsertPayload.slice(i, i + chunkSize));
          }
          alert(`Berhasil mengimpor ${upsertPayload.length} data jadwal.`);
          fetchSchedules();
        } else {
          alert('Format data tidak dikenali (Gunakan CSV Flat atau Matrix)');
        }
      } catch (err) {
        console.error('Import Error:', err);
        alert(err instanceof Error ? err.message : 'Gagal mengimpor file.');
      } finally {
        setIsImporting(false);
        if (e.target) e.target.value = '';
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleStatusUpdate = async (pic: string, date: string, status: string) => {
    if (!user) return; // RBAC: Public users cannot edit

    const currentStatus = scheduleGrid[pic]?.[date] || '';
    
    // Admin/Superadmin can update directly without approval
    if (isAdmin || isSuperadmin) {
       try {
         await taskService.upsertSchedules([{ pic_name: pic, schedule_date: date, status: status }]);
         fetchSchedules();
         onNotif?.("Jadwal updated!");
       } catch (err) {
         console.error("Failed to update schedule:", err);
       }
    } else {
      // Regular PICs (including owners) must go through the Approval Gate
      setRequestModal({ pic, date, currentStatus, newStatus: status });
    }
    setEditingCell(null);
  };

  const handleDownloadTemplate = () => {
    const data = [
      ["Name", "02 Jan 2026", "03 Jan 2026", "04 Jan 2026", "05 Jan 2026"],
      ["Fachrul Wisnu Novianto", "WFO", "WFH", "LIBUR", ""]
    ];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template_Matrix");
    XLSX.writeFile(wb, "Template_Schedule_Sakti_Matrix.xlsx");
  };

  const handleExport = () => {
    // Generate grid for Excel
    const header = ['PIC', ...days.map(d => format(d, 'dd/MM/yyyy'))];
    const dataRows = pics.map(pic => [
      pic,
      ...days.map(d => scheduleGrid[pic]?.[format(d, 'yyyy-MM-dd')] || '-')
    ]);

    const ws = XLSX.utils.aoa_to_sheet([header, ...dataRows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Schedule");
    XLSX.writeFile(wb, `OMDEDY_Schedule_${format(currentMonth, 'MMM_yyyy')}.xlsx`);
  };

  return (
    <div className="space-y-6 animate-in fade-in slide-in-from-bottom-4 duration-700">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-6">
        <div className="flex items-center gap-4">
          <div className="w-14 h-14 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
            <Calendar className="w-7 h-7 text-indigo-400" />
          </div>
          <div>
            <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">Om Dedy Schedule</h2>
            <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] leading-none mt-1">Resource Capacity Monitoring</p>
          </div>
        </div>

        <div className="flex items-center gap-4 bg-slate-900/80 p-1.5 rounded-2xl border border-slate-800 shadow-xl">
           <button onClick={() => setCurrentMonth(addMonths(currentMonth, -1))} className="p-2.5 hover:bg-slate-800 rounded-xl text-slate-400 transition-all hover:text-white">
             <ChevronLeft className="w-5 h-5" />
           </button>
           <div className="px-6 py-1 border-x border-slate-800">
             <span className="text-lg font-black text-white italic tracking-widest uppercase">{format(currentMonth, 'MMMM yyyy')}</span>
           </div>
           <button onClick={() => setCurrentMonth(addMonths(currentMonth, 1))} className="p-2.5 hover:bg-slate-800 rounded-xl text-slate-400 transition-all hover:text-white">
             <ChevronRight className="w-5 h-5" />
           </button>
        </div>

        <div className="flex items-center gap-3">
          {user && (user.access_level === 'Superadmin' || user.access_level === 'Admin') && (
            <>
              <button 
                onClick={() => setShowBatchModal(true)}
                className="flex items-center gap-2 px-5 py-2.5 bg-slate-800 hover:bg-slate-700 text-slate-200 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-slate-700"
              >
                <Plus className="w-4 h-4" />
                Input Massal
              </button>
              <button 
                onClick={() => setActiveView('RESCHEDULE')}
                className="flex items-center gap-2 px-5 py-2.5 bg-amber-500/10 hover:bg-amber-500/20 text-amber-500 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-amber-500/20 relative"
              >
                <History className="w-4 h-4" />
                Reschedule
                {rescheduleRequests.filter(r => r.status === 'Pending').length > 0 && (
                  <span className="absolute -top-1 -right-1 w-4 h-4 bg-rose-500 text-white text-[8px] rounded-full flex items-center justify-center animate-bounce shadow-lg ring-2 ring-slate-950">
                    {rescheduleRequests.filter(r => r.status === 'Pending').length}
                  </span>
                )}
              </button>
            </>
          )}
          {user && (
            <label className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-xl text-[10px] font-black uppercase tracking-widest cursor-pointer shadow-lg shadow-indigo-500/20 transition-all active:scale-95 group">
              <Upload className="w-4 h-4 group-hover:bounce" />
              {isImporting ? 'Parsing...' : 'Smart Import'}
              <input type="file" className="hidden" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} disabled={isImporting} />
            </label>
          )}
          {user && (
            <button onClick={handleExport} className="flex items-center gap-2 px-5 py-2.5 bg-black/20 hover:bg-black/30 text-slate-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 border border-white/5">
              <Download className="w-4 h-4" />
              Export Grid
            </button>
          )}
          {user && (
            <button 
              onClick={handleDownloadTemplate}
              className="flex items-center gap-2 px-5 py-2.5 bg-slate-900 border border-slate-800 hover:bg-slate-800 text-slate-400 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all active:scale-95 shadow-xl"
            >
              <History className="w-4 h-4" />
              Template Excel
            </button>
          )}
        </div>

        {showBatchModal && (
          <BatchManualEntryModal 
            users={users} 
            onClose={() => setShowBatchModal(false)} 
            onSuccess={() => {
              fetchSchedules();
            }} 
          />
        )}
      </div>

      {/* Monthly Statistics Summary */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-emerald-500/5 transition-all">
          <div className="w-12 h-12 bg-emerald-500/10 rounded-2xl flex items-center justify-center border border-emerald-500/20 shadow-lg shadow-emerald-500/5">
            < Rocket className="w-6 h-6 text-emerald-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Total WFO</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFO}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-indigo-500/5 transition-all">
          <div className="w-12 h-12 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
            <Activity className="w-6 h-6 text-indigo-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Total WFH</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFH}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-rose-500/5 transition-all">
          <div className="w-12 h-12 bg-rose-500/10 rounded-2xl flex items-center justify-center border border-rose-500/20 shadow-lg shadow-rose-500/5">
            <Filter className="w-6 h-6 text-rose-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Absence/Libur</p>
            <h4 className="text-2xl font-black text-white italic">{stats.CUTI_LIBUR}</h4>
          </div>
        </div>
        <div className="bg-slate-900/50 border border-slate-800 rounded-3xl p-5 flex items-center gap-4 group hover:bg-amber-500/5 transition-all">
          <div className="w-12 h-12 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20 shadow-lg shadow-amber-500/5">
            <LayoutGrid className="w-6 h-6 text-amber-400" />
          </div>
          <div>
            <p className="text-[9px] font-black text-slate-500 uppercase tracking-[0.2em] mb-1">Outstation/WFC</p>
            <h4 className="text-2xl font-black text-white italic">{stats.WFC}</h4>
          </div>
        </div>
      </div>

      <div className="bg-slate-900/30 border border-slate-800/80 rounded-[2.5rem] overflow-hidden flex flex-col backdrop-blur-xl shadow-2xl">
        <div className="overflow-auto relative min-h-[500px]">
          <table className="w-full border-collapse">
            <thead>
              <tr className="bg-slate-900/90 border-b border-slate-800 sticky top-0 z-40 backdrop-blur-md">
                <th className="sticky left-0 z-50 bg-slate-950 px-8 py-5 border-r border-slate-800 min-w-[240px] text-left">
                  <div className="flex items-center gap-3">
                    <div className="w-2 h-2 rounded-full bg-indigo-500 animate-pulse" />
                    <span className="text-[10px] font-black text-slate-500 uppercase tracking-[0.3em]">PIC RESOURCE</span>
                  </div>
                </th>
                {days.map((d, di) => (
                  <th key={getSafeKey({id: d.toISOString()}, di, 'omdedy-header')} className={cn(
                    "px-4 py-5 border-r border-slate-800/20 min-w-[70px] transition-all",
                    isToday(d) ? "bg-indigo-500/20 ring-1 ring-inset ring-indigo-500/30" : "hover:bg-white/[0.02]"
                  )}>
                    <div className="flex flex-col gap-1 items-center relative group/header">
                       <span className={cn(
                         "text-xl font-black italic tracking-tighter",
                         isToday(d) ? "text-indigo-400" : (HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] ? "text-rose-500" : (isWeekend(d) ? "text-rose-500/80" : "text-white"))
                       )}>
                         {format(d, 'dd')}
                       </span>
                       <span className={cn(
                         "text-[8px] font-black text-slate-500 uppercase tracking-widest",
                         CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')] && "text-amber-500"
                       )}>
                         {format(d, 'EEE')}
                       </span>
                       
                       {(HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] || CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')]) && (
                         <div className="absolute top-full mt-2 hidden group-hover/header:block z-[100] bg-slate-800 border border-slate-700 p-2 rounded shadow-xl text-[10px] font-bold text-white whitespace-nowrap pointer-events-none">
                           {HOLIDAYS_2026[format(d, 'yyyy-MM-dd')] || CUTI_BERSAMA_2026[format(d, 'yyyy-MM-dd')]}
                         </div>
                       )}
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {loading ? (
                <tr>
                  <td colSpan={days.length + 1} className="py-32">
                     <div className="flex flex-col items-center gap-4">
                        <Activity className="w-10 h-10 text-indigo-500 animate-pulse" />
                        <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">Synchronizing Capacity Grid...</p>
                     </div>
                  </td>
                </tr>
              ) : pics.length === 0 ? (
                <tr>
                   <td colSpan={days.length + 1} className="py-20 text-center">
                     <p className="text-[10px] font-black text-slate-500 uppercase tracking-widest italic">No PIC Data Available for this period</p>
                   </td>
                </tr>
              ) : pics.map((pic, pi) => (
                <tr key={getSafeKey({id: pic}, pi, 'resource-row')} className="group border-b border-white/[0.02] hover:bg-white/[0.01] transition-all">
                   <td className="sticky left-0 z-30 bg-slate-950/95 backdrop-blur-md px-6 py-4 border-r border-slate-800 transition-colors">
                     <div className="flex flex-col gap-1.5">
                       <span className="font-black text-sm italic text-slate-300 tracking-tighter group-hover:text-indigo-400">
                         {pic}
                       </span>
                       <div className="flex items-center gap-2">
                         <div className="flex items-center gap-1 bg-emerald-500/10 px-1.5 py-0.5 rounded-md border border-emerald-500/20">
                           <span className="text-[8px] font-black text-emerald-500">O: {picStats[pic]?.WFO || 0}</span>
                         </div>
                         <div className="flex items-center gap-1 bg-indigo-500/10 px-1.5 py-0.5 rounded-md border border-indigo-500/20">
                           <span className="text-[8px] font-black text-indigo-400">H: {picStats[pic]?.WFH || 0}</span>
                         </div>
                         <div className="flex items-center gap-1 bg-slate-800 px-1.5 py-0.5 rounded-md border border-slate-700">
                           <span className="text-[8px] font-black text-slate-400">T: {picStats[pic]?.TOTAL || 0}</span>
                         </div>
                       </div>
                     </div>
                   </td>
                   {days.map((d, di) => {
                     const dateKey = format(d, 'yyyy-MM-dd');
                     const holiday = HOLIDAYS_2026[dateKey];
                     const cuti = CUTI_BERSAMA_2026[dateKey];
                     const status = scheduleGrid[pic]?.[dateKey];
                     
                     const displayStatus = status || (holiday ? 'LIBUR' : null);
                     
                     return (
                       <td key={getSafeKey({id: `${pic}-${dateKey}`}, di, 'omdedy-cell')} className={cn(
                         "p-2 border-r border-slate-800/10 text-center relative group/cell",
                         isToday(d) && "bg-indigo-500/5",
                         holiday && "bg-rose-500/5",
                         cuti && "bg-amber-500/5"
                       )}>
                         <div className="flex justify-center items-center h-full min-h-[40px]">
                            {(holiday || cuti) && (
                               <div className="absolute bottom-full mb-2 hidden group-hover/cell:block z-[100] bg-slate-800 border border-slate-700 p-2 rounded shadow-xl text-[10px] font-bold text-white whitespace-nowrap pointer-events-none">
                                 {holiday || cuti}
                               </div>
                            )}
                            
                            {editingCell?.pic === pic && editingCell?.date === dateKey ? (
                              <div className="absolute inset-x-1 z-[110] bg-slate-950/95 backdrop-blur-md p-2 flex flex-col gap-1.5 shadow-2xl border border-indigo-500/50 rounded-2xl animate-in zoom-in-95 duration-200">
                                 {['WFO', 'WFH', 'WFC', 'LIBUR', 'A2', 'DELETE'].map((st, si) => (
                                   <button 
                                     key={getSafeKey({id: st}, si, 'status-opt')}
                                     onClick={() => {
                                       if (st === 'DELETE') {
                                          handleStatusUpdate(pic, dateKey, '');
                                       } else {
                                         handleStatusUpdate(pic, dateKey, st);
                                       }
                                     }}
                                     className={cn(
                                       "text-[9px] font-black px-3 py-1.5 rounded-xl grow hover:scale-105 active:scale-95 transition-all text-center border uppercase tracking-widest",
                                       st === 'DELETE' ? "bg-rose-500/20 text-rose-500 border-rose-500/30" : (SCHEDULE_STATUS_COLORS[st] || "bg-slate-800 text-slate-400 border-slate-700")
                                     )}
                                   >
                                     {st}
                                   </button>
                                 ))}
                                 <button onClick={() => setEditingCell(null)} className="text-[8px] font-black text-slate-500 bg-black/40 py-1.5 rounded-xl hover:text-white transition-colors">BATAL</button>
                              </div>
                            ) : (
                              displayStatus ? (
                               <motion.div 
                                 whileHover={{ scale: 1.1, zIndex: 10 }}
                                 onClick={() => {
                                   const isOwner = fullName === pic;
                                   if (isOwner || isAdmin || isSuperadmin) {
                                     setEditingCell({ pic, date: dateKey });
                                   } else {
                                     setOwnershipModal({
                                       isOpen: true,
                                       picName: pic,
                                       itemName: `Jadwal ${dateKey}`,
                                       type: 'schedule',
                                       onCancel: () => setOwnershipModal(prev => ({ ...prev, isOpen: false })),
                                       onConfirm: () => {
                                         setOwnershipModal(prev => ({ ...prev, isOpen: false }));
                                         setEditingCell({ pic, date: dateKey });
                                       }
                                     });
                                   }
                                 }}
                                 className={cn(
                                   "px-2.5 py-1.5 rounded-xl text-[10px] font-black border text-center shadow-lg transition-all select-none",
                                   (fullName === pic || isAdmin || isSuperadmin) ? "cursor-pointer hover:ring-2 hover:ring-indigo-500/50" : "cursor-help",
                                   holiday ? "text-rose-100 bg-rose-600 border-rose-500" : (SCHEDULE_STATUS_COLORS[displayStatus] || 'text-slate-500 bg-slate-800/10 border-slate-800')
                                 )}
                               >
                                 {displayStatus}
                               </motion.div>
                             ) : (
                               <div 
                                 onClick={() => {
                                   const isOwner = fullName === pic;
                                   if (isOwner || isAdmin || isSuperadmin) {
                                     setEditingCell({ pic, date: dateKey });
                                   } else {
                                     setOwnershipModal({
                                       isOpen: true,
                                       picName: pic,
                                       itemName: `Jadwal ${dateKey}`,
                                       type: 'schedule',
                                       onCancel: () => setOwnershipModal(prev => ({ ...prev, isOpen: false })),
                                       onConfirm: () => {
                                         setOwnershipModal(prev => ({ ...prev, isOpen: false }));
                                         setEditingCell({ pic, date: dateKey });
                                       }
                                     });
                                   }
                                 }}
                                 className={cn(
                                   "w-3 h-3 rounded-full bg-slate-800/50 transition-all shadow-inner",
                                   (fullName === pic || isAdmin || isSuperadmin) ? "cursor-pointer hover:bg-slate-600 hover:scale-125" : ""
                                 )} 
                               />
                             )
                            )}
                         </div>
                       </td>
                     );
                   })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Legend and Modals */}
      {requestModal && (
        <div className="fixed inset-0 z-[200] bg-slate-950/80 backdrop-blur-md flex items-center justify-center p-4">
           <motion.div 
             initial={{ opacity: 0, scale: 0.95, y: 20 }}
             animate={{ opacity: 1, scale: 1, y: 0 }}
             className="bg-slate-900 border border-white/10 rounded-[3rem] w-full max-w-lg shadow-[0_0_50px_rgba(0,0,0,0.5)] p-10 relative overflow-hidden"
           >
              {/* Glass Decorations */}
              <div className="absolute top-0 right-0 w-32 h-32 bg-indigo-500/10 rounded-full blur-3xl -translate-y-1/2 translate-x-1/2" />
              <div className="absolute bottom-0 left-0 w-32 h-32 bg-amber-500/5 rounded-full blur-3xl translate-y-1/2 -translate-x-1/2" />

              <div className="relative">
                <div className="flex items-center gap-5 mb-8">
                  <div className="w-14 h-14 bg-indigo-500/10 rounded-2xl flex items-center justify-center border border-indigo-500/20 shadow-lg shadow-indigo-500/5">
                    <Clock className="w-7 h-7 text-indigo-400" />
                  </div>
                  <div>
                    <h3 className="text-2xl font-black text-white italic uppercase tracking-tighter">Konfirmasi Perubahan Jadwal</h3>
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-[0.3em] mt-1">Status: Pending Approval</p>
                  </div>
                </div>
                <div className="space-y-5 mb-10">
                  <div className="p-6 bg-slate-950/40 rounded-3xl border border-white/5 backdrop-blur-sm">
                    <div className="flex justify-between items-center mb-4">
                      <div>
                        <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-1">PIC TARGET</p>
                        <p className="text-md font-black text-white italic">{requestModal.pic}</p>
                      </div>
                      <div className="text-right">
                        <p className="text-[10px] text-slate-500 font-black uppercase tracking-widest mb-1">TANGGAL</p>
                        <p className="text-md font-black text-slate-300 italic">{format(new Date(requestModal.date), 'dd MMMM yyyy')}</p>
                      </div>
                    </div>

                    <div className="flex items-center gap-6 p-4 bg-black/40 rounded-2xl border border-white/5 relative overflow-hidden group">
                      <div className="flex-1">
                        <p className="text-[8px] text-slate-600 font-black uppercase tracking-widest mb-1">Original</p>
                        <span className="text-xs font-black text-slate-500">{requestModal.currentStatus || 'KOSONG'}</span>
                      </div>
                      <div className="w-8 h-8 bg-indigo-500/10 rounded-full flex items-center justify-center border border-indigo-500/20">
                        <ChevronRight className="w-4 h-4 text-indigo-400" />
                      </div>
                      <div className="flex-1 text-right">
                        <p className="text-[8px] text-indigo-500 font-black uppercase tracking-widest mb-1">Proposed</p>
                        <span className="text-xs font-black text-indigo-400">{requestModal.newStatus || 'DELETE'}</span>
                      </div>
                    </div>
                  </div>

                  {/* Two-Way Swap Feature */}
                  <div className="p-6 bg-slate-950/20 rounded-3xl border border-white/5">
                    <label className="flex items-center gap-3 cursor-pointer group mb-4">
                      <input 
                        type="checkbox" 
                        checked={requestModal.isSwap || false}
                        onChange={(e) => setRequestModal({ ...requestModal, isSwap: e.target.checked })}
                        className="w-5 h-5 rounded-lg border-slate-700 bg-slate-950 text-indigo-600 focus:ring-indigo-500/50"
                      />
                      <span className="text-xs font-black text-slate-300 uppercase tracking-widest group-hover:text-white transition-colors">Tukar 2 Arah (Opsional)</span>
                    </label>

                    {requestModal.isSwap && (
                      <motion.div 
                        initial={{ opacity: 0, height: 0 }}
                        animate={{ opacity: 1, height: 'auto' }}
                        className="space-y-4 pt-2 overflow-hidden"
                      >
                        <div className="grid grid-cols-2 gap-4">
                          <div className="space-y-2">
                            <label className="text-[8px] text-slate-500 font-black uppercase tracking-widest ml-1">Tanggal Pengganti</label>
                            <input 
                              type="date"
                              value={requestModal.swapDate || ''}
                              onChange={(e) => {
                                const date = e.target.value;
                                const swapCurrentStatus = scheduleGrid[requestModal.pic]?.[date] || '';
                                setRequestModal({ ...requestModal, swapDate: date, swapCurrentStatus });
                              }}
                              className="w-full bg-slate-900 border border-slate-700 rounded-xl px-4 py-2.5 text-xs text-white focus:outline-none focus:ring-1 focus:ring-indigo-500/50"
                            />
                          </div>
                          <div className="space-y-2">
                            <label className="text-[8px] text-slate-500 font-black uppercase tracking-widest ml-1">Status Pengganti</label>
                            <select 
                              value={requestModal.swapStatus || ''}
                              onChange={(e) => setRequestModal({ ...requestModal, swapStatus: e.target.value })}
                              className="w-full bg-slate-950 border border-white/10 rounded-xl px-4 py-2.5 text-xs text-indigo-400 font-black uppercase focus:outline-none focus:ring-1 focus:ring-indigo-500/50"
                            >
                              <option value="">Pilih Status</option>
                              {['WFO', 'WFH', 'A2', 'LIBUR'].map(st => (
                                <option key={st} value={st}>{st}</option>
                              ))}
                            </select>
                          </div>
                        </div>
                        {requestModal.swapDate && (
                          <div className="p-3 bg-black/20 rounded-xl border border-white/5 text-center">
                            <p className="text-[8px] text-slate-600 font-black uppercase tracking-tighter mb-1">Status Saat Ini pada {requestModal.swapDate}</p>
                            <span className="text-[10px] font-black text-slate-400">{requestModal.swapCurrentStatus || 'KOSONG'}</span>
                          </div>
                        )}
                      </motion.div>
                    )}
                  </div>

                  <div className="space-y-3">
                    <label className="text-[10px] text-slate-500 font-black uppercase tracking-widest ml-2">Alasan Reschedule</label>
                    <textarea 
                      id="reschedule-reason"
                      className="w-full bg-slate-950/50 border border-white/10 rounded-3xl p-5 text-sm text-white focus:outline-none focus:ring-2 focus:ring-indigo-500/30 transition-all placeholder:text-slate-700 min-h-[120px]"
                      placeholder="Contoh: Sakit, Ada keperluan mendesak, Tukar shift..."
                      rows={3}
                    />
                  </div>
                </div>

                <div className="flex gap-4">
                  <button 
                    onClick={() => setRequestModal(null)}
                    className="flex-1 py-4 bg-slate-800/50 hover:bg-slate-800 text-slate-400 rounded-2xl text-[11px] font-black uppercase tracking-[0.2em] transition-all border border-white/5"
                  >
                    Batalkan
                  </button>
                  <button 
                    onClick={async () => {
                      const reason = (document.getElementById('reschedule-reason') as HTMLTextAreaElement).value;
                      if (!reason.trim()) {
                        alert('Harap masukkan alasan penukaran jadwal.');
                        return;
                      }

                      if (requestModal.isSwap) {
                        if (!requestModal.swapDate || !requestModal.swapStatus) {
                          alert('Harap lengkapi detail tukar jadwal.');
                          return;
                        }
                        if (requestModal.swapDate === requestModal.date) {
                          alert('Tanggal pengganti tidak boleh sama dengan tanggal utama.');
                          return;
                        }

                        // Check duplicate for swap date
                        try {
                          const isSwapPending = await taskService.checkExistingRescheduleRequest(requestModal.pic, requestModal.swapDate);
                          if (isSwapPending) {
                            alert(`Tanggal pengganti (${requestModal.swapDate}) sedang dalam status Pending approval.`);
                            return;
                          }
                        } catch (e) { console.error(e); }
                      }

                      try {
                        setLoading(true);
                        
                        // Single payload for multi-way swaps
                        const payload = {
                          pic_name: requestModal.pic,
                          schedule_date: requestModal.date,
                          original_status: requestModal.currentStatus,
                          new_status: requestModal.newStatus,
                          reason: reason.trim() + (requestModal.isSwap ? ` (Tukar dengan ${requestModal.swapDate})` : ''),
                          requested_by: fullName || user?.email || user?.name || 'User',
                          swap_date: requestModal.isSwap ? requestModal.swapDate : null,
                          swap_status: requestModal.isSwap ? requestModal.swapStatus : null
                        };

                        await taskService.createRescheduleRequest(payload);
                        onNotif?.("Request submitted for approval.");
                        setRequestModal(null);
                        fetchSchedules();
                      } catch (err: any) {
                        console.error(err);
                        alert(`Gagal mengirim permohonan: ${err.message || 'Unknown Error'}`);
                      } finally {
                        setLoading(false);
                      }
                    }}
                    className="flex-3 py-4 bg-indigo-600 hover:bg-indigo-500 text-white rounded-2xl text-[11px] font-black uppercase tracking-[0.2em] shadow-xl shadow-indigo-500/20 transition-all border border-white/20 active:scale-95"
                  >
                    Kirim Pengajuan
                  </button>
                </div>
              </div>
           </motion.div>
        </div>
      )}

      {showRescheduleRequests && (
        <div className="fixed inset-0 z-[200] bg-slate-950/80 backdrop-blur-sm flex items-center justify-center p-4">
           <motion.div 
             initial={{ opacity: 0, scale: 0.95, y: 20 }}
             animate={{ opacity: 1, scale: 1, y: 0 }}
             className="bg-slate-900 border border-slate-800 rounded-[2.5rem] w-full max-w-4xl max-h-[85vh] shadow-2xl flex flex-col relative overflow-hidden"
           >
              <div className="p-8 border-b border-slate-800 flex justify-between items-center">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20">
                    <History className="w-6 h-6 text-amber-500" />
                  </div>
                  <div>
                    <h3 className="text-xl font-black text-white italic uppercase tracking-tight">Om Dedy: APPROVAL MENU</h3>
                    <p className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Incoming and Pending Requests</p>
                  </div>
                </div>
                <button onClick={() => setShowRescheduleRequests(false)} className="p-3 bg-slate-800 hover:bg-slate-700 text-white rounded-2xl transition-all">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-8">
                {rescheduleRequests.length === 0 ? (
                  <div className="flex flex-col items-center justify-center py-20 gap-4 opacity-40">
                    <History className="w-16 h-16 text-slate-500" />
                    <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">No Pending Requests Found</p>
                  </div>
                ) : (
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                    {rescheduleRequests.map((req: any, ri: number) => (
                      <div key={getSafeKey(req, ri, 'resched-req')} className="bg-slate-950/50 border border-slate-800 p-6 rounded-3xl space-y-4 hover:border-indigo-500/30 transition-all">
                        <div className="flex justify-between items-start">
                          <div>
                            <p className="text-xs font-black text-white italic">{req.pic_name}</p>
                            <p className="text-[10px] text-slate-500 font-bold uppercase">{req.schedule_date}</p>
                          </div>
                          <div className={cn(
                            "px-3 py-1 text-[10px] font-black uppercase tracking-widest rounded-lg border",
                            req.status === 'Approved' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                            req.status === 'Rejected' ? "bg-rose-500/10 text-rose-500 border-rose-500/20" :
                            "bg-amber-500/10 text-amber-500 border-amber-500/20"
                          )}>
                            {req.status}
                          </div>
                        </div>
                        
                        <div className="flex items-center gap-4 bg-black/20 p-4 rounded-2xl">
                          <div className="flex-1 text-center">
                            <p className="text-[8px] text-slate-500 font-bold uppercase mb-1">From</p>
                            <span className="text-[10px] font-black text-slate-400 bg-slate-800 px-2.5 py-1 rounded-lg border border-slate-700">{req.original_status || 'BLANK'}</span>
                          </div>
                          <ChevronRight className="w-4 h-4 text-slate-700" />
                          <div className="flex-1 text-center">
                            <p className="text-[8px] text-indigo-500 font-bold uppercase mb-1">To</p>
                            <span className="text-[10px] font-black text-indigo-400 bg-indigo-500/10 px-2.5 py-1 rounded-lg border border-indigo-500/20">{req.new_status}</span>
                          </div>
                        </div>

                        <div className="p-4 bg-slate-900 rounded-2xl border border-slate-800/50">
                          <p className="text-[8px] text-slate-500 font-bold uppercase mb-2">Reason</p>
                          <p className="text-xs text-slate-300 italic">"{req.reason}"</p>
                        </div>
                        
                        {req.status === 'Pending' && (
                          <div className="flex gap-3">
                            <button 
                              onClick={async () => {
                                if (confirm('Tolak request ini?')) {
                                  try {
                                    await taskService.updateRescheduleRequestStatus(req.id, 'Rejected', user?.email || 'Admin');
                                    fetchSchedules();
                                  } catch (err) { alert('Gagal memproses approval'); }
                                }
                              }}
                              className="flex-1 py-3 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 text-slate-500 rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all"
                            >
                              Reject
                            </button>
                            <button 
                              onClick={async () => {
                                if (confirm('Setujui request ini? Jadwal akan terupdate otomatis.')) {
                                  try {
                                    await taskService.updateRescheduleRequestStatus(req.id, 'Approved', user?.email || 'Admin');
                                    fetchSchedules();
                                  } catch (err) { alert('Gagal memproses approval'); }
                                }
                              }}
                              className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-2xl text-[10px] font-black uppercase tracking-widest transition-all shadow-lg shadow-emerald-500/20"
                            >
                              Approve
                            </button>
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                )}
              </div>
           </motion.div>
        </div>
      )}

      <div className="bg-slate-900/50 border border-slate-800 p-6 rounded-[2rem] flex flex-wrap gap-8 items-center backdrop-blur-md">
         <div className="flex items-center gap-3 pr-8 border-r border-slate-800">
            <Filter className="w-4 h-4 text-slate-500" />
            <span className="text-[10px] font-black text-slate-500 uppercase tracking-[0.3em]">Status Legend</span>
         </div>
         {[
           { label: 'WFO - Work From Office', color: 'emerald', code: 'WFO' },
           { label: 'WFH - Work From Home', color: 'indigo', code: 'WFH' },
           { label: 'WFC/A2 - Outstation', color: 'amber', code: 'WFC' },
           { label: 'LIBUR/LATE - Absence', color: 'rose', code: 'LIBUR' },
         ].map(item => (
           <div key={item.code} className="flex items-center gap-3">
              <div className={cn(
                "w-4 h-4 rounded-lg border",
                SCHEDULE_STATUS_COLORS[item.code]
              )} />
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">{item.label}</span>
           </div>
         ))}
      </div>
    </div>
  );
}

function PersonelManagement({ users, projects, currentUser, onRefresh, isAdmin }: { users: AppUser[], projects: Project[], currentUser: any, onRefresh: () => void, isAdmin: boolean }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [accessFilter, setAccessFilter] = useState('all');
  const [selectedPIC, setSelectedPIC] = useState<AppUser | null>(null);
  
  const [confirmConfig, setConfirmConfig] = useState<{ isOpen: boolean, title: string, description: string, onConfirm: () => void, variant?: 'danger' | 'primary' } | null>(null);
  const [showAddModal, setShowAddModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [editingUser, setEditingUser] = useState<AppUser | null>(null);
  const [notif, setNotif] = useState<string | null>(null);

  const filteredUsers = (users || []).filter(u => {
    const matchesSearch = u.name?.toLowerCase().includes(searchTerm.toLowerCase()) || u.email?.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesAccess = accessFilter === 'all' || u.access_level === accessFilter;
    return matchesSearch && matchesAccess;
  });

  const startEditing = (user: AppUser) => {
    setEditingUser(user);
    setShowEditModal(true);
  };

  const handleDelete = async (user: AppUser) => {
    if (!isAdmin) return;
    setConfirmConfig({
      isOpen: true,
      title: 'Hapus Personel?',
      description: `Apakah Anda yakin ingin menghapus ${user.name} (${user.email}) dari sistem? Tindakan ini tidak dapat dibatalkan.`,
      variant: 'danger',
      onConfirm: async () => {
        try {
          await taskService.deleteUser(user.id, currentUser?.email || 'System');
          setNotif(`Personnel ${user.name} removed successfully`);
          onRefresh();
          if (selectedPIC?.id === user.id) setSelectedPIC(null);
        } catch (err) {
          console.error('Delete user failed:', err);
        }
      }
    });
  };

  return (
    <div className="space-y-6">
      <SuccessNotification show={!!notif} message={notif || ''} onClose={() => setNotif(null)} />
      
      <AnimatePresence>
        {confirmConfig && (
          <ConfirmModal 
            isOpen={confirmConfig.isOpen}
            onClose={() => setConfirmConfig(null)}
            onConfirm={confirmConfig.onConfirm}
            title={confirmConfig.title}
            description={confirmConfig.description}
            variant={confirmConfig.variant}
            confirmText={confirmConfig.variant === 'primary' ? 'Ya, Simpan' : 'Hapus'}
          />
        )}
        {showAddModal && (
          <AddPersonnelModal 
            isAdmin={isAdmin}
            currentUserEmail={currentUser?.email || 'System'}
            onClose={() => setShowAddModal(false)}
            onSuccess={() => {
              setNotif('New personnel onboarded successfully');
              onRefresh();
            }}
          />
        )}
        {showEditModal && editingUser && (
          <EditPersonnelModal 
            user={editingUser}
            isAdmin={isAdmin}
            currentUserEmail={currentUser?.email || 'System'}
            onClose={() => {
              setShowEditModal(false);
              setEditingUser(null);
            }}
            onConfirmSave={async (updatedData) => {
              try {
                await taskService.updateUser(editingUser.id, updatedData, currentUser?.email || 'System');
                setNotif('Data berhasil diperbarui');
                setShowEditModal(false);
                setEditingUser(null);
                onRefresh();
              } catch (err: any) {
                alert('Update failed: ' + err.message);
              }
            }}
          />
        )}
      </AnimatePresence>

      <div className="flex flex-col md:flex-row gap-4 items-center justify-between mb-8">
        <div className="flex gap-4 w-full md:w-auto">
          <div className="relative flex-1 md:w-64">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
            <input 
              type="text"
              placeholder="Search PIC or Email..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full bg-slate-900 border border-slate-800 rounded-xl pl-10 pr-4 py-2 text-xs text-slate-200 focus:border-indigo-500 outline-none transition-all"
            />
          </div>
          <select 
            value={accessFilter}
            onChange={(e) => setAccessFilter(e.target.value)}
            className="bg-slate-900 border border-slate-800 rounded-xl px-4 py-2 text-xs text-slate-400 focus:border-indigo-500 outline-none"
          >
            <option value="all">All Access Levels</option>
            <option value="Superadmin">Superadmin</option>
            <option value="Admin">Admin</option>
            <option value="PIC">PIC</option>
            <option value="Developer">Developer</option>
            <option value="QA">QA</option>
          </select>
        </div>
        {isAdmin && (
          <button 
            onClick={() => setShowAddModal(true)}
            className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white px-4 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center gap-2 transition-all active:scale-95"
          >
            <UserPlus className="w-4 h-4" /> Add Personnel
          </button>
        )}
      </div>

      <div className="bg-slate-900/50 border border-slate-800/60 rounded-2xl overflow-hidden shadow-2xl">
        <table className="w-full text-left">
          <thead>
            <tr className="border-b border-slate-800 bg-slate-900/50">
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Nama PIC</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Email PIC</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Password</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Access Level</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Role</th>
              <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] text-right">Action</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800/50">
            {filteredUsers.map((u, i) => {
              const canEdit = isAdmin || u.email === currentUser?.email;
              const personnelKey = getSafeKey(u, i, 'user');

              return (
                <tr 
                  key={personnelKey} 
                  className="hover:bg-indigo-500/5 transition-all group cursor-pointer border-b border-slate-800/30"
                  onClick={() => setSelectedPIC(u)}
                >
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-3">
                      <div className="w-8 h-8 rounded-full bg-slate-800 border border-slate-700 flex items-center justify-center font-bold text-[10px] text-indigo-400 uppercase">
                        {u.name?.charAt(0)}
                      </div>
                      <span className="font-bold text-slate-200">{u.name}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4">
                     <span className="text-[11px] text-slate-400 font-mono">{u.email}</span>
                  </td>
                  <td className="px-6 py-4">
                     <span className="text-[11px] text-slate-500 font-mono">••••••••</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="px-2 py-1 rounded bg-indigo-500/10 border border-indigo-500/20 text-indigo-400 text-[10px] font-black uppercase tracking-wider">
                      {u.access_level}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="font-black text-slate-600 text-[10px] uppercase tracking-[0.15em]">{u.role}</span>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity">
                      {canEdit && (
                        <button 
                          onClick={e => { e.stopPropagation(); startEditing(u); }}
                          className="p-2 hover:bg-slate-800 rounded-lg text-slate-500 hover:text-indigo-400"
                          title="Edit Personnel"
                        >
                          <Edit3 className="w-4 h-4" />
                        </button>
                      )}
                      <button 
                        onClick={e => { e.stopPropagation(); setSelectedPIC(u); }}
                        className="p-2 hover:bg-slate-800 rounded-lg text-slate-500 hover:text-indigo-400"
                        title="View Info"
                      >
                        <ExternalLink className="w-4 h-4" />
                      </button>
                      {isAdmin && (
                        <button 
                          onClick={e => { e.stopPropagation(); handleDelete(u); }}
                          className="p-2 hover:bg-rose-500/10 rounded-lg text-slate-500 hover:text-rose-500"
                          title="Hapus Personel"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <AnimatePresence>
        {selectedPIC && (
          <div className="fixed inset-0 z-[60] flex justify-end">
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="w-full max-w-md bg-slate-900 border-l border-slate-800 shadow-2xl flex flex-col"
            >
              <div className="p-6 border-b border-slate-800 flex items-center justify-between bg-slate-900/50 backdrop-blur-md">
                <div className="flex items-center gap-4">
                  <div className="w-12 h-12 rounded-2xl bg-indigo-600 flex items-center justify-center text-xl font-black text-white italic">
                    {selectedPIC.name?.charAt(0)}
                  </div>
                  <div>
                    <h3 className="font-black text-white uppercase italic tracking-tighter text-lg">{selectedPIC.name}</h3>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{selectedPIC.email}</p>
                  </div>
                </div>
                <button 
                  onClick={() => setSelectedPIC(null)}
                  className="p-2 hover:bg-slate-800 rounded-xl transition-colors text-slate-500"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
              <div className="flex-1 overflow-auto p-6 space-y-6">
                <div>
                   <h4 className="text-[10px] font-black text-slate-600 uppercase tracking-[0.2em] mb-4">Credentials</h4>
                   <div className="space-y-4">
                      <div className="space-y-1.5">
                        <label className="text-[9px] font-bold text-slate-500 uppercase px-1">Crypto-Key (Password)</label>
                        <input 
                          type="password"
                          defaultValue={selectedPIC.password || ''}
                          className="w-full bg-slate-950 border border-slate-800 rounded-lg px-3 py-2 text-xs text-slate-300 outline-none focus:border-indigo-500 transition-all font-mono"
                          placeholder="Set password..."
                        />
                      </div>
                   </div>
                </div>

                <div>
                   <h4 className="text-[10px] font-black text-slate-600 uppercase tracking-[0.2em] mb-4">Assigned Projects</h4>
                   <div className="space-y-3">
                      {projects.filter(p => true).map((p, i) => (
                        <div key={getSafeKey(p, i, 'pic-project')} className="bg-slate-950/50 border border-slate-800 rounded-xl p-4 flex items-center justify-between hover:border-indigo-500/50 transition-all group">
                            <div>
                              <p className="text-xs font-bold text-slate-200 uppercase">{p.name}</p>
                              <p className="text-[9px] text-slate-500 font-bold uppercase tracking-tighter mt-1">Status: <span className="text-indigo-400">{p.status}</span></p>
                            </div>
                            <button className="p-2 bg-slate-800 rounded-lg text-slate-500 hover:text-white transition-opacity opacity-0 group-hover:opacity-100">
                              <ArrowRight className="w-4 h-4" />
                            </button>
                          </div>
                      ))}
                   </div>
                </div>
              </div>
            </motion.div>
            <div className="absolute inset-0 bg-black/40 backdrop-blur-sm -z-10" onClick={() => setSelectedPIC(null)} />
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

function KanbanView({ projects, tasks, onOpenGantt, onUpdateProject }: { projects: Project[], tasks: Task[], onOpenGantt: (id: string) => void, onUpdateProject: (id: string, updates: Partial<Project>) => void }) {
  const [picFilter, setPicFilter] = useState('all');
  const [projectSearch, setProjectSearch] = useState('');

  const columns = [
    { id: ProjectStatus.FSD_PROGRESS, label: 'FSD on Progress', color: 'indigo' },
    { id: ProjectStatus.FSD_REVIEW, label: 'FSD on Review', color: 'amber' },
    { id: ProjectStatus.SIT_PROGRESS, label: 'SIT on Progress', color: 'emerald' },
    { id: ProjectStatus.UAT_PROGRESS, label: 'UAT on Progress', color: 'cyan' },
    { id: ProjectStatus.PROJECT_LATE, label: 'Project Late', color: 'rose' },
  ];

  const safeProjects = projects || [];
  const safeTasks = tasks || [];

  const filteredProjects = safeProjects.filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(projectSearch.toLowerCase());
    return matchesSearch;
  });

  const getProjectsByStatus = (status: string) => {
    return (filteredProjects || []).filter(p => p.status === status);
  };

  return (
    <div className="flex flex-col h-full space-y-6">
      <header className="sticky top-0 z-20 bg-[#020617]/80 backdrop-blur-md py-4 border-b border-slate-800/50 mb-6 flex items-center justify-between gap-4">
        <div className="flex items-center gap-4 flex-1 max-w-2xl">
          <div className="relative flex-1">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-3.5 h-3.5 text-slate-500" />
            <input 
              type="text"
              placeholder="Filter Project Name..."
              value={projectSearch}
              onChange={(e) => setProjectSearch(e.target.value)}
              className="w-full bg-slate-900/50 border border-slate-800 rounded-xl pl-9 pr-4 py-2 text-[11px] text-slate-300 focus:border-indigo-500 outline-none"
            />
          </div>
          <select 
            value={picFilter}
            onChange={(e) => setPicFilter(e.target.value)}
            className="bg-slate-900/50 border border-slate-800 rounded-xl px-4 py-2 text-[11px] text-slate-400 outline-none"
          >
            <option value="all">All PICs</option>
            {/* Unique PICs from tasks or users */}
          </select>
        </div>
        <div className="flex items-center gap-2 px-3 py-1 bg-slate-900/50 border border-slate-800 rounded-lg">
          <div className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
          <span className="text-[9px] font-black text-slate-500 uppercase tracking-widest">Real-time Sync Active</span>
        </div>
      </header>

      <div className="flex-1 overflow-x-auto snap-x pb-6 scrollbar-thin scrollbar-thumb-slate-700 scrollbar-track-transparent">
        <div className="flex gap-6 h-full min-w-max pr-24">
          {columns.map((col, i) => {
            const colProjects = col.id === 'Project Late' 
              ? filteredProjects.filter(p => p.status === 'Project Late') // Or logic for late
              : getProjectsByStatus(col.id);

            return (
              <div key={getSafeKey(col, i, 'kanban-col')} className="flex-1 w-[350px] min-w-[350px] shrink-0 bg-slate-900/30 rounded-2xl border border-slate-800/40 p-4 flex flex-col gap-4 snap-center">
                <div className="flex items-center justify-between px-2">
                  <div className="flex items-center gap-2">
                    <div className={`w-2 h-2 rounded-full bg-${col.color}-500 shadow-[0_0_8px_rgba(var(--${col.color}-500),0.4)]`} />
                    <h3 className="text-[10px] font-black text-slate-400 uppercase tracking-[0.15em]">{col.label}</h3>
                  </div>
                  <span className="bg-slate-800 px-2 py-0.5 rounded text-[10px] font-mono text-slate-500">{colProjects.length}</span>
                </div>

                <div className="flex-1 overflow-y-auto space-y-4 scrollbar-hide">
                      {(colProjects || []).map((p, i) => {
                    const projectTasks = safeTasks.filter(t => t.project_id === p.id);
                    const doneTasks = projectTasks.filter(t => t.status === TaskStatus.DONE).length;
                    const progress = projectTasks.length > 0 ? (doneTasks / projectTasks.length) * 100 : 0;
                    
                    return (
                      <motion.div 
                        layoutId={p.id}
                        key={getSafeKey(p, i, 'kanban-project')}
                        onClick={() => onUpdateProject(p.id, { status: col.id as ProjectStatus })} // Simulation of Drag/Click to move
                        className="bg-slate-900 border border-slate-800 hover:border-indigo-500/50 rounded-xl p-4 shadow-xl cursor-pointer transition-all hover:bg-slate-800/40 group relative overflow-hidden"
                      >
                        <div className="relative z-10">
                          <h4 className="text-xs font-black text-white italic uppercase tracking-tighter mb-2">{p.name}</h4>
                          <div className="flex flex-col gap-1.5 mb-4">
                             <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                               <UserIcon className="w-3 h-3 text-indigo-500" />
                               <span>PIC: <span className="text-slate-300">{p.pic_name || p.leader_email || 'Unassigned'}</span></span>
                             </div>
                             <div className="flex items-center gap-2 text-[9px] font-bold text-slate-500 uppercase tracking-widest">
                               <Clock className="w-3 h-3 text-slate-600" />
                               <span>Updated: <span className="font-mono">{p.updated_at ? format(new Date(p.updated_at), 'MM/dd') : 'N/A'}</span></span>
                             </div>
                          </div>
                          
                          <div className="space-y-1">
                            <div className="flex justify-between text-[9px] font-bold uppercase tracking-tighter italic">
                              <span className="text-slate-500">Infrastructure Health</span>
                              <span className="text-indigo-400">{Math.round(progress)}%</span>
                            </div>
                            <div className="h-1 bg-slate-800 rounded-full overflow-hidden">
                              <motion.div 
                                initial={{ width: 0 }}
                                animate={{ width: `${progress}%` }}
                                className="h-full bg-indigo-500 shadow-[0_0_10px_#6366f1]"
                              />
                            </div>
                          </div>

                          <div className="mt-4 pt-4 border-t border-slate-800 flex items-center justify-between">
                            <button 
                              onClick={(e) => { e.stopPropagation(); onOpenGantt(p.id); }}
                              className="px-3 py-1.5 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 rounded-lg text-[9px] font-black uppercase tracking-widest text-white transition-all shadow-lg shadow-indigo-600/10 active:scale-95"
                            >
                              Open Gantt Detail
                            </button>
                            <div className="w-6 h-6 rounded-full bg-slate-950 flex items-center justify-center font-black text-[8px] text-indigo-500 border border-slate-800">
                              {p.name.charAt(0)}
                            </div>
                          </div>
                        </div>
                        <div className="absolute top-0 right-0 w-16 h-16 bg-indigo-500/5 blur-2xl -mr-8 -mt-8" />
                      </motion.div>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}


function GanttDetailView({ 
  user, users, projectId, setFocusedProjectId, projects, setProjects, tasks, hierarchicalTasks, expandedRows, scale, 
  setRefreshKey, handleToggleExpand, handleUpdateTask, handleOpenAudit, handleDeleteTask, 
  setScale, setTasks, onReschedule, onNotif, isAdmin, isSuperadmin
}: any) {
  const navigate = useNavigate();
  const [activeTab, setActiveTab] = useState<'TASKS' | 'AUDIT'>('TASKS');
  const [rescheduleLogs, setRescheduleLogs] = useState<ProjectRescheduleLog[]>([]);
  // If projectId is null, we show the high-level Global Gantt
  const isGlobalView = !projectId;
  
  const currentProject = (projects || []).find((p: any) => p.id === projectId);

  const handleAddInlineL1 = async () => {
    if (!projectId) return;
    const tempId = crypto.randomUUID();
    const customId = `#PH-${Math.random().toString(36).substr(2, 4).toUpperCase()}`;
      const defaultL1 = {
        id: tempId,
        custom_id: customId,
        project_id: projectId,
        parent_id: null,
        title: '', 
        duration_hours: 0,
        man_hours: 0,
        status: TaskStatus.ON_PROGRESS,
        start_time: null, 
        end_time: null,
        assignee: user?.name || '',
        updated_at: new Date().toISOString()
      };
    try {
      setTasks((prev: Task[]) => [...prev, defaultL1]);
      const { data, error } = await supabase.from('tasks').insert([defaultL1]).select().single();
      if (error) {
        setTasks((prev: Task[]) => prev.filter(t => t.id !== tempId));
        throw error;
      }
      if (data) {
        setTasks((prev: Task[]) => prev.map(t => t.id === tempId ? data : t));
      }
    } catch (err: any) {
      console.error("Failed to add inline L1:", err);
    }
  };

  const handleAddInlineL2 = async (parentId: string) => {
    if (!projectId) return;
    const parent = tasks.find((t: any) => t.id === parentId);
    const tempId = crypto.randomUUID();
    const customId = `#TS-${Math.random().toString(36).substr(2, 4).toUpperCase()}`;
      const defaultL2 = {
        id: tempId,
        custom_id: customId,
        project_id: projectId,
        parent_id: parentId,
        title: '', 
        duration_hours: 0,
        man_hours: 0,
        status: TaskStatus.ON_PROGRESS,
        start_time: (parent?.start_time && parent.start_time !== '') ? parent.start_time : new Date().toISOString(),
        end_time: (parent?.end_time && parent.end_time !== '') ? parent.end_time : addDays(new Date(), 1).toISOString(),
        assignee: user?.name || '',
        updated_at: new Date().toISOString()
      };
    try {
      setTasks((prev: Task[]) => [...prev, defaultL2]);
      if (!expandedRows.has(parentId)) {
        handleToggleExpand(parentId);
      }
      const { data, error } = await supabase.from('tasks').insert([defaultL2]).select().single();
      if (error) {
        setTasks((prev: Task[]) => prev.filter(t => t.id !== tempId));
        throw error;
      }
      if (data) {
        setTasks((prev: Task[]) => prev.map(t => t.id === tempId ? data : t));
      }
    } catch (err: any) {
      console.error("Failed to add inline L2:", err);
    }
  };

  const projectStats = useMemo(() => {
    if (!projectId || !tasks) return { minStart: null, maxEnd: null, totalManHours: 0, totalTasks: 0, totalChildTasks: 0, remainingManHours: 0 };
    
    // THE "MATHEMATICIAN" FIX: Project-Level Scoped Scoped aggregation
    const pTasks = tasks.filter((t: any) => t.project_id === projectId);
    if (pTasks.length === 0) return { minStart: null, maxEnd: null, totalManHours: 0, totalTasks: 0, totalChildTasks: 0, remainingManHours: 0 };

    const validDates = pTasks
      .flatMap(t => [t.start_time, t.end_time])
      .filter(d => {
        const ts = d ? new Date(d).getTime() : 0;
        return ts > 86400000 && !isNaN(ts); // Filter out 1970/Invalid dates
      });

    const minStart = validDates.length > 0 ? new Date(Math.min(...validDates.map(d => new Date(d).getTime()))) : null;
    const maxEnd = validDates.length > 0 ? new Date(Math.max(...validDates.map(d => new Date(d).getTime()))) : null;
    
    const totalTasks = pTasks.filter(t => !t.parent_id).length;
    const totalChildTasks = pTasks.filter(t => !!t.parent_id).length;
    
    // THE "TRUE SUM" PROTOCOL: Direct aggregation of the man_hours column
    const totalManHours = pTasks.reduce((sum, t) => sum + (Number(t.man_hours) || 0), 0);
    const remainingManHours = pTasks
      .filter(t => t.status !== TaskStatus.DONE)
      .reduce((sum, t) => sum + (Number(t.man_hours) || 0), 0);

    return { 
      minStart, 
      maxEnd, 
      totalManHours,
      remainingManHours: Math.max(0, remainingManHours),
      totalTasks,
      totalChildTasks
    };
  }, [tasks, projectId]);

  const displayStart = currentProject?.start_date || (projectStats.minStart && !isNaN(projectStats.minStart.getTime()) ? format(projectStats.minStart, 'yyyy-MM-dd') : '');
  const displayEnd = currentProject?.end_date || (projectStats.maxEnd && !isNaN(projectStats.maxEnd.getTime()) ? format(projectStats.maxEnd, 'yyyy-MM-dd') : '');
  
  useEffect(() => {
    if (projectId) {
      taskService.getProjectRescheduleLogs(projectId).then(setRescheduleLogs);
    }
  }, [projectId]);

  
  // RBAC & Filter logic
  
  const filteredTasks = useMemo(() => {
    if (isGlobalView) return tasks || []; 
    const pTasks = (tasks || []).filter((t: any) => t.project_id === projectId);
    return pTasks;
  }, [tasks, projectId, isGlobalView]);

  // Recalculate hierarchy for filtered tasks
  const projectTree = useMemo(() => {
    if (isGlobalView) {
      const globalRoots = (projects || []).map(p => {
        const pTasks = (tasks || []).filter(t => t.project_id === p.id);
        
        // Build sub-hierarchy (Phases -> Tasks)
        const l1Map = new Map();
        const l1Roots: any[] = [];
        pTasks.forEach(t => {
          if (!t.parent_id) {
            const node = { ...t, children: [] };
            l1Map.set(t.id, node);
            l1Roots.push(node);
          }
        });
        pTasks.forEach(t => {
          if (t.parent_id && l1Map.has(t.parent_id)) {
            l1Map.get(t.parent_id).children.push({ ...t });
          }
        });

        // 1. Sync Phase ranges from children
        l1Roots.forEach(l1 => {
          if (l1.children.length > 0) {
            const range = calculateParentRange(l1.children);
            if (range.from_date) l1.start_time = range.from_date;
            if (range.to_date) l1.end_time = range.to_date;
          }
        });

        // 2. Sync Project range from all tasks
        const projectRange = calculateParentRange(pTasks);
        
        // Use specified project dates if available and tasks don't exceed them, or if no tasks
        const baseStart = projectRange.from_date ? new Date(projectRange.from_date).getTime() : (p.start_date ? new Date(p.start_date).getTime() : new Date(p.created_at).getTime());
        const baseEnd = projectRange.to_date ? new Date(projectRange.to_date).getTime() : (p.end_date ? new Date(p.end_date).getTime() : new Date(p.created_at).getTime() + 86400000);

        // TRUE MATH AGGREGATION for Global View
        const leafTasks = pTasks.filter(t => !pTasks.some(other => other.parent_id === t.id));
        const totalProjectManHours = leafTasks.reduce((sum, t) => sum + (Number(t.man_hours) || 0), 0);

        return {
          id: p.id,
          title: p.name,
          isProject: true,
          status: p.status,
          start_time: new Date(baseStart).toISOString(),
          end_time: new Date(baseEnd).toISOString(),
          assignee: p.pic_name || p.leader_email,
          duration_hours: totalProjectManHours,
          duration_minutes: 0,
          children: l1Roots
        };
      });
      
      const map = new Map();
      globalRoots.forEach(r => map.set(r.id, r)); 
      return { roots: globalRoots, map };
    }

    const roots: any[] = [];
    const map = new Map();
    (filteredTasks || []).forEach((t: any) => map.set(t.id, { ...t, children: [] }));
    (filteredTasks || []).forEach((t: any) => {
      if (t.parent_id && map.has(t.parent_id)) {
        map.get(t.parent_id).children.push(map.get(t.id));
      } else if (!t.parent_id) {
        roots.push(map.get(t.id));
      }
    });

    // Recursive inheritance for dates
    const syncBoundaries = (nodes: any[]) => {
      nodes.forEach(node => {
        if (node.children && node.children.length > 0) {
          syncBoundaries(node.children);
          const range = calculateParentRange(node.children);
          if (range.from_date && !node.start_time) node.start_time = range.from_date;
          if (range.to_date && !node.end_time) node.end_time = range.to_date;
        }
      });
    };

    syncBoundaries(roots);

    return { roots, map };
  }, [filteredTasks, isGlobalView, projects, tasks]);

  const handleExportWBS = () => {
    if (!currentProject) return;
    
    // Convert hierarchical roots into flat list for excel
    const flatList: any[] = [];
    const recurse = (t: any, level: number) => {
      flatList.push({
        'WBS Level': level === 1 ? 'Title Task' : 'Child Task',
        'Task Name': t.title || t.name || 'Untitled',
        'From Date': t.start_time ? format(new Date(t.start_time), 'yyyy-MM-dd') : '-',
        'To Date': t.end_time ? format(new Date(t.end_time), 'yyyy-MM-dd') : '-',
        'Man Hours': t.man_hours || 0,
        'Status': t.status || 'Unknown'
      });
      if (t.children && t.children.length > 0) {
        t.children.forEach((c: any) => recurse(c, level + 1));
      }
    };

    projectTree.roots.forEach((r: any) => recurse(r, 1));
    handleExcelExport(flatList, `Project_WBS_${currentProject.name.replace(/\s+/g, '_')}`);
  };

  if (!projectTree || !projectTree.roots || (projectTree.roots.length === 0 && !isGlobalView)) {
    return (
      <div className="h-full flex items-center justify-center bg-[#020617] text-slate-500 font-bold uppercase tracking-widest flex-col gap-4">
        <div className="w-12 h-12 rounded-full border-2 border-slate-800 flex items-center justify-center opacity-50">
          <Plus className="w-6 h-6" />
        </div>
        {user ? 'Timeline Partition Empty' : 'No tasks assigned yet'}
        {!isGlobalView && user && (
           <button 
             onClick={handleAddInlineL1}
             className="mt-4 bg-indigo-600 hover:bg-indigo-500 text-white px-6 py-2 rounded-xl text-[10px] font-black uppercase tracking-widest flex items-center gap-2"
           >
             <Plus className="w-4 h-4" /> Add First Phase (L1)
           </button>
        )}
      </div>
    );
  }

  return (
    <div className="flex flex-col h-full bg-[#020617] overflow-hidden">
       <AnimatePresence mode="wait">
       <motion.div 
         key={projectId || 'global'}
         initial={{ opacity: 0, scale: 0.98 }}
         animate={{ opacity: 1, scale: 1 }}
         exit={{ opacity: 0, scale: 1.02 }}
         transition={{ duration: 0.3 }}
         className="flex flex-col h-full justify-start gap-4 p-4"
       >
          {/* TOP: Task Manager (Only in Detail View) */}
          {!isGlobalView && (
            <div className="h-auto max-h-[480px] flex flex-col bg-slate-950/20 border border-slate-800/60 rounded-2xl overflow-y-auto shadow-2xl shrink-0 scrollbar-hide">
      <div className="p-4 border-b border-slate-800/60 flex flex-wrap items-center justify-between bg-zinc-900/40 gap-4">
                <div className="flex items-center gap-6">
                  <div className="flex bg-slate-950 p-1 rounded-xl border border-white/5">
                    <button 
                      onClick={() => setActiveTab('TASKS')}
                      className={cn(
                        "px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                        activeTab === 'TASKS' ? "bg-indigo-600 text-white shadow-lg shadow-indigo-600/20" : "text-slate-500 hover:text-slate-300"
                      )}
                    >
                      Infrastructure Breakdown
                    </button>
                    <button 
                      onClick={() => setActiveTab('AUDIT')}
                      className={cn(
                        "px-4 py-1.5 rounded-lg text-[9px] font-black uppercase tracking-widest transition-all",
                        activeTab === 'AUDIT' ? "bg-indigo-600 text-white shadow-lg shadow-indigo-600/20" : "text-slate-500 hover:text-slate-300"
                      )}
                    >
                      Audit Trail
                    </button>
                  </div>

                   {currentProject && (
                    <div className="flex flex-wrap items-center gap-4 pl-4 border-l border-white/5">
                      <div className="flex flex-col">
                        <div className="flex flex-wrap items-center gap-3">
                           <div className="flex items-center gap-2 bg-slate-950 px-3 py-1.5 rounded-lg border border-white/5 shadow-inner group z-[100] relative pointer-events-auto">
                              <Calendar className="w-3.5 h-3.5 text-indigo-400" />
                              <div className="flex items-center gap-1">
                                <input 
                                  type="date"
                                  value={displayStart}
                                  onChange={async (e) => {
                                    const newDate = e.target.value;
                                    if (!newDate || !projectId) return;
                                    try {
                                      const { data, error } = await supabase
                                        .from('projects')
                                        .update({ start_date: newDate })
                                        .eq('id', projectId)
                                        .select()
                                        .single();

                                      if (error) throw error;
                                      
                                      if (data) {
                                        setProjects(prev => prev.map(p => p.id === projectId ? data : p));
                                      }

                                      await taskService.createProjectRescheduleLog({
                                        project_id: projectId,
                                        changed_by: user?.name || user?.email || 'User',
                                        old_start_date: displayStart,
                                        old_end_date: displayEnd,
                                        new_start_date: newDate,
                                        new_end_date: displayEnd,
                                        reason: 'Direct Header Adjustment'
                                      });
                                      onNotif?.("Project timeline updated!");
                                      // setRefreshKey?.((prev: number) => prev + 1); // No longer needed if we update state directly
                                    } catch (err) {
                                      console.error("Failed to update project start date:", err);
                                    }
                                  }}
                                  className="bg-transparent text-[9px] sm:text-[10px] text-indigo-100 font-mono italic font-bold outline-none cursor-pointer hover:text-white transition-colors relative z-[110]"
                                  disabled={!user}
                                />
                                <span className="text-slate-500 mx-0.5">-</span>
                                <input 
                                  type="date"
                                  value={displayEnd}
                                  onChange={async (e) => {
                                    const newDate = e.target.value;
                                    if (!newDate || !projectId) return;
                                    try {
                                      const { data, error } = await supabase
                                        .from('projects')
                                        .update({ end_date: newDate })
                                        .eq('id', projectId)
                                        .select()
                                        .single();

                                      if (error) throw error;

                                      if (data) {
                                        setProjects(prev => prev.map(p => p.id === projectId ? data : p));
                                      }

                                      await taskService.createProjectRescheduleLog({
                                        project_id: projectId,
                                        changed_by: user?.name || user?.email || 'User',
                                        old_start_date: displayStart,
                                        old_end_date: displayEnd,
                                        new_start_date: displayStart,
                                        new_end_date: newDate,
                                        reason: 'Direct Header Adjustment'
                                      });
                                      onNotif?.("Project timeline updated!");
                                      // setRefreshKey?.((prev: number) => prev + 1);
                                    } catch (err) {
                                      console.error("Failed to update project end date:", err);
                                    }
                                  }}
                                  className="bg-transparent text-[9px] sm:text-[10px] text-indigo-100 font-mono italic font-bold outline-none cursor-pointer hover:text-white transition-colors relative z-[110]"
                                  disabled={!user}
                                />
                              </div>
                            </div>
                           <div className="flex items-center gap-2">
                             {user && (
                               <button 
                                 onClick={() => onReschedule(currentProject)}
                                 className="p-2 bg-slate-800 hover:bg-slate-700 text-slate-400 hover:text-indigo-400 rounded-lg transition-colors border border-white/5 shadow-lg"
                                 title="Reschedule Project Timeline"
                               >
                                 <ArrowDown className="w-3.5 h-3.5 rotate-[-90deg]" />
                               </button>
                             )}
                             {user && (
                               <button 
                                 onClick={handleExportWBS}
                                 className="bg-indigo-600/10 hover:bg-indigo-600/20 text-indigo-400 p-2 rounded-lg border border-indigo-500/20 transition-all"
                                 title="Export to Excel"
                               >
                                 <Download className="w-3.5 h-3.5" />
                               </button>
                             )}
                           </div>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
                
                {activeTab === 'TASKS' && user && (
                  <button 
                    onClick={handleAddInlineL1}
                    className="bg-indigo-600 hover:bg-indigo-500 text-white px-5 py-2.5 rounded-xl text-[10px] font-black uppercase tracking-widest shadow-xl shadow-indigo-600/20 transition-all active:scale-95 flex items-center gap-2 border border-indigo-400/30"
                  >
                    <Plus className="w-4 h-4" />
                    + Add New Phase (L1)
                  </button>
                )}
              </div>
              
              <div className="flex-1 overflow-auto scrollbar-hide bg-slate-950/40 relative">
                {activeTab === 'TASKS' ? (
                  <div className="w-full overflow-x-auto scrollbar-hide">
                    <div className="min-w-[800px]">
                      <GanttTree 
                        user={user}
                        users={users}
                        roots={projectTree.roots} 
                        map={projectTree.map} 
                        tasks={filteredTasks}
                        projects={projects}
                        expandedRows={expandedRows}
                        onToggleExpand={handleToggleExpand}
                        onUpdateTask={handleUpdateTask}
                        onOpenAudit={handleOpenAudit}
                        onDeleteTask={handleDeleteTask}
                        onAddSubTask={handleAddInlineL2}
                        disabled={!user}
                      />
                    </div>
                  </div>
                ) : (
                  <div className="p-8">
                     <AuditLogTable logs={rescheduleLogs} />
                  </div>
                )}
              </div>
            </div>
          )}

         {/* Alternative TOP for Global View: Simple Legend / Filter */}
         {isGlobalView && (
           <div className="p-4 bg-slate-950/20 border-b border-slate-800/40 flex items-center justify-between rounded-2xl">
              <div>
                <h3 className="text-white font-black uppercase italic tracking-tighter text-lg">Om Dedy Timeline Overview</h3>
                
              </div>
              <div className="flex gap-6">
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-emerald-500/40 border border-emerald-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Done</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-indigo-500/40 border border-indigo-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">In Progress</span>
                </div>
                <div className="flex items-center gap-2">
                  <div className="w-3 h-3 bg-amber-500/40 border border-amber-500 rounded" />
                  <span className="text-[10px] text-slate-400 font-bold uppercase tracking-widest">Review</span>
                </div>
              </div>
           </div>
         )}

         {/* BOTTOM: Temporal Visualizer */}
         <div className="flex-1 min-h-0 bg-[#020617] border-t border-slate-800/60 flex flex-col">
           <div className="p-4 flex items-center justify-between bg-slate-950/30">
             <div className="flex items-center gap-3">
               <LayoutGrid className="w-4 h-4 text-indigo-500" />
                <span className="text-[10px] font-black text-slate-500 uppercase tracking-widest">Temporal Visualizer ({scale.charAt(0) + scale.slice(1).toLowerCase()} Resolution)</span>
             </div>
             <div className="flex bg-slate-900 p-1 rounded-lg border border-slate-800 scale-90">
               {(['MONTH', 'WEEK', 'DAY', 'HOUR'] as ViewScale[]).map((s, si) => (
                 <button
                   key={`detail-scale-${s}-${si}`}
                   onClick={() => setScale(s)}
                   className={cn(
                     "px-3 py-1 rounded-md text-[10px] font-bold tracking-wider transition-all",
                     scale === s ? "bg-slate-800 text-indigo-400 shadow-sm" : "text-slate-500 hover:text-slate-300"
                   )}
                 >
                   {s}
                 </button>
               ))}
             </div>
           </div>
           <div className="flex-1 overflow-auto bg-[#020617] scrollbar-hide relative border-t border-slate-800/40">
             <div className="flex min-w-[3600px] items-start h-full">
                <TemporalVisualizer 
                  user={user}
                  scale={scale} 
                  tasks={filteredTasks} 
                  projectId={projectId}
                  hierarchicalTasks={projectTree}
                  expandedRows={expandedRows}
                  onToggleExpand={handleToggleExpand}
                  setTasks={setTasks} 
                  projects={projects}
                  isGlobalView={isGlobalView}
                  onSetFocus={(id: string) => navigate(`/project/${id}`)}
                />
             </div>
           </div>
         </div>
       </motion.div>
       </AnimatePresence>
    </div>
  );
}

// --- Anti-Stuttering Input Component ---
function LocalInput({ 
  value, 
  onChange, 
  className, 
  placeholder, 
  type = 'text',
  required = false,
  autoFocus = false,
  disabled = false
}: { 
  value: string, 
  onChange: (v: string) => void, 
  className?: string, 
  placeholder?: string,
  type?: string,
  required?: boolean,
  autoFocus?: boolean,
  disabled?: boolean
}) {
  const [local, setLocal] = useState(value);
  useEffect(() => { setLocal(value); }, [value]);

  if (disabled) return <span className={cn(className, "opacity-70 cursor-default")}>{value || placeholder}</span>;

  return (
    <input 
      type={type}
      required={required}
      autoFocus={autoFocus}
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => local !== value && onChange(local)}
      onKeyDown={(e) => e.key === 'Enter' && e.currentTarget.blur()}
      className={className}
      placeholder={placeholder}
      disabled={disabled}
    />
  );
}

function AddPersonnelModal({ 
  onClose, 
  onSuccess, 
  isAdmin,
  currentUserEmail
}: { 
  onClose: () => void, 
  onSuccess: () => void, 
  isAdmin: boolean,
  currentUserEmail: string
}) {
  const [formData, setFormData] = useState({
    name: '',
    email: '',
    access_level: 'PIC',
    role: 'Staff',
    password: 'password123'
  });
  const [error, setError] = useState<string | null>(null);

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    try {
      await taskService.createUser(formData, currentUserEmail);
      onSuccess();
      onClose();
    } catch (err: any) {
      setError(err.message || 'Failed to add personnel');
    }
  };

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-[95%] max-w-2xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden"
      >
        <div className="p-8 border-b border-white/5 bg-slate-900/50">
          <h3 className="text-2xl font-black text-white italic tracking-tighter uppercase">Onboard Personnel</h3>
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Register new PIC into the OD Ecosystem</p>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-5">
          {error && <div className="p-3 bg-rose-500/10 border border-rose-500/20 text-rose-400 text-[10px] font-bold uppercase rounded-xl tracking-wider">{error}</div>}
          
          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Nama PIC</label>
            <LocalInput 
              required
              value={formData.name}
              onChange={v => setFormData({...formData, name: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-bold"
              placeholder="e.g., John Doe"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Email PIC</label>
            <LocalInput 
              required
              type="email"
              value={formData.email}
              onChange={v => setFormData({...formData, email: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-mono"
              placeholder="email@example.com"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Access Level</label>
              <select 
                value={formData.access_level}
                onChange={e => setFormData({...formData, access_level: e.target.value})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-indigo-400 outline-none focus:border-indigo-500 font-bold"
              >
                <option value="Superadmin">Superadmin</option>
                <option value="Admin">Admin</option>
                <option value="PIC">PIC</option>
                <option value="Developer">Developer</option>
                <option value="QA">QA</option>
              </select>
            </div>
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Role (Text)</label>
              <LocalInput 
                value={formData.role}
                onChange={v => setFormData({...formData, role: v})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-bold"
                placeholder="Developer, QA, etc."
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Security-Key</label>
            <LocalInput 
              required
              type="password"
              value={formData.password}
              onChange={v => setFormData({...formData, password: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-mono"
              placeholder="••••••••"
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              type="submit"
              className="flex-3 py-3 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-indigo-500/20 transition-all"
            >
              Onboard User
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function EditPersonnelModal({ 
  user,
  onClose, 
  onConfirmSave, 
  isAdmin,
  currentUserEmail
}: { 
  user: AppUser,
  onClose: () => void, 
  onConfirmSave: (data: Partial<AppUser>) => void, 
  isAdmin: boolean,
  currentUserEmail: string
}) {
  const [formData, setFormData] = useState({
    name: user.name || '',
    email: user.email || '',
    access_level: user?.access_level || 'PIC',
    role: user.role || 'Staff',
    password: user.password || ''
  });

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    onConfirmSave(formData);
  };

  return (
    <div className="fixed inset-0 z-[120] flex items-center justify-center p-4">
      <motion.div 
        initial={{ opacity: 0 }}
        animate={{ opacity: 1 }}
        exit={{ opacity: 0 }}
        onClick={onClose}
        className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
      />
      <motion.div 
        initial={{ opacity: 0, scale: 0.9, y: 20 }}
        animate={{ opacity: 1, scale: 1, y: 0 }}
        exit={{ opacity: 0, scale: 0.9, y: 20 }}
        className="relative w-[95%] max-w-2xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden"
      >
        <div className="p-8 border-b border-white/5 bg-slate-900/50">
          <h3 className="text-2xl font-black text-white italic tracking-tighter uppercase">Edit Personnel</h3>
          <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Modify credentials for {user.name}</p>
        </div>

        <form onSubmit={handleSubmit} className="p-8 space-y-5">
          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Nama PIC</label>
            <LocalInput 
              required
              value={formData.name}
              onChange={v => setFormData({...formData, name: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 transition-all font-bold"
              placeholder="Full Name"
            />
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Email PIC</label>
            <LocalInput 
              required
              type="email"
              value={formData.email}
              onChange={v => setFormData({...formData, email: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-400 outline-none focus:border-indigo-500 transition-all font-mono read-only:opacity-60"
              placeholder="email@example.com"
            />
          </div>

          <div className="grid grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Access Level</label>
              <select 
                value={formData.access_level}
                disabled={!isAdmin}
                onChange={e => setFormData({...formData, access_level: e.target.value})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-indigo-400 outline-none focus:border-indigo-500 font-bold disabled:opacity-60"
              >
                <option value="Superadmin">Superadmin</option>
                <option value="Admin">Admin</option>
                <option value="PIC">PIC</option>
                <option value="Developer">Developer</option>
                <option value="QA">QA</option>
              </select>
            </div>
            <div className="space-y-1.5">
              <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Role (Text)</label>
              <LocalInput 
                value={formData.role}
                onChange={v => setFormData({...formData, role: v})}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-bold"
                placeholder="Developer, QA, etc."
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-[9px] font-bold text-slate-500 uppercase tracking-[0.2em] px-1">Security-Key</label>
            <LocalInput 
              type="password"
              value={formData.password}
              onChange={v => setFormData({...formData, password: v})}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-4 py-2.5 text-xs text-slate-200 outline-none focus:border-indigo-500 font-mono"
              placeholder="••••••••"
            />
          </div>

          <div className="pt-4 flex gap-4">
            <button 
              type="button"
              onClick={onClose}
              className="flex-1 py-3 bg-slate-800 hover:bg-slate-700 text-slate-300 text-[10px] font-black uppercase tracking-widest rounded-xl transition-all"
            >
              Cancel
            </button>
            <button 
              type="submit"
              className="flex-3 py-3 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[10px] font-black uppercase tracking-widest rounded-xl shadow-lg shadow-indigo-500/20 transition-all font-black uppercase tracking-widest shadow-lg shadow-indigo-500/20 flex items-center justify-center gap-2"
            >
              Update Data
            </button>
          </div>
        </form>
      </motion.div>
    </div>
  );
}

function RescheduleRequestsView({ requests, onRefresh, user, isLoading }: { requests: any[], onRefresh: () => void, user: any, isLoading?: boolean }) {
  const [confirmData, setConfirmData] = useState<{ id: string, status: 'Approved' | 'Rejected' } | null>(null);
  const [activeTab, setActiveTab] = useState<'PENDING' | 'HISTORY'>('PENDING');
  const [deleteId, setDeleteId] = useState<string | null>(null);

  useEffect(() => {
    onRefresh();
  }, []);

  const isAdminAccess = useMemo(() => {
    if (!user) return false;
    const level = user.access_level?.toLowerCase() || '';
    return level.includes('admin');
  }, [user]);

  if (!user || !isAdminAccess) {
    return (
      <div className="flex items-center justify-center h-full">
        <div className="text-center space-y-4">
          <ShieldAlert className="w-16 h-16 text-rose-500 mx-auto" />
          <h2 className="text-2xl font-black text-white uppercase italic">Unauthorized Access</h2>
          <p className="text-slate-500">You do not have permission to access the Approval Menu.</p>
        </div>
      </div>
    );
  }

  const handleAction = async () => {
    if (!confirmData) return;
    try {
      await taskService.updateRescheduleRequestStatus(confirmData.id, confirmData.status, user.email || 'Admin');
      onRefresh();
    } catch (err) { 
      alert('Gagal memproses approval'); 
    } finally {
      setConfirmData(null);
    }
  };

  const handleDelete = async () => {
    if (!deleteId) return;
    try {
      await taskService.deleteRescheduleRequest(deleteId, user.email || 'Admin');
      onRefresh();
    } catch (err) {
      alert('Gagal menghapus request');
    } finally {
      setDeleteId(null);
    }
  };

  const pendingRequests = requests.filter(r => r.status === 'Pending');
  const historyRequests = requests.filter(r => r.status !== 'Pending');
  const currentRequests = activeTab === 'PENDING' ? pendingRequests : historyRequests;

  return (
    <div className="space-y-8 animate-in fade-in slide-in-from-bottom-4 duration-700 h-full flex flex-col">
      <ConfirmModal 
        isOpen={!!confirmData}
        onClose={() => setConfirmData(null)}
        onConfirm={handleAction}
        title={confirmData?.status === 'Approved' ? "Setujui Request?" : "Tolak Request?"}
        description={confirmData?.status === 'Approved' 
          ? "Jadwal personel akan diperbarui secara otomatis di sistem." 
          : "Permohonan reschedule ini akan ditolak dan personel akan diberitahu."}
        variant={confirmData?.status === 'Approved' ? 'primary' : 'danger'}
        confirmText={confirmData?.status === 'Approved' ? "Approve" : "Reject"}
      />

      <ConfirmModal 
        isOpen={!!deleteId}
        onClose={() => setDeleteId(null)}
        onConfirm={handleDelete}
        title="Hapus History Request?"
        description="Data ini akan dihapus permanen dari audit trail reschedule."
        variant="danger"
        confirmText="Hapus Permanen"
      />

      <div className="flex items-center justify-between">
        <div className="flex items-center gap-4">
          <div className="w-14 h-14 bg-amber-500/10 rounded-2xl flex items-center justify-center border border-amber-500/20 shadow-lg shadow-amber-500/5">
            <History className="w-7 h-7 text-amber-400" />
          </div>
          <div>
            <h2 className="text-2xl font-black text-white tracking-tighter italic uppercase">APPROVAL MENU</h2>
            <p className="text-xs text-slate-500 font-bold uppercase tracking-[0.2em] mt-1">Status Change Requests</p>
          </div>
        </div>

        <div className="flex items-center gap-3">
          <button 
            onClick={onRefresh}
            disabled={isLoading}
            className="p-2.5 bg-slate-900/50 border border-slate-800 rounded-xl text-slate-400 hover:text-white transition-all disabled:opacity-50"
            title="Refresh Requests"
          >
            <RefreshCw className={cn("w-5 h-5", isLoading && "animate-spin")} />
          </button>

          <div className="flex bg-slate-900/50 p-1 rounded-xl border border-white/5">
            <button 
              onClick={() => setActiveTab('PENDING')}
              className={cn(
                "px-6 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all flex items-center gap-2",
                activeTab === 'PENDING' ? "bg-amber-600 text-white shadow-lg shadow-amber-600/20" : "text-slate-500 hover:text-slate-300"
              )}
            >
              Pending Requests
              <span className="bg-black/40 px-1.5 py-0.5 rounded-md text-[8px]">{pendingRequests.length}</span>
            </button>
            <button 
              onClick={() => setActiveTab('HISTORY')}
              className={cn(
                "px-6 py-2 rounded-lg text-[10px] font-black uppercase tracking-widest transition-all flex items-center gap-2",
                activeTab === 'HISTORY' ? "bg-slate-700 text-white" : "text-slate-500 hover:text-slate-300"
              )}
            >
              Resolution History
            </button>
          </div>
        </div>
      </div>

      <div className="flex-1 overflow-y-auto pr-4 scrollbar-hide">
        {isLoading ? (
          <div className="h-full bg-slate-900/50 border border-slate-800 rounded-[2.5rem] p-20 flex flex-col items-center justify-center gap-4">
             <div className="w-12 h-12 border-4 border-amber-500/20 border-t-amber-500 rounded-full animate-spin" />
             <p className="text-slate-500 font-bold uppercase tracking-widest text-[10px]">Fetching requests...</p>
          </div>
        ) : currentRequests.length === 0 ? (
          <div className="h-full bg-slate-900/50 border border-slate-800 rounded-[2.5rem] p-20 flex flex-col items-center justify-center gap-4">
            <History className="w-16 h-16 text-slate-700" />
            <p className="text-[10px] font-black text-slate-500 uppercase tracking-[0.4em]">No {activeTab === 'PENDING' ? 'Pending' : 'Resolved'} Requests</p>
          </div>
        ) : (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 pb-20">
            {currentRequests.map((req: any, i: number) => {
              return (
                <div key={getSafeKey(req, i, 'request')} className="bg-slate-900 border border-slate-800 p-6 rounded-[2rem] space-y-5 hover:border-indigo-500/30 transition-all shadow-xl group relative">
                <div className="flex justify-between items-start">
                  <div>
                    <p className="text-sm font-black text-white italic tracking-tight">{req.pic_name}</p>
                    <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest">{req.schedule_date}</p>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className={cn(
                      "px-3 py-1 text-[9px] font-black uppercase tracking-widest rounded-lg border",
                      req.status === 'Approved' ? "bg-emerald-500/10 text-emerald-500 border-emerald-500/20" :
                      req.status === 'Rejected' ? "bg-rose-500/10 text-rose-500 border-rose-500/20" :
                      "bg-amber-500/10 text-amber-500 border-amber-500/20"
                    )}>
                      {req.status}
                    </div>
                    {activeTab === 'HISTORY' && (
                      <button 
                        onClick={() => setDeleteId(req.id)}
                        className="p-1.5 text-slate-600 hover:text-rose-500 hover:bg-rose-500/10 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                      >
                        <Trash2 className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>
                </div>
                
                <div className="space-y-3">
                  <div className="flex items-center gap-4 bg-slate-950/50 p-4 rounded-2xl border border-slate-800/50">
                    <div className="flex-1 text-center">
                      <p className="text-[8px] text-slate-600 font-bold uppercase mb-1">From</p>
                      <span className="text-[10px] font-black text-slate-500">{req.original_status || 'KOSONG'}</span>
                    </div>
                    <ChevronRight className="w-4 h-4 text-slate-700" />
                    <div className="flex-1 text-center">
                      <p className="text-[8px] text-indigo-500 font-bold uppercase mb-1">To</p>
                      <span className="text-[10px] font-black text-indigo-400">{req.new_status}</span>
                    </div>
                  </div>

                  {req.swap_date && (
                    <div className="flex items-center gap-3 p-3 bg-indigo-500/5 rounded-xl border border-indigo-500/10">
                      <Layers className="w-3.5 h-3.5 text-indigo-400" />
                      <div>
                        <p className="text-[7px] text-slate-500 font-black uppercase tracking-widest">Two-Way Swap Detected</p>
                        <p className="text-[9px] text-slate-400 font-bold">Swap with {req.swap_date} ({req.swap_status})</p>
                      </div>
                    </div>
                  )}
                </div>

                <div className="p-4 bg-slate-950/30 rounded-2xl border border-white/5 h-20 overflow-y-auto">
                  <p className="text-[8px] text-slate-600 font-bold uppercase mb-1">Reason</p>
                  <p className="text-xs text-slate-400 italic font-medium leading-relaxed">"{req.reason}"</p>
                </div>

                {activeTab === 'HISTORY' && (
                  <div className="pt-4 border-t border-white/5 grid grid-cols-2 gap-4">
                    <div>
                      <p className="text-[7px] text-slate-600 font-bold uppercase tracking-widest">Processed By</p>
                      <p className="text-[9px] text-slate-400 font-mono mt-0.5">{req.processed_by || 'Admin'}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-[7px] text-slate-600 font-bold uppercase tracking-widest">Action Date</p>
                      <p className="text-[9px] text-slate-400 font-mono mt-0.5">
                        {req.updated_at ? format(new Date(req.updated_at), 'dd/MM/yy HH:mm') : '-'}
                      </p>
                    </div>
                  </div>
                )}

                {req.status === 'Pending' && (
                  <div className="pt-2 flex gap-3 text-xs">
                    <button 
                      onClick={() => setConfirmData({ id: req.id, status: 'Rejected' })}
                      className="flex-1 py-3 bg-slate-800 hover:bg-rose-500/10 hover:text-rose-400 text-slate-500 rounded-xl text-[10px] font-black uppercase tracking-widest transition-all"
                    >
                      Reject
                    </button>
                    <button 
                      onClick={() => setConfirmData({ id: req.id, status: 'Approved' })}
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-500 text-white rounded-xl text-[10px] font-black uppercase tracking-widest transition-all shadow-lg shadow-emerald-500/20"
                    >
                      Approve
                    </button>
                  </div>
                )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function AuditLogView({ logs, projects, users }: { logs: AuditLog[], projects: Project[], users: AppUser[] }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [dateFilter, setDateFilter] = useState('');
  const [selectedLog, setSelectedLog] = useState<AuditLog | null>(null);

  const filteredLogs = logs.filter(log => {
    const actorName = users.find(u => u.email === log.actor)?.name || log.actor || 'System';
    const matchesActor = actorName.toLowerCase().includes(searchTerm.toLowerCase());
    const matchesDate = !dateFilter || (log.created_at && log.created_at.startsWith(dateFilter));
    return matchesActor && matchesDate;
  });

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between gap-4 bg-slate-900/50 p-4 rounded-2xl border border-slate-800">
        <div className="flex-1 relative">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-slate-500" />
          <input 
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Search by Personnel Name..."
            className="w-full bg-slate-950/50 border border-slate-800 rounded-xl pl-10 pr-4 py-2.5 text-xs text-slate-300 outline-none focus:border-indigo-500 transition-all font-bold"
          />
        </div>
        <div className="flex items-center gap-3">
          <label className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Filter Date</label>
          <input 
            type="date"
            value={dateFilter}
            onChange={(e) => setDateFilter(e.target.value)}
            className="bg-slate-900 border border-slate-700 text-white rounded-xl px-4 py-2.5 text-xs outline-none focus:border-indigo-500 transition-all font-bold"
          />
        </div>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl flex flex-col max-h-[calc(100vh-300px)]">
        <div className="flex-1 overflow-y-auto scrollbar-hide">
          <table className="w-full text-left">
            <thead className="sticky top-0 z-10 bg-slate-900 shadow-md">
              <tr className="border-b border-slate-800">
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Timestamp</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Actor (Name)</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Project Mapping</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em]">Operational Action</th>
                <th className="px-6 py-4 text-[10px] font-bold text-slate-500 uppercase tracking-[0.2em] text-right">Insight</th>
              </tr>
            </thead>
        <tbody className="divide-y divide-slate-800/50 text-[12px] text-slate-400">
          {(() => {
            const uniqueLogs = Array.from(new Map((filteredLogs || []).filter(l => !!l).map(l => [l.id, l])).values());
            return uniqueLogs.map((log, i) => {
              const project = projects.find(p => p.id === log.project_id);
              const actorUser = users.find(u => u.email === log.actor);
              const combinedKey = getSafeKey(log, i, 'audit-table');
              return (
                <tr key={combinedKey} className="hover:bg-white/[0.01] transition-colors group">
                    <td className="px-6 py-4 font-mono text-[10px] text-slate-500 whitespace-nowrap italic">
                      {log.created_at ? format(new Date(log.created_at), 'dd/MM/yyyy HH:mm') : 'N/A'}
                    </td>
                    <td className="px-6 py-4">
                      <div className="flex items-center gap-2">
                        <div className="w-6 h-6 rounded-lg bg-indigo-500/10 border border-indigo-500/30 flex items-center justify-center font-black text-[9px] text-indigo-400">
                          {actorUser?.name?.charAt(0) || 'S'}
                        </div>
                        <div className="flex flex-col">
                          <span className="font-black text-slate-200 tracking-tight leading-none mb-1">{actorUser?.name || log.actor || 'System'}</span>
                          <span className="text-[9px] text-slate-600 font-bold uppercase tracking-widest">{log.actor}</span>
                        </div>
                      </div>
                    </td>
                    <td className="px-6 py-4">
                      <span className="px-2 py-1 bg-slate-950 border border-slate-800 rounded text-[10px] text-indigo-400 uppercase font-black tracking-tighter">
                        {project?.name || 'Central Engine'}
                      </span>
                    </td>
                    <td className="px-6 py-4">
                       <span className={cn(
                         "text-[10px] font-black uppercase tracking-widest px-2 py-0.5 rounded",
                         log.action?.includes('Create') ? "bg-emerald-500/10 text-emerald-400 border border-emerald-500/20" :
                         log.action?.includes('Delete') ? "bg-rose-500/10 text-rose-400 border border-rose-500/20" :
                         "bg-indigo-500/10 text-indigo-400 border border-indigo-500/20"
                       )}>
                        {log.action || 'Unknown'}
                       </span>
                    </td>
                    <td className="px-6 py-4 text-right">
                       <button 
                         onClick={() => setSelectedLog(log)}
                         className="text-[10px] font-black uppercase text-white transition-colors tracking-widest px-4 py-1.5 bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 rounded-lg shadow-lg shadow-indigo-500/20 transition-all active:scale-95"
                       >
                         View Details
                       </button>
                    </td>
                  </tr>
                );
              });
            })()}
        </tbody>
          </table>
        </div>
      </div>

      <AnimatePresence>
        {selectedLog && (
          <div className="fixed inset-0 z-[120] flex items-center justify-center p-6">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setSelectedLog(null)}
              className="absolute inset-0 bg-slate-950/90 backdrop-blur-md"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-4xl bg-slate-900 border border-slate-800 rounded-3xl shadow-2xl overflow-hidden max-h-[80vh] flex flex-col"
            >
              <div className="p-6 border-b border-slate-800 flex items-center justify-between">
                <div>
                  <h3 className="text-xl font-black text-white italic uppercase tracking-tighter">Payload Analysis</h3>
                  <p className="text-[10px] text-slate-500 font-bold uppercase tracking-widest mt-1">Transaction ID: {selectedLog.id}</p>
                </div>
                <button onClick={() => setSelectedLog(null)} className="p-2 hover:bg-slate-800 rounded-xl transition-colors">
                  <X className="w-5 h-5 text-slate-500" />
                </button>
              </div>
              <div className="p-8 overflow-y-auto bg-slate-950/50 grow">
                <div className="space-y-6">
                  <label className="text-[10px] font-black text-white uppercase tracking-[0.2em] px-2 block border-l-2 border-indigo-500">Field Comparison Logic</label>
                  <div className="grid grid-cols-2 gap-4">
                    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 text-[10px] font-bold text-slate-500 uppercase">Property</div>
                    <div className="bg-slate-900 border border-slate-800 rounded-xl p-4 text-[10px] font-bold text-slate-500 uppercase">Transition</div>
                  </div>
                  {Object.keys({ ...(selectedLog.old_payload || {}), ...(selectedLog.new_payload || {}) }).map((key, ki) => {
                    // Exclude metadata fields
                    if (['updated_at', 'id', 'created_at', 'project_id', 'task_id', 'user_id', 'created_by_name'].includes(key)) return null;

                    const oldVal = (selectedLog.old_payload as any)?.[key];
                    const newVal = (selectedLog.new_payload as any)?.[key];
                    if (JSON.stringify(oldVal) === JSON.stringify(newVal)) return null;

                    const formatVal = (v: any) => {
                      if (v === null || v === undefined) return 'None';
                      if (typeof v === 'string') return v;
                      return JSON.stringify(v);
                    };

                    return (
                      <div key={`log-detail-payload-${key}-${ki}`} className="grid grid-cols-[150px_1fr] gap-4 items-center py-2 border-b border-white/5">
                        <div className="text-[10px] font-black text-slate-500 uppercase tracking-wider">{key.replace(/_/g, ' ')}</div>
                        <div className="flex items-center gap-3 overflow-hidden">
                           <div className="line-through text-rose-500/60 text-[10px] font-medium truncate">{formatVal(oldVal)}</div>
                           <div className="text-slate-700 font-bold">➔</div>
                           <div className="text-emerald-400 text-[11px] font-black italic">{formatVal(newVal)}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  );
}

const DeferredTextarea = ({ value, onSave, className, placeholder, disabled }: { value: string, onSave: (v: string) => void, className?: string, placeholder?: string, disabled?: boolean }) => {
  const [local, setLocal] = useState(value);
  useEffect(() => setLocal(value), [value]);

  if (disabled) return <div className={cn(className, "opacity-70 text-[8px] overflow-hidden")}>{value || placeholder}</div>;

  return (
    <textarea 
      value={local}
      onChange={(e) => setLocal(e.target.value)}
      onBlur={() => { if (local !== value) onSave(local); }}
      className={className}
      placeholder={placeholder}
      disabled={disabled}
    />
  );
};

function AuditLogTable({ logs }: { logs: ProjectRescheduleLog[] }) {
  if (logs.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center py-20 text-slate-500 gap-3 border-2 border-dashed border-slate-800 rounded-3xl">
        <History className="w-8 h-8 opacity-20" />
        <p className="text-[10px] font-black uppercase tracking-widest">No Reschedule Events Logged</p>
      </div>
    );
  }

  return (
    <div className="overflow-hidden border border-white/5 rounded-2xl bg-slate-900/40">
      <div className="overflow-x-auto">
        <table className="w-full text-left border-collapse min-w-[800px]">
          <thead>
            <tr className="bg-slate-950/50 border-b border-white/5">
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Timestamp</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Authorized By</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Previous Timeline</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest text-center">Target Timeline</th>
              <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-widest">Reason / Justification</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-white/5">
            {logs.map((log, i) => (
              <tr key={getSafeKey(log, i, 'resched-log')} className="hover:bg-white/[0.02] transition-colors">
                <td className="px-6 py-4">
                  <div className="flex flex-col">
                    <span className="text-xs text-slate-200 font-mono italic">{format(new Date(log.created_at), 'MMM dd, yyyy')}</span>
                    <span className="text-[10px] text-slate-600 font-bold">{format(new Date(log.created_at), 'HH:mm:ss')}</span>
                  </div>
                </td>
                <td className="px-6 py-4">
                  <div className="flex items-center gap-2">
                    <div className="w-6 h-6 rounded-full bg-indigo-500/20 border border-indigo-500/30 flex items-center justify-center text-[10px] font-black text-indigo-400">
                      {log.changed_by?.charAt(0) || 'A'}
                    </div>
                    <span className="text-xs font-bold text-slate-300">{log.changed_by}</span>
                  </div>
                </td>
                <td className="px-6 py-4">
                   <div className="flex flex-col items-center gap-1 opacity-50">
                      <span className="text-[9px] text-slate-400 font-mono">{log.old_start_date}</span>
                      <ArrowDown className="w-3 h-3 text-slate-600" />
                      <span className="text-[9px] text-slate-400 font-mono">{log.old_end_date}</span>
                   </div>
                </td>
                <td className="px-6 py-4">
                   <div className="flex flex-col items-center gap-1">
                      <span className="text-[9px] text-emerald-400 font-black font-mono">{log.new_start_date}</span>
                      <ArrowDown className="w-3 h-3 text-indigo-500" />
                      <span className="text-[9px] text-emerald-400 font-black font-mono">{log.new_end_date}</span>
                   </div>
                </td>
                <td className="px-6 py-4">
                  <div className="p-3 bg-slate-950/50 rounded-xl border border-white/5 max-w-md">
                     <p className="text-xs text-slate-400 leading-relaxed italic">"{log.reason}"</p>
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function GanttTree({ user, users, roots, map, tasks, projects, expandedRows, onToggleExpand, onUpdateTask, onOpenAudit, onAddSubTask, onDeleteTask, disabled }: any) {
  
  const renderTaskRows = (task: any, level: number = 0, index: number = 0, parentTask: any = null) => {
    if (!task) return null;
    const isExpanded = expandedRows.has(task.id);
    const children = task.children || [];
    const isProject = !!task.isProject;
    const health = getTaskHealth(task);
    const proj = projects.find((p: any) => p.id === task.project_id) || (isProject ? task : null);

    return (
      <React.Fragment key={`${task.custom_id || task.id}-${index}`}>
        <tr 
          onClick={() => !isProject && children.length > 0 && onToggleExpand(task.id)}
          className={cn(
            "border-b border-white/5 transition-all group cursor-pointer",
            level === 0 ? "bg-slate-900/40" : "bg-slate-800/10",
            level === 1 && isExpanded ? "bg-slate-800/50" : "hover:bg-white/[0.02]",
            health === 'OVERDUE' && "border-l-4 border-l-rose-500 bg-rose-500/5",
            health === 'OVER SLA' && "border-l-4 border-l-amber-500 bg-amber-500/5"
          )}
        >
          {/* Node Selector / Title */}
          <td className="px-6 py-4">
            <div className={cn(
              "flex items-center gap-3",
              level === 1 ? "pl-10" : level > 1 ? "pl-16" : "pl-0"
            )}>
              {(isProject || children.length > 0) && (
                <span className="text-indigo-500 font-mono w-4" onClick={() => onToggleExpand(task.id)}>
                  {isExpanded ? "▼" : "▶"}
                </span>
              )}
              {isProject ? <FolderKanban className="w-4 h-4 text-indigo-400" /> : level > 0 ? <span className="text-indigo-500/60 font-black text-lg select-none leading-none">↳</span> : <Layers className="w-4 h-4 text-slate-500" />}
              <div className="flex-1 min-w-0" onClick={e => e.stopPropagation()}>
                <div className="flex items-center gap-2">
                  {task.custom_id && (
                    <span className="shrink-0 px-1.5 py-0.5 rounded bg-slate-800 border border-slate-700 text-[8px] font-black text-slate-400 font-mono tracking-tighter shadow-sm select-none">
                      {task.custom_id}
                    </span>
                  )}
                  <EditableInput 
                    value={task.title} 
                    onSave={(v) => onUpdateTask(task.id, 'title', v)}
                    className="text-xs font-black text-white italic truncate tracking-tight uppercase bg-transparent outline-none border-none focus:ring-1 focus:ring-emerald-500 w-full"
                    placeholder="ENTER TASK NAME..."
                    disabled={disabled}
                  />
                  <HealthBadge health={health} />
                  {(() => {
                    const collisions = getCollision(task, tasks, projects);
                    const conflict = level > 0 && parentTask && isWBSConflict(parentTask, task);
                    
                    return (
                          <div className="flex items-center gap-1">
                            {conflict && (
                              <div className="relative group/conflict inline-flex items-center justify-center ml-1">
                                <div className="text-amber-500 animate-pulse cursor-help">
                                  <AlertTriangle className="w-4 h-4" />
                                </div>
                                <div className="absolute bottom-full mb-2 hidden group-hover/conflict:block w-max max-w-xs bg-slate-900 border border-amber-500/30 text-amber-200 text-[9px] rounded p-2 z-[100] shadow-xl font-bold uppercase tracking-widest leading-tight">
                                  OUTSIDE PARENT RANGE: Tanggal di luar jadwal task utama!
                                </div>
                              </div>
                            )}
                          </div>
                    );
                  })()}
                </div>
                {level === 0 && !isProject && (() => {
                    const children = map.get(task.id) || [];
                    if (children.length > 0) {
                      const capacity = task.man_hours || 0;
                      const used = children.reduce((sum: number, c: any) => sum + (c.man_hours || 0), 0);
                      if (used > capacity) return <span className="text-[7px] font-black text-rose-500 uppercase mt-0.5">⚠️ Over-Allocated</span>;
                      if (used < capacity) return <span className="text-[7px] font-black text-amber-500 uppercase mt-0.5">⚠️ Under-Allocated</span>;
                      return <span className="text-[7px] font-black text-emerald-500 uppercase mt-0.5">✅ Alokasi Pas</span>;
                    }
                    return null;
                })()}
              </div>
            </div>
          </td>

          {/* Man-Hours */}
          <td className="px-2 py-4 text-center">
             <div className="flex flex-col items-center gap-1" onClick={e => e.stopPropagation()}>
                <div className="flex items-center gap-1">
                  <EditableInput 
                    type="number"
                    value={(task.man_hours || 0).toString()} 
                    onSave={(v) => onUpdateTask(task.id, 'man_hours', parseFloat(v) || 0)}
                    className="w-12 bg-slate-800 text-[10px] text-center rounded border border-slate-700 text-indigo-400 font-black"
                    disabled={disabled}
                  />
                  <span className="text-[8px] text-slate-600 font-bold">h</span>
                </div>
                <div className="text-[7px] text-slate-500 font-black uppercase tracking-tighter bg-slate-900/50 px-1.5 py-0.5 rounded border border-white/5">
                  {formatWorkday(task.man_hours || 0)}
                </div>
             </div>
          </td>

          {/* Assignee */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
            <EditableInput 
              value={task.assignee || ''} 
              onSave={(v) => onUpdateTask(task.id, 'assignee', v)}
              className="bg-slate-800/80 text-[10px] font-bold text-center p-1 rounded border border-slate-700 text-slate-300 outline-none focus:border-indigo-500 w-full"
              placeholder="PIC"
              disabled={disabled}
            />
          </td>

          {/* Dev */}
          <td className="px-2 py-4" onClick={e => e.stopPropagation()}>
            {level > 0 && (
              <EditableInput 
                value={task.developer_name || ''} 
                onSave={(v) => onUpdateTask(task.id, 'developer_name', v)}
                className="bg-slate-950/50 border border-slate-800 text-[10px] px-1.5 py-1 rounded text-indigo-400/80 text-center hover:border-slate-700 transition-all font-mono w-full"
                placeholder="Dev"
                disabled={disabled}
              />
            )}
          </td>

          {/* QA */}
          <td className="px-2 py-4" onClick={e => e.stopPropagation()}>
            {level > 0 && (
              <EditableInput 
                value={task.qa_name || ''} 
                onSave={(v) => onUpdateTask(task.id, 'qa_name', v)}
                className="bg-slate-950/50 border border-slate-800 text-[10px] px-1.5 py-1 rounded text-purple-400/80 text-center hover:border-slate-700 transition-all font-mono w-full"
                placeholder="QA"
                disabled={disabled}
              />
            )}
          </td>

          {/* Dates */}
          <td 
            className="px-1 py-4 text-center cursor-pointer hover:bg-slate-800/80 transition-all group/date relative" 
            onClick={(e) => { 
              if (disabled) return;
              e.stopPropagation(); 
              const input = e.currentTarget.querySelector('input');
              if (input) {
                try { input.showPicker(); } catch(err) { input.focus(); }
              }
            }}
          >
            <div className="flex flex-col items-center">
              <input 
                type="date" 
                value={task.start_time ? format(new Date(task.start_time), "yyyy-MM-dd") : ''}
                onChange={(e) => {
                  const val = e.target.value;
                  onUpdateTask(task.id, 'start_time', val ? new Date(val).toISOString() : null);
                }}
                min={proj?.start_date ? format(new Date(proj.start_date), "yyyy-MM-dd") : ''}
                max={task.end_time ? format(new Date(task.end_time), "yyyy-MM-dd") : (parentTask ? format(new Date(parentTask.end_time), "yyyy-MM-dd") : '')}
                className="bg-slate-900 border border-slate-700 text-[10px] text-white font-mono focus:text-indigo-400 outline-none w-full text-center cursor-pointer font-bold rounded px-1 py-0.5"
                disabled={disabled}
              />
              <span className="text-[7px] text-indigo-500/50 font-black uppercase opacity-0 group-hover/date:opacity-100 transition-all absolute -top-1">Start</span>
            </div>
          </td>

          <td 
            className="px-1 py-4 text-center cursor-pointer hover:bg-slate-800/80 transition-all group/date relative" 
            onClick={(e) => { 
              if (disabled) return;
              e.stopPropagation(); 
              const input = e.currentTarget.querySelector('input');
              if (input) {
                try { input.showPicker(); } catch(err) { input.focus(); }
              }
            }}
          >
            <div className="flex flex-col items-center">
              <input 
                type="date" 
                value={task.end_time ? format(new Date(task.end_time), "yyyy-MM-dd") : ''}
                onChange={(e) => {
                  const val = e.target.value;
                  onUpdateTask(task.id, 'end_time', val ? new Date(val).toISOString() : null);
                }}
                min={task.start_time ? format(new Date(task.start_time), "yyyy-MM-dd") : (proj?.start_date ? format(new Date(proj.start_date), "yyyy-MM-dd") : '')}
                max={parentTask ? format(new Date(parentTask.end_time), "yyyy-MM-dd") : ''}
                className="bg-slate-900 border border-slate-700 text-[10px] text-white font-mono focus:text-indigo-400 outline-none w-full text-center cursor-pointer font-bold rounded px-1 py-0.5"
                disabled={disabled}
              />
              <span className="text-[7px] text-indigo-500/50 font-black uppercase opacity-0 group-hover/date:opacity-100 transition-all absolute -top-1">End</span>
            </div>
          </td>

          {/* Status */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
            <div className="flex justify-center">
              <TaskStatusSelector 
                status={task.status || TaskStatus.TODO} 
                onUpdate={(v) => onUpdateTask(task.id, 'status', v)} 
                disabled={disabled}
              />
            </div>
          </td>

          {/* Fachrul Feedback */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
             <div className="flex items-center gap-1.5 px-2 py-1.5 bg-slate-900/50 rounded-lg border border-slate-800/50 group-hover:border-slate-700/50 transition-colors">
                <div className="flex flex-col gap-0.5 min-w-[50px]">
                  <ApprovalBadge value={task.approval_fachrul} label="Fachrul" onUpdate={(v) => onUpdateTask(task.id, 'approval_fachrul', v)} disabled={disabled} />
                </div>
                <DeferredTextarea 
                   value={task.suggestion_fachrul || ''}
                   onSave={(v) => onUpdateTask(task.id, 'suggestion_fachrul', v)}
                   className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-200 focus:ring-1 focus:ring-indigo-500/30 outline-none w-full min-h-[32px] max-h-[32px] resize-none scrollbar-hide font-medium"
                   placeholder="..."
                   disabled={disabled}
                />
             </div>
          </td>

          {/* Barra Feedback */}
          <td className="px-4 py-4" onClick={e => e.stopPropagation()}>
             <div className="flex items-center gap-1.5 px-2 py-1.5 bg-slate-900/50 rounded-lg border border-slate-800/50 group-hover:border-slate-700/50 transition-colors">
                <div className="flex flex-col gap-0.5 min-w-[50px]">
                  <ApprovalBadge value={task.approval_barra} label="Barra" onUpdate={(v) => onUpdateTask(task.id, 'approval_barra', v)} disabled={disabled} />
                </div>
                <DeferredTextarea 
                   value={task.suggestion_barra || ''}
                   onSave={(v) => onUpdateTask(task.id, 'suggestion_barra', v)}
                   className="bg-slate-950 border border-slate-800 rounded px-2 py-1 text-[9px] text-slate-200 focus:ring-1 focus:ring-indigo-500/30 outline-none w-full min-h-[32px] max-h-[32px] resize-none scrollbar-hide font-medium"
                   placeholder="..."
                   disabled={disabled}
                />
             </div>
          </td>

          {/* Hidden Actions Column if needed, or just combine in list */}
          <td className="px-6 py-4 text-right">
            <div className="flex items-center justify-end gap-2 opacity-0 group-hover:opacity-100 transition-opacity" onClick={e => e.stopPropagation()}>
              {!isProject && level === 0 && user && (
                <button 
                  onClick={() => onAddSubTask(task.id)}
                  className="bg-gradient-to-r from-indigo-600 to-violet-600 hover:from-indigo-500 hover:to-violet-500 text-white text-[8px] font-black uppercase px-2 py-1 rounded shadow-lg shadow-indigo-500/20 transition-all border border-white/10 active:scale-95"
                >
                  + Sub-task
                </button>
              )}
              <button 
                onClick={() => onOpenAudit(task)}
                className="p-1.5 bg-slate-800 hover:bg-slate-700 text-slate-500 hover:text-indigo-400 rounded-lg transition-all border border-slate-700"
              >
                <History className="w-3.5 h-3.5" />
              </button>
              {user && (
                <button 
                  onClick={() => onDeleteTask(task.id)}
                  className="p-1.5 bg-slate-800 hover:bg-rose-500/20 text-slate-600 hover:text-rose-500 rounded-lg transition-all border border-slate-700"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                </button>
              )}
            </div>
          </td>
        </tr>
        
        {isExpanded && (() => {
          const uniqueChildren = (children || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
          return uniqueChildren.map((sub: any, sidx: number) => (
            <React.Fragment key={`${sub.custom_id || sub.id}-${sidx}`}>
              {renderTaskRows(sub, level + 1, sidx, task)}
            </React.Fragment>
          ));
        })()}
      </React.Fragment>
    );
  };

  return (
    <div className="overflow-x-auto w-full scrollbar-thin scrollbar-thumb-slate-800 scrollbar-track-transparent">
      <table className="w-full text-left border-collapse min-w-[1600px]">
        <thead className="sticky top-0 z-40 bg-slate-900 border-b border-white/5">
          <tr className="shadow-xl">
            <th className="px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-[0.2em] min-w-[300px]">Hierarchy & Governance</th>
            <th className="px-2 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest text-center w-24">Man-Hours</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-40 text-center">PIC</th>
            <th className="px-2 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">Dev</th>
            <th className="px-2 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">QA</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">Start</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-24 text-center">End</th>
            <th className="px-4 py-4 text-[9px] font-black text-slate-600 uppercase tracking-widest w-32 text-center">Status</th>
            <th className="px-4 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest w-[250px]"><div className="flex items-center justify-center gap-1"><UserIcon className="w-2.5 h-2.5"/> Fachrul Feedback</div></th>
            <th className="px-4 py-4 text-[9px] font-black text-indigo-500 uppercase tracking-widest w-[250px]"><div className="flex items-center justify-center gap-1"><UserIcon className="w-2.5 h-2.5"/> Barra Feedback</div></th>
            <th className="px-6 py-4 text-[9px] font-black text-slate-500 uppercase tracking-widest w-40 text-right">Comm Ops</th>
          </tr>
        </thead>
        <tbody className="divide-y divide-white/5">
          {(() => {
            // THE "TERMINATOR" FIX: Unique Task Guard
            const uniqueRoots = (roots || []).filter((v, i, a) => !!v && a.findIndex(t => t.id === v.id) === i);
            if (uniqueRoots.length === 0) {
              return (
                <tr>
                  <td colSpan={11} className="py-10 text-center">
                    <div className="flex flex-col items-center justify-center space-y-4">
                      <div className="w-12 h-12 bg-slate-900/50 rounded-xl flex items-center justify-center border border-slate-800">
                        <Layers className="w-6 h-6 text-slate-700" />
                      </div>
                      <div className="space-y-1">
                        <h4 className="text-slate-300 font-bold text-sm">No Governance Nodes Found</h4>
                        <p className="text-slate-500 text-[10px] font-black uppercase tracking-[0.2em]">Project scope is currently undefined</p>
                      </div>
                    </div>
                  </td>
                </tr>
              );
            }
            return uniqueRoots.map((task: any, idx: number) => {
              const combinedKey = `${task.custom_id || task.id}-${idx}`;
              return <React.Fragment key={combinedKey}>{renderTaskRows(task, 0, idx, null)}</React.Fragment>;
            });
          })()}
        </tbody>
      </table>
    </div>
  );
}

function TemporalVisualizer({ user, scale, tasks, projectId, hierarchicalTasks, expandedRows, onToggleExpand, setTasks, projects, isGlobalView, onSetFocus }: { user: any, scale: ViewScale, tasks: Task[], projectId?: string | null, hierarchicalTasks: any, expandedRows: Set<string>, onToggleExpand: (id: string) => void, setTasks: React.Dispatch<React.SetStateAction<Task[]>>, projects: Project[], isGlobalView?: boolean, onSetFocus?: (id: string) => void }) {
  const { gridStart, gridEnd, intervals, totalDuration } = useMemo(() => {
    return calculateTimelineRange(scale, tasks, projects, projectId, isGlobalView);
  }, [tasks, scale, projects, projectId, isGlobalView]);

  const CELL_WIDTH = scale === 'HOUR' ? 80 : 100;
  const gridWidth = intervals.length * CELL_WIDTH;

  // RED NEEDLE: Precision Today Indicator
  const todayPos = useMemo(() => {
    const now = new Date();
    const startMs = gridStart.getTime();
    if (now.getTime() < startMs || now.getTime() > gridEnd.getTime()) return null;
    return ((now.getTime() - startMs) / totalDuration) * 100;
  }, [gridStart, gridEnd, totalDuration]);

  return (
    <div className="flex h-full w-full overflow-x-auto scrollbar-hide">
      <div className="flex min-w-max h-full">
        {/* Left Column: Task/Project Labels */}
        <div className="w-[280px] border-r border-slate-800 bg-slate-950/90 sticky left-0 z-30 flex flex-col shrink-0">
          <div className="h-10 border-b border-white/5 flex items-center px-6 bg-slate-900/95 sticky top-0 shadow-lg">
            <span className="text-[10px] font-black text-indigo-400 uppercase tracking-[0.2em]">Entity Structure</span>
          </div>
          <div className="flex-1">
            {(() => {
               const uniqueRoots = (hierarchicalTasks.roots || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
               if (uniqueRoots.length === 0) {
                 return (
                   <div className="py-8 px-4 text-center opacity-30 italic text-[10px] uppercase tracking-widest text-slate-500">
                     Empty Roadmap
                   </div>
                 );
               }
               return uniqueRoots.map((root: any, rIdx: number) => {
                 const isExpanded = expandedRows.has(root.id);
                 const rootKey = `${root.custom_id || root.id}-${rIdx}`;
                 return (
                   <div key={rootKey} className="flex flex-col">
                      <div 
                        onClick={() => onToggleExpand(root.id)}
                        className={cn(
                          "flex items-center px-6 border-b border-white/[0.02] text-slate-100 font-black text-[10px] uppercase truncate tracking-widest cursor-pointer hover:bg-indigo-600/5 transition-all group border-l-4 border-l-transparent",
                          isExpanded && "border-l-indigo-600 bg-indigo-600/[0.02]",
                          isGlobalView ? "h-[64px]" : "h-[56px]"
                        )}
                      >
                        <div className="mr-3 text-indigo-500/50 group-hover:text-indigo-400 transition-colors">
                          {isExpanded ? <ChevronDown className="w-3.5 h-3.5" /> : <ChevronRight className="w-3.5 h-3.5" />}
                        </div>
                        <span className="truncate">{root.title || root.custom_id}</span>
                      </div>
                      
                      {isExpanded && (() => {
                        const uniqueChildren = (root.children || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
                        return uniqueChildren.map((child: any, cIdx: number) => {
                          const isChildExpanded = expandedRows.has(child.id);
                          const childKey = `${child.custom_id || child.id}-${cIdx}`;
                          return (
                            <React.Fragment key={childKey}>
                              <div 
                                onClick={() => onToggleExpand(child.id)}
                                className={cn(
                                  "h-[48px] flex items-center px-6 pl-12 border-b border-white/[0.01] text-slate-400 text-[9px] font-bold truncate italic cursor-pointer hover:bg-white/5 transition-colors group"
                                )}
                              >
                                <div className="mr-2 opacity-50 group-hover:opacity-100">
                                  {(child.children && child.children.length > 0) ? (isChildExpanded ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />) : null}
                                </div>
                                <span className="opacity-60 mr-1.5">↳</span> {child.title || child.custom_id}
                              </div>
                              
                              {isChildExpanded && (() => {
                                 const uniqueSubs = (child.children || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
                                 return uniqueSubs.map((sub: any, sIdx: number) => {
                                   const subKey = `${sub.custom_id || sub.id}-${sIdx}`;
                                   return (
                                     <div key={subKey} className="h-[40px] flex items-center px-6 pl-16 border-b border-white/[0.005] text-slate-500 text-[8px] font-medium truncate opacity-70">
                                        <span className="opacity-40 mr-1.5">↳</span> {sub.title || sub.custom_id}
                                     </div>
                                   );
                                 });
                              })()}
                            </React.Fragment>
                          );
                        });
                      })()}
                   </div>
                 );
               });
            })()}
          </div>
        </div>

        <div className="flex-1 overflow-visible relative" style={{ width: gridWidth }}>
          {/* Time Header */}
          <div className="flex sticky top-0 bg-slate-900/95 backdrop-blur-md border-b border-white/5 z-20 h-10 shadow-lg">
            {(intervals || []).map((dt, i) => {
              if (!dt || isNaN(dt.getTime())) return null;
              const iso = dt.toISOString();
              return (
                <div 
                  key={`${iso}-${i}`} 
                  style={{ width: CELL_WIDTH }}
                  className="flex-shrink-0 border-r border-white/5 flex flex-col justify-center px-3"
                >
                  <span className="text-[10px] font-black text-slate-400 uppercase tracking-tighter">
                    {scale === 'HOUR' ? format(dt, 'HH:mm') : scale === 'MONTH' ? format(dt, 'MMM yyyy') : format(dt, 'MMM dd')}
                  </span>
                  <span className="text-[9px] text-slate-600 font-black uppercase tracking-widest opacity-60">
                     {scale === 'HOUR' ? 'REL' : format(dt, 'EEE')}
                  </span>
                </div>
              );
            })}
          </div>

          {/* Grid Lines */}
          <div className="relative min-h-full">
            <div className="absolute inset-0 flex pointer-events-none z-0">
              {(intervals || []).map((dt, i) => (
                <div key={`grid-line-${i}`} style={{ width: CELL_WIDTH }} className="flex-shrink-0 border-r border-white/5" />
              ))}
            </div>

            {/* RED NEEDLE */}
            {todayPos !== null && (
              <div 
                className="absolute top-0 bottom-0 w-[2px] bg-rose-500 z-40 pointer-events-none shadow-[0_0_10px_rgba(244,63,94,0.6)]"
                style={{ left: `${todayPos}%` }}
              >
                <div className="absolute top-0 left-1/2 -translate-x-1/2 w-3 h-3 bg-rose-500 rotate-45 border-2 border-white/20 shadow-lg" />
              </div>
            )}

            {/* Task Bars aligned with Tree rows */}
            <div className="relative z-10 w-full">
              {(() => {
                const uniqueRoots = (hierarchicalTasks.roots || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
                return uniqueRoots.map((root: any, rIdx: number) => {
                  const isExpanded = expandedRows.has(root.id);
                  return (
                    <div key={`${root.custom_id || root.id}-${rIdx}-timeline`} className="flex flex-col">
                      <div className={cn(
                        "flex items-center border-b border-white/[0.02] relative",
                        isGlobalView ? "h-[64px]" : "h-[56px]"
                      )}>
                        <GanttBar 
                          user={user}
                          task={root} 
                          tasks={tasks}
                          projects={projects}
                          setTasks={setTasks}
                          scale={scale} 
                          gridStart={gridStart}
                          gridEnd={gridEnd}
                          totalDuration={totalDuration}
                          isLevel1={true}
                          isProjectBar={root.isProject}
                          onSetFocus={onSetFocus}
                          index={rIdx}
                        />
                      </div>
                      
                      {isExpanded && (() => {
                        const uniqueChildren = (root.children || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
                        return uniqueChildren.map((child: any, cIdx: number) => {
                          const isChildExpanded = expandedRows.has(child.id);
                          return (
                            <React.Fragment key={`${child.custom_id || child.id}-${cIdx}-timeline`}>
                              <div className="h-[48px] flex items-center border-b border-white/[0.01] relative">
                                <GanttBar 
                                  user={user}
                                  task={child} 
                                  tasks={tasks}
                                  projects={projects}
                                  setTasks={setTasks}
                                  scale={scale} 
                                  gridStart={gridStart}
                                  gridEnd={gridEnd}
                                  totalDuration={totalDuration}
                                  isLevel1={!isGlobalView}
                                  isSmall={isGlobalView}
                                  index={cIdx}
                                />
                              </div>
                              
                              {isChildExpanded && (() => {
                                 const uniqueSubs = (child.children || []).filter((v: any, i: number, a: any[]) => !!v && a.findIndex(t => t.id === v.id) === i);
                                 return uniqueSubs.map((sub: any, sIdx: number) => (
                                   <div key={`${sub.custom_id || sub.id}-${sIdx}-timeline`} className="h-[40px] flex items-center border-b border-white/[0.005] relative">
                                     <GanttBar 
                                        user={user}
                                        task={sub} 
                                        tasks={tasks}
                                        projects={projects}
                                        setTasks={setTasks}
                                        scale={scale} 
                                        gridStart={gridStart}
                                        gridEnd={gridEnd}
                                        totalDuration={totalDuration}
                                        isLevel1={false}
                                        isSmall={true}
                                        index={sIdx}
                                      />
                                   </div>
                                 ));
                              })()}
                            </React.Fragment>
                          );
                        });
                      })()}
                   </div>
                  );
                });
              })()}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function GanttBar({ user, task, tasks, projects, setTasks, scale, gridStart, gridEnd, totalDuration, isLevel1, isProjectBar, onSetFocus, index }: any) {
  const [isDragging, setIsDragging] = useState(false);
  const [dragOffset, setDragOffset] = useState(0); 
  const [showPopover, setShowPopover] = useState(false);

  const viewStartMs = startOfDay(gridStart).getTime();
  const totalViewRangeMs = totalDuration;

  // DUAL-DATE PRECISION ENGINE with Robust Fallbacks
  const fromMs = useMemo(() => {
    const d = task.start_time || (task as any).from_date || (task as any).start_date;
    if (!d) return null;
    const date = new Date(d);
    return isNaN(date.getTime()) ? null : startOfDay(date).getTime();
  }, [task.start_time, (task as any).from_date, (task as any).start_date]);

  const toMs = useMemo(() => {
    const d = task.end_time || (task as any).to_date || (task as any).end_date;
    if (!d) return null;
    const date = new Date(d);
    return isNaN(date.getTime()) ? null : startOfDay(date).getTime();
  }, [task.end_time, (task as any).to_date, (task as any).end_date]);

  const left = fromMs !== null ? ((fromMs - viewStartMs) / totalViewRangeMs) * 100 : 0;
  const width = (fromMs !== null && toMs !== null) ? ((toMs - fromMs) / totalViewRangeMs) * 100 : 0;

  const isVisible = (fromMs !== null && toMs !== null) && 
                    (toMs >= viewStartMs) && 
                    (fromMs <= gridEnd.getTime()) &&
                    fromMs > 86400000; // Ignore dates near 1970

  if (!isVisible || fromMs === null || toMs === null) return null;

  const handlePointerDown = (e: React.PointerEvent) => {
    if (isProjectBar || task.status === TaskStatus.DONE || !user || user?.access_level === 'Viewer') return;
    
    setIsDragging(true);
    setShowPopover(false);
    const startX = e.clientX;
    const container = e.currentTarget.closest('[style*="width"]');
    const containerWidth = container?.clientWidth || 1000;
    
    const handleMove = (moveEvent: PointerEvent) => {
      const deltaPx = moveEvent.clientX - startX;
      const deltaTime = (deltaPx / containerWidth) * totalDuration;
      setDragOffset(deltaTime);
    };

    const handleUp = async () => {
      setIsDragging(false);
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);

      if (Math.abs(dragOffset) > 2 * 60000) {
        const newStart = new Date(fromMs + dragOffset);
        const currentDuration = toMs - fromMs;
        const updatedTask = { 
          ...task, 
          start_time: newStart.toISOString(),
          end_time: new Date(newStart.getTime() + currentDuration).toISOString() 
        };
        setTasks((prev: Task[]) => prev.map((t: any) => t.id === task.id ? updatedTask : t));
        setDragOffset(0);
        try {
          await taskService.updateTask(task.id, {
            start_time: updatedTask.start_time,
            end_time: updatedTask.end_time
          }, user?.email || 'Administrator');
        } catch (err) {
          console.error('Update failed:', err);
        }
      } else {
        setDragOffset(0);
      }
    };
    window.addEventListener('pointermove', handleMove);
    window.addEventListener('pointerup', handleUp);
  };

  const isApproved = task.status === TaskStatus.DONE;
  const health = getTaskHealth(task);
  const timeProgress = Math.min(100, Math.max(0, ((new Date().getTime() - fromMs) / (toMs - fromMs)) * 100));

  return (
    <div key={`${task.id}-${index}-${Math.random()}`} className={cn("relative flex items-center w-full", isProjectBar ? "h-16" : (isLevel1 ? "h-14" : "h-10"))}>
      <motion.div
        initial={false}
        animate={{ 
          left: `${left + (dragOffset / totalViewRangeMs) * 100}%`, 
          width: `${Math.max(0.5, width)}%`,
          scale: isDragging ? 1.02 : 1,
          zIndex: isDragging ? 50 : 10
        }}
        transition={{ type: "spring", stiffness: 400, damping: 40 }}
        onPointerDown={handlePointerDown}
        onClick={(e) => {
          if (isProjectBar && onSetFocus) {
            e.stopPropagation();
            onSetFocus(task.id);
          } else {
            setShowPopover(!showPopover);
          }
        }}
        className={cn(
          "absolute rounded-lg flex items-center px-4 shadow-xl select-none transition-all duration-300",
          isLevel1 || isProjectBar 
            ? "h-10 border border-white/10 bg-gradient-to-r from-purple-600 to-indigo-600 shadow-[0_0_20px_rgba(124,58,237,0.3)]" 
            : "h-6 border border-white/10 bg-gradient-to-r from-cyan-400 to-blue-500 shadow-[0_0_15px_rgba(34,211,238,0.2)]",
          isDragging && "ring-2 ring-white/50 z-[100]",
          isApproved && "from-emerald-500 to-teal-600 opacity-80",
          health === 'OVERDUE' && "from-rose-500 to-red-700 shadow-rose-500/40",
          "hover:backdrop-brightness-110 active:scale-[0.98] cursor-pointer"
        )}
      >
        {/* Subtle Inner Glow */}
        <div className="absolute inset-0 bg-white/5 rounded-lg pointer-events-none" />
        
        <span className={cn(
          "font-black whitespace-nowrap overflow-hidden uppercase tracking-tighter text-white drop-shadow-md",
          isLevel1 || isProjectBar ? "text-[10px]" : "text-[8px]"
        )}>
           {task.title || task.custom_id}
        </span>

        {/* Progress Overlay */}
        {!isApproved && !isProjectBar && (
          <div className="absolute bottom-0 left-0 h-[2px] bg-white/30 rounded-full" style={{ width: `${timeProgress}%` }} />
        )}

        <AnimatePresence>
          {showPopover && (
            <motion.div 
              initial={{ opacity: 0, y: 10, scale: 0.9 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: 10, scale: 0.9 }}
              className="absolute bottom-full mb-3 left-1/2 -translate-x-1/2 w-64 bg-slate-900/95 border border-slate-700 rounded-2xl p-5 shadow-[0_20px_50px_rgba(0,0,0,0.5)] z-[100] backdrop-blur-2xl"
              onClick={(e) => e.stopPropagation()}
            >
              <div className="space-y-4">
                <div className="flex justify-between items-start">
                  <h4 className="text-[11px] font-black text-indigo-400 uppercase tracking-widest">{task.title || task.custom_id}</h4>
                  <HealthBadge health={health} />
                </div>
                
                <div className="grid grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Initialization</p>
                    <p className="text-[10px] font-mono text-slate-200">{format(new Date(fromMs!), 'MMM dd, yyyy')}</p>
                  </div>
                  <div className="space-y-1">
                    <p className="text-[8px] font-black text-slate-500 uppercase tracking-widest">Termination</p>
                    <p className="text-[10px] font-mono text-slate-200">{format(new Date(toMs!), 'MMM dd, yyyy')}</p>
                  </div>
                </div>

                <div className="space-y-2">
                  <div className="flex justify-between text-[8px] font-black uppercase tracking-widest">
                    <span className="text-slate-500">Node Completion</span>
                    <span className="text-indigo-400">{Math.round(timeProgress)}%</span>
                  </div>
                  <div className="h-2 bg-slate-800/50 rounded-full overflow-hidden border border-white/5">
                    <div className="h-full bg-gradient-to-r from-indigo-500 to-violet-500 shadow-[0_0_10px_#6366f1]" style={{ width: `${timeProgress}%` }} />
                  </div>
                </div>

                <div className="pt-3 border-t border-slate-800 flex justify-between items-center">
                  <span className="text-[9px] font-black text-slate-500 uppercase">Inputted Man-Hours</span>
                  <span className="text-[11px] font-black text-indigo-400">
                    {task.man_hours || 0} HOURS
                  </span>
                </div>
              </div>
              <div className="absolute top-full left-1/2 -translate-x-1/2 border-8 border-transparent border-t-slate-900/95" />
            </motion.div>
          )}
        </AnimatePresence>
      </motion.div>
    </div>
  );
}
