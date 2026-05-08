import { 
  addDays, 
  addWeeks, 
  addMonths, 
  subDays, 
  startOfDay, 
  endOfDay, 
  eachHourOfInterval, 
  eachDayOfInterval, 
  eachWeekOfInterval, 
  eachMonthOfInterval,
  startOfMonth,
  endOfMonth as dateFnsEndOfMonth,
  startOfWeek,
  endOfWeek,
  addHours
} from 'date-fns';
import { Task, Project } from '../types';

export interface GanttTimelineRange {
  gridStart: Date;
  gridEnd: Date;
  intervals: Date[];
  totalDuration: number;
}

export const sanitizeDate = (value: any): string | null => {
  if (!value || value === '') return null;
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : (typeof value === 'string' ? value : date.toISOString());
};

export const normalizeDate = (value: any): Date | null => {
  const sanitized = sanitizeDate(value);
  return sanitized ? new Date(sanitized) : null;
};

export const calculateParentRange = (children: Task[]): { from_date: string | null; to_date: string | null } => {
  if (!children || children.length === 0) return { from_date: null, to_date: null };

  const startDates = children
    .map(t => normalizeDate(t.start_time))
    .filter((d): d is Date => d !== null)
    .map(d => d.getTime());
  
  const endDates = children
    .map(t => normalizeDate(t.end_time))
    .filter((d): d is Date => d !== null)
    .map(d => d.getTime());

  if (startDates.length === 0) return { from_date: null, to_date: null };

  return {
    from_date: new Date(Math.min(...startDates)).toISOString(),
    to_date: new Date(Math.max(...endDates)).toISOString()
  };
};

export const calculateTimelineRange = (
  scale: 'HOUR' | 'DAY' | 'WEEK' | 'MONTH',
  tasks: Task[],
  projects: Project[],
  projectId?: string | null,
  isGlobalView?: boolean
): GanttTimelineRange => {
  let minTs = Date.now();
  let maxTs = Date.now() + 86400000 * 7;
  
  const project = projectId ? projects.find(p => p.id === projectId) : null;
  const projectTasks = projectId ? (tasks || []).filter(t => t.project_id === projectId) : (tasks || []);

  // THE "MATHEMATICAL ANCHOR" PROTOCOL:
  // 1. Identify all relevant dates (tasks + project boundaries)
  const taskDates = (projectTasks || []).flatMap(t => {
    const s = t.start_time || (t as any).start_date;
    const e = t.end_time || (t as any).end_date;
    return [
      s ? new Date(s).getTime() : null,
      e ? new Date(e).getTime() : null
    ];
  });

  const pDates = (isGlobalView ? (projects || []) : (project ? [project] : [])).flatMap(p => [
    p.start_date ? new Date(p.start_date).getTime() : null,
    p.end_date ? new Date(p.end_date).getTime() : null
  ]);

  const validDates = [...taskDates, ...pDates].filter((t): t is number => t !== null && !isNaN(t));

  if (validDates.length > 0) {
    minTs = Math.min(...validDates);
    maxTs = Math.max(...validDates);
  } else if (project && project.start_date) {
    minTs = new Date(project.start_date).getTime();
    maxTs = project.end_date ? new Date(project.end_date).getTime() : minTs + 86400000 * 7;
  }

  // Fallback for completely empty state
  if (isNaN(minTs) || isNaN(maxTs)) {
    minTs = Date.now();
    maxTs = Date.now() + 86400000 * 7;
  }

  // THE "MATHEMATICAL ANCHOR" PROTOCOL: 
  // No buffer at the start to ensure 0% alignment with the absolute earliest date
  const start = new Date(minTs);
  const bufferTail = scale === 'HOUR' ? 0 : (scale === 'MONTH' ? 30 : 5);
  const end = addDays(new Date(maxTs), isGlobalView ? 10 : bufferTail);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) {
     return {
        gridStart: new Date(),
        gridEnd: addDays(new Date(), 7),
        intervals: [],
        totalDuration: 86400000 * 7
      };
  }

  let generatedIntervals: Date[] = [];
  const weekOpts = { weekStartsOn: 1 as 0 | 1 | 2 | 3 | 4 | 5 | 6 };

  switch(scale) {
    case 'HOUR':
      generatedIntervals = eachHourOfInterval({ start: startOfDay(start), end: endOfDay(end) });
      break;
    case 'DAY':
      generatedIntervals = eachDayOfInterval({ start: startOfDay(start), end: endOfDay(end) });
      break;
    case 'WEEK':
      generatedIntervals = eachWeekOfInterval({ 
        start: startOfWeek(start, weekOpts), 
        end: endOfWeek(end, weekOpts) 
      });
      break;
    case 'MONTH':
      generatedIntervals = eachMonthOfInterval({ 
        start: startOfMonth(start), 
        end: dateFnsEndOfMonth(end) 
      });
      break;
  }

  if (generatedIntervals.length === 0) {
    generatedIntervals = [startOfDay(start)];
  }

  const firstStart = generatedIntervals[0];
  const lastInterval = generatedIntervals[generatedIntervals.length - 1];
  let gridEnd: Date;
  
  switch(scale) {
    case 'HOUR': gridEnd = addHours(lastInterval, 1); break;
    case 'DAY': gridEnd = addDays(lastInterval, 1); break;
    case 'WEEK': gridEnd = addWeeks(lastInterval, 1); break;
    case 'MONTH': gridEnd = addMonths(lastInterval, 1); break;
    default: gridEnd = lastInterval;
  }

  return {
    gridStart: firstStart,
    gridEnd: gridEnd,
    intervals: generatedIntervals,
    totalDuration: Math.max(3600000, gridEnd.getTime() - firstStart.getTime())
  };
};

export const getDaysDiff = (d1: Date | string, d2: Date | string): number => {
  const start = startOfDay(new Date(d1));
  const end = startOfDay(new Date(d2));
  return Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24));
};

export const calculateBarCoordinates = (
  taskStart: string | null | undefined,
  taskEnd: string | null | undefined,
  timelineRange: GanttTimelineRange,
  scale: 'HOUR' | 'DAY' | 'WEEK' | 'MONTH'
) => {
  const start = normalizeDate(taskStart);
  const end = normalizeDate(taskEnd);

  if (!start || !end || !timelineRange.gridStart) {
    return { left: 0, width: 0, isVisible: false };
  }

  const gridStartMs = timelineRange.gridStart.getTime();
  const taskStartMs = start.getTime();
  const taskEndMs = end.getTime();
  
  const CELL_WIDTH = scale === 'HOUR' ? 80 : 100;

  // Calculate milliseconds per cell based on scale
  let msPerCell = 86400000; // Default DAY
  if (scale === 'HOUR') msPerCell = 3600000;
  if (scale === 'WEEK') msPerCell = 86400000 * 7;
  if (scale === 'MONTH') {
     // Month is tricky because it's variable. 
     // We should use the actual month boundaries if we want perfection.
     // For now, let's stick to percentage or a robust month approximation.
     // But the user wants "Mathematical Anchor".
  }

  // REFINED PIXEL LOGIC: 
  // We use percentages relative to totalDuration * gridWidth 
  // OR we calculate absolute pixels. 
  // Percentages are fine if the parent container is correctly sized to gridWidth.
  const leftPercent = ((taskStartMs - gridStartMs) / timelineRange.totalDuration) * 100;
  const widthPercent = ((taskEndMs - taskStartMs) / timelineRange.totalDuration) * 100;

  let left = leftPercent;
  let width = widthPercent;

  // Clamping for visualization if it overflows the grid range
  if (left < 0) {
    width = Math.max(0, width + left);
    left = 0;
  }
  
  if (left + width > 100) {
    width = Math.max(0, 100 - left);
  }

  // Minimum visible width (1px equivalent in percent)
  width = Math.max(0.1, width);
  
  const isVisible = (taskEndMs > gridStartMs) && (taskStartMs < timelineRange.gridEnd.getTime());

  return {
    left,
    width,
    isVisible
  };
};
