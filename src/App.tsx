import { type ReactNode, useEffect, useMemo, useRef, useState } from "react";
import { AlertTriangle, ChevronDown, Circle, Upload } from "lucide-react";
import { Bar, BarChart, Cell, Legend, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { VirtualTable } from "@/components/common/virtual-table";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

type Status = "open" | "inprogress" | "done" | "other";

type Task = {
  taskId: string;
  taskUrl: string;
  issueType: string;
  assignee: string;
  storyPoint: number;
  name: string;
  weeks: string[];
  epicLink: string;
  module: string;
  status: Status;
};

function weekIndex(week: string) {
  const m = week.match(/^W(\d{1,2})$/i);
  return m ? Number(m[1]) : 999;
}

function previousWeek(week: string) {
  const m = week.match(/^W(\d{1,2})$/i);
  if (!m) return "";
  const n = Number(m[1]);
  if (n <= 1) return "";
  return `W${String(n - 1).padStart(2, "0")}`;
}

function statusIndex(status: Status) {
  if (status === "open") return 0;
  if (status === "inprogress") return 1;
  if (status === "done") return 2;
  return 3;
}

function statusLabel(status: Status) {
  if (status === "inprogress") return "in progress";
  return status;
}

function twoDigits(value: number) {
  return String(value).padStart(2, "0");
}

function isoWeekInfo(date: Date) {
  const temp = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = temp.getUTCDay() || 7;
  temp.setUTCDate(temp.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(temp.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((temp.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
  return { week, year: temp.getUTCFullYear() };
}

function weekRangeMonToFri(date: Date) {
  const d = new Date(date);
  const day = d.getDay();
  const mondayDiff = day === 0 ? -6 : 1 - day;
  const monday = new Date(d);
  monday.setDate(d.getDate() + mondayDiff);
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  return { monday, friday };
}

function formatDate(date: Date) {
  return `${twoDigits(date.getDate())}/${twoDigits(date.getMonth() + 1)}/${date.getFullYear()}`;
}

function formatClock(date: Date) {
  return `${twoDigits(date.getHours())}:${twoDigits(date.getMinutes())}:${twoDigits(date.getSeconds())}`;
}

function weekdayName(date: Date) {
  return date.toLocaleDateString("en-US", { weekday: "long" });
}

type ParseWorkerResponse =
  | { type: "success"; requestId: number; rowCount: number; tasks: Task[] }
  | { type: "error"; requestId: number; rowCount: number; error: string };

function sanitizeCell(value: string | number) {
  return String(value ?? "").replace(/\r?\n/g, " ").replace(/\t/g, " ").trim();
}

function escapeHtml(value: string | number) {
  return sanitizeCell(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function displayCount(value: number) {
  return value === 0 ? "-" : String(value);
}

function StatusBadge({ status }: { status: Status | "-" }) {
  if (status === "-") return <span>-</span>;
  const color =
    status === "open"
      ? "text-sky-600"
      : status === "inprogress"
        ? "text-amber-500"
        : status === "done"
          ? "text-emerald-600"
          : "text-slate-500";
  return (
    <span className={`inline-flex items-center gap-1 ${color}`}>
      <Circle className="h-3.5 w-3.5 fill-current" />
      {statusLabel(status)}
    </span>
  );
}

function AccordionSection({
  title,
  description,
  children
}: {
  title: string;
  description: string;
  children: ReactNode;
}) {
  return (
    <details open className="group overflow-hidden rounded-2xl border border-white/50 bg-white/70 shadow-[0_15px_45px_rgba(50,50,93,0.09)] backdrop-blur-xl">
      <summary className="flex cursor-pointer list-none items-center justify-between gap-3 border-b bg-white px-5 py-4">
        <div>
          <div className="text-base font-semibold">{title}</div>
          <div className="text-sm text-muted-foreground">{description}</div>
        </div>
        <ChevronDown className="h-5 w-5 text-muted-foreground transition-transform group-open:rotate-180" />
      </summary>
      <div className="p-5">{children}</div>
    </details>
  );
}

export default function App() {
  const [rowCount, setRowCount] = useState(0);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [error, setError] = useState("");
  const [toast, setToast] = useState<{ message: string; type: "success" | "error" } | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const uploadTokenRef = useRef(0);
  const [now, setNow] = useState(new Date());
  const currentWeekInfo = isoWeekInfo(now);
  const currentWeekCode = `W${twoDigits(currentWeekInfo.week)}`;
  const currentRange = weekRangeMonToFri(now);
  const [activeTab, setActiveTab] = useState<"main" | "manager">("main");

  const [projectAllWeeks, setProjectAllWeeks] = useState(true);
  const [projectWeekFilters, setProjectWeekFilters] = useState<string[]>([]);
  const [projectModuleFilter, setProjectModuleFilter] = useState("all");
  const [projectAssigneeFilter, setProjectAssigneeFilter] = useState("all");
  const [assigneeAllWeeks, setAssigneeAllWeeks] = useState(true);
  const [assigneeWeekFilters, setAssigneeWeekFilters] = useState<string[]>([]);
  const [compareWeekA, setCompareWeekA] = useState("");
  const [compareWeekB, setCompareWeekB] = useState("");
  const [managerWeek, setManagerWeek] = useState("");

  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  const handleFileUpload = async (file: File | null) => {
    if (!file) return;

    const currentToken = ++uploadTokenRef.current;
    setIsParsing(true);

    try {
      setError("");
      setToast(null);
      const data = await file.arrayBuffer();
      const worker = new Worker(new URL("./excel.worker.ts", import.meta.url), { type: "module" });
      let workerResult: ParseWorkerResponse;
      try {
        workerResult = await new Promise<ParseWorkerResponse>((resolve, reject) => {
          const requestId = currentToken;
          worker.onmessage = (event: MessageEvent<ParseWorkerResponse>) => {
            if (event.data.requestId !== requestId) return;
            resolve(event.data);
          };
          worker.onerror = () => reject(new Error("Worker parse failed"));
          worker.postMessage({ type: "parse", requestId, buffer: data }, [data]);
        });
      } finally {
        worker.terminate();
      }

      if (currentToken !== uploadTokenRef.current) return;

      if (workerResult.type === "error") {
        setError(workerResult.error);
        setRowCount(workerResult.rowCount);
        setTasks([]);
        return;
      }

      setRowCount(workerResult.rowCount);
      setTasks(workerResult.tasks);
      setProjectAllWeeks(true);
      setProjectWeekFilters([]);
      setProjectModuleFilter("all");
      setProjectAssigneeFilter("all");
      setAssigneeAllWeeks(true);
      setAssigneeWeekFilters([]);
      const weekSet = new Set<string>();
      workerResult.tasks.forEach((task) => task.weeks.forEach((week) => weekSet.add(week)));
      const sortedWeeks = Array.from(weekSet).sort((a, b) => weekIndex(a) - weekIndex(b));
      setCompareWeekA(sortedWeeks.length >= 2 ? sortedWeeks[sortedWeeks.length - 2] : sortedWeeks[0] ?? "");
      setCompareWeekB(sortedWeeks.length >= 1 ? sortedWeeks[sortedWeeks.length - 1] : "");
      setManagerWeek(sortedWeeks.includes(currentWeekCode) ? currentWeekCode : (sortedWeeks[sortedWeeks.length - 1] ?? ""));
    } catch {
      if (currentToken !== uploadTokenRef.current) return;
      setError("Unable to read file. Please check Excel/CSV format.");
      setRowCount(0);
      setTasks([]);
    } finally {
      if (currentToken === uploadTokenRef.current) setIsParsing(false);
    }
  };

  const allWeeks = useMemo(() => {
    const set = new Set<string>();
    tasks.forEach((task) => {
      task.weeks.forEach((week) => set.add(week));
    });
    return Array.from(set).sort((a, b) => weekIndex(a) - weekIndex(b));
  }, [tasks]);

  const projectModules = useMemo(() => {
    const set = new Set<string>();
    tasks.forEach((task) => {
      const matchWeek =
        projectAllWeeks ||
        projectWeekFilters.length === 0 ||
        task.weeks.some((week) => projectWeekFilters.includes(week));
      if (matchWeek) set.add(task.module);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [tasks, projectAllWeeks, projectWeekFilters]);

  const projectAssignees = useMemo(() => {
    const set = new Set<string>();
    tasks.forEach((task) => {
      const matchWeek =
        projectAllWeeks ||
        projectWeekFilters.length === 0 ||
        task.weeks.some((week) => projectWeekFilters.includes(week));
      const matchModule = projectModuleFilter === "all" || task.module === projectModuleFilter;
      if (matchWeek && matchModule) set.add(task.assignee);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [tasks, projectAllWeeks, projectWeekFilters, projectModuleFilter]);

  const projectWeekSummary = useMemo(() => {
    if (projectAllWeeks) return "All weeks";
    if (!projectWeekFilters.length) return "Select weeks";
    const ordered = allWeeks.filter((week) => projectWeekFilters.includes(week));
    return ordered.join(", ");
  }, [projectAllWeeks, projectWeekFilters, allWeeks]);

  const projectRows = useMemo(() => {
    const taskStatusByWeek = new Map<string, Map<string, Status>>();
    tasks.forEach((task) => {
      const taskKey = task.taskId || `${task.module}||${task.name}`;
      const prev = taskStatusByWeek.get(taskKey) ?? new Map<string, Status>();
      task.weeks.forEach((week) => {
        const current = prev.get(week);
        if (!current || statusIndex(task.status) > statusIndex(current)) {
          prev.set(week, task.status);
        }
      });
      taskStatusByWeek.set(taskKey, prev);
    });

    const selectedWeeks = projectAllWeeks || projectWeekFilters.length === 0 ? allWeeks : projectWeekFilters;
    const filtered = tasks
      .filter((task) => {
        const matchWeek =
          selectedWeeks.length === 0 || task.weeks.some((week) => selectedWeeks.includes(week));
        const matchModule = projectModuleFilter === "all" || task.module === projectModuleFilter;
        const matchAssignee = projectAssigneeFilter === "all" || task.assignee === projectAssigneeFilter;
        return matchWeek && matchModule && matchAssignee;
      })
      .map((task) => ({
        task,
        warning: selectedWeeks.some((week) => {
          const prevWeek = previousWeek(week);
          if (!prevWeek) return false;
          return (taskStatusByWeek.get(task.taskId || `${task.module}||${task.name}`)?.get(prevWeek) ?? "-") === "done";
        })
      }))
      .sort(
        (a, b) =>
          a.task.module.localeCompare(b.task.module) ||
          a.task.assignee.localeCompare(b.task.assignee) ||
          statusIndex(a.task.status) - statusIndex(b.task.status) ||
          a.task.name.localeCompare(b.task.name)
      );

    const moduleCount = new Map<string, number>();
    const assigneeCount = new Map<string, number>();
    filtered.forEach((row) => {
      const task = row.task;
      moduleCount.set(task.module, (moduleCount.get(task.module) ?? 0) + 1);
      const key = `${task.module}||${task.assignee}`;
      assigneeCount.set(key, (assigneeCount.get(key) ?? 0) + 1);
    });

    const seenModule = new Set<string>();
    const seenAssignee = new Set<string>();

    return filtered.map((row) => {
      const task = row.task;
      const assigneeKey = `${task.module}||${task.assignee}`;
      const showModule = !seenModule.has(task.module);
      const showAssignee = !seenAssignee.has(assigneeKey);
      if (showModule) seenModule.add(task.module);
      if (showAssignee) seenAssignee.add(assigneeKey);

      return {
        module: task.module,
        assignee: task.assignee,
        taskKey: task.taskId || `${task.module}||${task.name}`,
        taskId: task.taskId || "-",
        taskUrl: task.taskUrl || "",
        issueType: task.issueType || "-",
        taskName: task.name || "-",
        epicLink: task.epicLink || "-",
        status: task.status,
        storyPoint: task.storyPoint,
        weeks: task.weeks.join(","),
        prevDoneStillAppear: row.warning,
        prevWeek: selectedWeeks.map((w) => previousWeek(w)).filter(Boolean).join(","),
        showModule,
        showAssignee,
        moduleRowSpan: moduleCount.get(task.module) ?? 1,
        assigneeRowSpan: assigneeCount.get(assigneeKey) ?? 1
      };
    });
  }, [tasks, projectAllWeeks, projectWeekFilters, allWeeks, projectModuleFilter, projectAssigneeFilter]);

  const assigneeRows = useMemo(() => {
    const map = new Map<
      string,
      { assignee: string; point1: number; point2: number; point3: number; point4: number; point5: number; total: number }
    >();

    tasks.forEach((task) => {
      const matched =
        assigneeAllWeeks ||
        assigneeWeekFilters.length === 0 ||
        task.weeks.some((week) => assigneeWeekFilters.includes(week));
      if (!matched) return;

      const prev = map.get(task.assignee) ?? {
        assignee: task.assignee,
        point1: 0,
        point2: 0,
        point3: 0,
        point4: 0,
        point5: 0,
        total: 0
      };
      if (task.storyPoint === 1) prev.point1 += 1;
      if (task.storyPoint === 2) prev.point2 += 1;
      if (task.storyPoint === 3) prev.point3 += 1;
      if (task.storyPoint === 4) prev.point4 += 1;
      if (task.storyPoint === 5) prev.point5 += 1;
      prev.total += 1;
      map.set(task.assignee, prev);
    });

    return Array.from(map.values()).sort((a, b) => b.total - a.total || a.assignee.localeCompare(b.assignee));
  }, [tasks, assigneeAllWeeks, assigneeWeekFilters]);

  const assigneeWeekSummary = useMemo(() => {
    if (assigneeAllWeeks) return "All weeks";
    if (!assigneeWeekFilters.length) return "Select weeks";
    const ordered = allWeeks.filter((week) => assigneeWeekFilters.includes(week));
    return ordered.join(", ");
  }, [assigneeAllWeeks, assigneeWeekFilters, allWeeks]);

  useEffect(() => {
    if (!allWeeks.length) {
      setManagerWeek("");
      return;
    }
    if (!managerWeek || !allWeeks.includes(managerWeek)) {
      setManagerWeek(allWeeks.includes(currentWeekCode) ? currentWeekCode : allWeeks[allWeeks.length - 1]);
    }
  }, [allWeeks, managerWeek, currentWeekCode]);

  const compareRows = useMemo(() => {
    if (!compareWeekA || !compareWeekB) return [] as Array<{
      taskKey: string;
      taskId: string;
      taskName: string;
      module: string;
      assignee: string;
      statusA: Status | "-";
      statusB: Status | "-";
      transition: string;
      invalidDoneBoth: boolean;
      missingNextWeekLabelNeedUpdate: boolean;
    }>;

    const taskMap = new Map<
      string,
      { taskId: string; taskName: string; module: string; assignee: string; statusByWeek: Map<string, Status> }
    >();

    tasks.forEach((task) => {
      const taskKey = task.taskId || `${task.module}||${task.name}`;
      const prev = taskMap.get(taskKey) ?? {
        taskId: task.taskId || "-",
        taskName: task.name || "-",
        module: task.module,
        assignee: task.assignee,
        statusByWeek: new Map<string, Status>()
      };

      task.weeks.forEach((week) => {
        const current = prev.statusByWeek.get(week);
        if (!current || statusIndex(task.status) > statusIndex(current)) {
          prev.statusByWeek.set(week, task.status);
        }
      });

      taskMap.set(taskKey, prev);
    });

    return Array.from(taskMap.entries())
      .map(([taskKey, task]) => {
        const statusA = task.statusByWeek.get(compareWeekA) ?? "-";
        const statusB = task.statusByWeek.get(compareWeekB) ?? "-";
        const invalidDoneBoth = statusA === "done" && statusB === "done";
        const missingNextWeekLabelNeedUpdate =
          statusA !== "-" && statusA !== "done" && statusB === "-";
        const transition =
          statusA === "-" && statusB === "-"
            ? "No data in both weeks"
            : statusA === "-"
              ? `New in ${compareWeekB}`
              : statusB === "-"
                ? `Missing in ${compareWeekB}`
                : statusA === statusB
                  ? "No change"
                  : `${statusA} -> ${statusB}`;

        return {
          taskKey,
          taskId: task.taskId,
          taskName: task.taskName,
          module: task.module,
          assignee: task.assignee,
          statusA,
          statusB,
          transition,
          invalidDoneBoth,
          missingNextWeekLabelNeedUpdate
        };
      })
      .filter((row) => row.statusA !== "-" || row.statusB !== "-")
      .sort(
        (a, b) =>
          a.module.localeCompare(b.module) ||
          a.assignee.localeCompare(b.assignee) ||
          Number(b.invalidDoneBoth) - Number(a.invalidDoneBoth) ||
          a.taskId.localeCompare(b.taskId)
      );
  }, [tasks, compareWeekA, compareWeekB]);

  const compareSummary = useMemo(() => {
    const changed = compareRows.filter((row) => row.statusA !== row.statusB).length;
    const invalidDoneBoth = compareRows.filter((row) => row.invalidDoneBoth).length;
    return { changed, invalidDoneBoth, total: compareRows.length };
  }, [compareRows]);

  const managerWeekCode = managerWeek || currentWeekCode;
  const managerPrevWeek = previousWeek(managerWeekCode);

  const managerCurrentWeekTasks = useMemo(
    () => tasks.filter((task) => task.weeks.includes(managerWeekCode)),
    [tasks, managerWeekCode]
  );

  const managerReminderRows = useMemo(() => {
    const taskMap = new Map<
      string,
      { taskId: string; taskName: string; assignee: string; module: string; statusByWeek: Map<string, Status> }
    >();

    tasks.forEach((task) => {
      const key = task.taskId || `${task.module}||${task.name}`;
      const prev =
        taskMap.get(key) ??
        { taskId: task.taskId || "-", taskName: task.name || "-", assignee: task.assignee, module: task.module, statusByWeek: new Map<string, Status>() };

      task.weeks.forEach((week) => {
        const current = prev.statusByWeek.get(week);
        if (!current || statusIndex(task.status) > statusIndex(current)) {
          prev.statusByWeek.set(week, task.status);
        }
      });
      taskMap.set(key, prev);
    });

    return Array.from(taskMap.entries())
      .map(([taskKey, task]) => {
        const prevStatus = managerPrevWeek ? task.statusByWeek.get(managerPrevWeek) ?? "-" : "-";
        const currentStatus = task.statusByWeek.get(managerWeekCode) ?? "-";
        const doneBoth = prevStatus === "done" && currentStatus === "done";
        const missingCurrent = prevStatus !== "-" && prevStatus !== "done" && currentStatus === "-";
        return {
          taskKey,
          taskId: task.taskId,
          taskName: task.taskName,
          assignee: task.assignee,
          module: task.module,
          doneBoth,
          missingCurrent
        };
      })
      .filter((row) => row.doneBoth || row.missingCurrent)
      .sort((a, b) => a.module.localeCompare(b.module) || a.assignee.localeCompare(b.assignee) || a.taskId.localeCompare(b.taskId));
  }, [tasks, managerWeekCode, managerPrevWeek]);

  const managerInProgressMultiWeek = useMemo(() => {
    const map = new Map<string, { taskId: string; taskName: string; assignee: string; module: string; weeks: string }>();
    managerCurrentWeekTasks.forEach((task) => {
      if (task.status !== "inprogress" || task.weeks.length < 2) return;
      const key = task.taskId || `${task.module}-${task.name}`;
      map.set(key, {
        taskId: task.taskId || "-",
        taskName: task.name || "-",
        assignee: task.assignee,
        module: task.module,
        weeks: task.weeks.join(", ")
      });
    });
    return Array.from(map.values()).sort((a, b) => a.assignee.localeCompare(b.assignee));
  }, [managerCurrentWeekTasks]);

  const managerWorkloadRows = useMemo(() => {
    const map = new Map<string, { assignee: string; light: number; medium: number; heavy: number; total: number }>();
    managerCurrentWeekTasks.forEach((task) => {
      const prev = map.get(task.assignee) ?? { assignee: task.assignee, light: 0, medium: 0, heavy: 0, total: 0 };
      if (task.storyPoint <= 2) prev.light += 1;
      else if (task.storyPoint === 3) prev.medium += 1;
      else prev.heavy += 1;
      prev.total += 1;
      map.set(task.assignee, prev);
    });
    return Array.from(map.values()).sort((a, b) => b.total - a.total || a.assignee.localeCompare(b.assignee));
  }, [managerCurrentWeekTasks]);

  const compareGroupedRows = useMemo(() => {
    const moduleCount = new Map<string, number>();
    const assigneeCount = new Map<string, number>();

    compareRows.forEach((row) => {
      moduleCount.set(row.module, (moduleCount.get(row.module) ?? 0) + 1);
      const assigneeKey = `${row.module}||${row.assignee}`;
      assigneeCount.set(assigneeKey, (assigneeCount.get(assigneeKey) ?? 0) + 1);
    });

    const seenModule = new Set<string>();
    const seenAssignee = new Set<string>();

    return compareRows.map((row) => {
      const assigneeKey = `${row.module}||${row.assignee}`;
      const showModule = !seenModule.has(row.module);
      const showAssignee = !seenAssignee.has(assigneeKey);
      if (showModule) seenModule.add(row.module);
      if (showAssignee) seenAssignee.add(assigneeKey);

      return {
        ...row,
        showModule,
        showAssignee,
        moduleRowSpan: moduleCount.get(row.module) ?? 1,
        assigneeRowSpan: assigneeCount.get(assigneeKey) ?? 1
      };
    });
  }, [compareRows]);

  const overviewStatus = useMemo(() => {
    const summary = { open: 0, inprogress: 0, done: 0, other: 0 };
    tasks.forEach((task) => {
      summary[task.status] += 1;
    });
    return summary;
  }, [tasks]);

  const projectPieData = useMemo(() => {
    const summary = { open: 0, inprogress: 0, done: 0, other: 0 };
    projectRows.forEach((row) => {
      summary[row.status] += 1;
    });
    return [
      { name: "open", value: summary.open, color: "#2563eb" },
      { name: "in progress", value: summary.inprogress, color: "#f59e0b" },
      { name: "done", value: summary.done, color: "#16a34a" },
      { name: "other", value: summary.other, color: "#94a3b8" }
    ];
  }, [projectRows]);

  const projectAssigneeChartData = useMemo(() => {
    const map = new Map<string, number>();
    projectRows.forEach((row) => {
      map.set(row.assignee, (map.get(row.assignee) ?? 0) + 1);
    });
    return Array.from(map.entries())
      .map(([assignee, count]) => ({ assignee, count }))
      .sort((a, b) => b.count - a.count || a.assignee.localeCompare(b.assignee))
      .slice(0, 10);
  }, [projectRows]);

  const assigneePieData = useMemo(() => {
    const summary = { C1: 0, C2: 0, C3: 0, C4: 0, C5: 0 };
    assigneeRows.forEach((row) => {
      summary.C1 += row.point1;
      summary.C2 += row.point2;
      summary.C3 += row.point3;
      summary.C4 += row.point4;
      summary.C5 += row.point5;
    });
    return [
      { name: "C1", value: summary.C1, color: "#38bdf8" },
      { name: "C2", value: summary.C2, color: "#818cf8" },
      { name: "C3", value: summary.C3, color: "#f59e0b" },
      { name: "C4", value: summary.C4, color: "#f97316" },
      { name: "C5", value: summary.C5, color: "#ef4444" }
    ];
  }, [assigneeRows]);

  const comparePieData = useMemo(
    () => [
      { name: "Status changed", value: compareSummary.changed, color: "#0ea5e9" },
      { name: "No change", value: Math.max(compareSummary.total - compareSummary.changed, 0), color: "#64748b" },
      { name: "Done in both", value: compareSummary.invalidDoneBoth, color: "#dc2626" }
    ],
    [compareSummary]
  );

  const showToast = (message: string, type: "success" | "error") => {
    setToast({ message, type });
    window.setTimeout(() => setToast(null), 2200);
  };

  const copyCsv = async (kind: "project" | "assignee") => {
    try {
      if (kind === "project") {
        const plainText = projectRows
          .map((r) =>
            [
              r.showModule ? r.module : "",
              r.showAssignee ? r.assignee : "",
              r.taskId,
              r.taskName,
              statusLabel(r.status),
              r.storyPoint
            ]
              .map(sanitizeCell)
              .join("\t")
          )
          .join("\n");

        const htmlRows = projectRows
          .map((r) => {
            const moduleCell = r.showModule ? `<td rowspan="${r.moduleRowSpan}">${escapeHtml(r.module)}</td>` : "";
            const assigneeCell = r.showAssignee ? `<td rowspan="${r.assigneeRowSpan}">${escapeHtml(r.assignee)}</td>` : "";
            return `<tr>${moduleCell}${assigneeCell}<td>${escapeHtml(r.taskId)}</td><td>${escapeHtml(r.taskName)}</td><td>${escapeHtml(
              statusLabel(r.status)
            )}</td><td>${escapeHtml(r.storyPoint)}</td></tr>`;
          })
          .join("");
        const htmlTable = `<table><tbody>${htmlRows}</tbody></table>`;

        if (navigator.clipboard.write && typeof ClipboardItem !== "undefined") {
          await navigator.clipboard.write([
            new ClipboardItem({
              "text/plain": new Blob([plainText], { type: "text/plain" }),
              "text/html": new Blob([htmlTable], { type: "text/html" })
            })
          ]);
        } else {
          await navigator.clipboard.writeText(plainText);
        }

        showToast("Table copied: project view", "success");
        return;
      }

      if (kind === "assignee") {
        const plainText = assigneeRows
          .map((r) => [r.assignee, r.point1, r.point2, r.point3, r.point4, r.point5].map(sanitizeCell).join("\t"))
          .join("\n");
        await navigator.clipboard.writeText(plainText);
        showToast("Table copied: assignee view", "success");
        return;
      }
    } catch {
      showToast("Copy failed. Please check clipboard permission.", "error");
    }
  };

  const totalTasks = tasks.length;
  const withWeek = tasks.filter((task) => task.weeks.length > 0).length;
  const hasData = rowCount > 0;

  return (
    <main className="mx-auto max-w-[1300px] p-4 md:p-8">
      <section className="mb-6 grid gap-3 md:grid-cols-3">
        <div className="rounded-2xl border border-cyan-200/70 bg-gradient-to-br from-cyan-500 via-blue-600 to-indigo-700 p-4 text-white shadow-[0_20px_45px_rgba(37,99,235,0.26)]">
          <div className="text-[11px] uppercase tracking-[0.18em] text-cyan-100">Current Week</div>
          <div className="mt-1 text-3xl font-bold">{currentWeekCode}</div>
          <div className="mt-1 text-xs text-cyan-100">
            {formatDate(currentRange.monday)} - {formatDate(currentRange.friday)}
          </div>
          <div className="mt-2 -ml-4 gap-1 flex w-[274px] items-center justify-start rounded-full rounded-l-none bg-white/20 px-3 py-1.5 text-base font-semibold tabular-nums">
            <span className="">{weekdayName(now)}</span>
            <span className="opacity-70">|</span>
            <span className="">{formatDate(now)}</span>
            <span className="opacity-70">|</span>
            <span className="">{formatClock(now)}</span>
          </div>
        </div>
        <Card className="border-0 bg-gradient-to-br from-indigo-500 to-violet-600 text-white shadow-lg">
          <CardHeader className="p-4">
            <CardDescription className="text-indigo-100">Task Summary</CardDescription>
            <CardTitle className="text-white text-3xl">{totalTasks}</CardTitle>
            <p className="text-sm text-indigo-100">Valid week labels: {withWeek}</p>
          </CardHeader>
        </Card>
        <Card className="border-0 bg-gradient-to-br from-slate-600 to-slate-800 text-white shadow-lg">
          <CardHeader className="p-4">
            <CardDescription className="text-slate-200">Status Summary</CardDescription>
            <div className="mt-1 grid grid-cols-2 gap-2 text-sm">
              <div>Open: <span className="font-semibold">{overviewStatus.open}</span></div>
              <div>In Progress: <span className="font-semibold">{overviewStatus.inprogress}</span></div>
              <div>Done: <span className="font-semibold">{overviewStatus.done}</span></div>
              <div>Other: <span className="font-semibold">{overviewStatus.other}</span></div>
            </div>
          </CardHeader>
        </Card>
      </section>

      <section className="mb-6 rounded-xl border border-white/60 bg-white/80 p-3 shadow-[0_8px_26px_rgba(15,23,42,0.08)] backdrop-blur-xl">
        <details open={!hasData}>
          <summary className="flex cursor-pointer list-none items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <Upload className="h-4 w-4 text-primary" />
              <h1 className="text-sm font-semibold tracking-tight">Data Input</h1>
            </div>
            <span className="text-xs text-muted-foreground">
              {isParsing ? "Parsing file in background..." : hasData ? "Data loaded - click to upload new file" : "Upload file to start"}
            </span>
          </summary>
          <div className="mt-3 grid gap-2 sm:grid-cols-[220px,1fr] sm:items-center">
            <Label htmlFor="file" className="text-xs text-muted-foreground">Excel/CSV File</Label>
            <Input
              id="file"
              type="file"
              accept=".xlsx,.xls,.csv"
              disabled={isParsing}
              onChange={(e) => handleFileUpload(e.target.files?.[0] ?? null)}
              className="h-9 text-sm"
            />
          </div>
        </details>
        {isParsing && <p className="mt-3 text-sm text-primary">Processing data in Web Worker...</p>}
        {error && <p className="mt-3 text-sm text-red-600">{error}</p>}
      </section>

      <section className="mb-5 flex gap-2">
        <button
          type="button"
          onClick={() => setActiveTab("main")}
          className={`rounded-full px-4 py-2 text-sm font-medium ${activeTab === "main" ? "bg-primary text-white" : "bg-white text-slate-700 border"}`}
        >
          Main Dashboard
        </button>
        <button
          type="button"
          onClick={() => setActiveTab("manager")}
          className={`rounded-full px-4 py-2 text-sm font-medium ${activeTab === "manager" ? "bg-primary text-white" : "bg-white text-slate-700 border"}`}
        >
          Manager View
        </button>
      </section>

      {activeTab === "main" ? (
      <section className="space-y-5">
        <AccordionSection
          title="1. Project Filter by Week"
          description="Grouped by project -> assignee -> task (rowspan/colspan)"
        >
          <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
            <div className="grid w-full max-w-4xl grid-cols-1 gap-3 sm:grid-cols-3">
              <div className="space-y-2">
                <Label>Weeks (multi-select)</Label>
                <details className="group relative">
                  <summary className="flex h-10 cursor-pointer list-none items-center justify-between rounded-md border border-input bg-background px-3 py-2 text-sm">
                    <span className="truncate">{projectWeekSummary}</span>
                  </summary>
                  <div className="absolute z-20 mt-2 w-full rounded-md border bg-card p-3 shadow-md">
                    <label className="mb-2 flex items-center gap-2 text-sm">
                      <input
                        type="checkbox"
                        checked={projectAllWeeks}
                        onChange={(e) => {
                          if (e.target.checked) {
                            setProjectAllWeeks(true);
                            setProjectWeekFilters([]);
                          } else {
                            setProjectAllWeeks(false);
                          }
                        }}
                      />
                      All weeks
                    </label>
                    <div className="max-h-44 space-y-1 overflow-auto border-t pt-2">
                      {allWeeks.map((week) => (
                        <label key={week} className="flex items-center gap-2 text-sm">
                          <input
                            type="checkbox"
                            checked={projectWeekFilters.includes(week)}
                            onChange={(e) => {
                              setProjectAllWeeks(false);
                              setProjectWeekFilters((prev) =>
                                e.target.checked ? [...prev, week] : prev.filter((w) => w !== week)
                              );
                            }}
                          />
                          {week}
                        </label>
                      ))}
                    </div>
                  </div>
                </details>
              </div>
              <div className="space-y-2">
                <Label>Project (optional)</Label>
                <Select
                  value={projectModuleFilter}
                  onValueChange={(value) => {
                    setProjectModuleFilter(value);
                    setProjectAssigneeFilter("all");
                  }}
                >
                  <SelectTrigger>
                    <SelectValue placeholder="All projects" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All projects</SelectItem>
                    {projectModules.map((module) => (
                      <SelectItem key={module} value={module}>
                        {module}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-2">
                <Label>Assignee (optional)</Label>
                <Select value={projectAssigneeFilter} onValueChange={setProjectAssigneeFilter}>
                  <SelectTrigger>
                    <SelectValue placeholder="All assignees" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="all">All assignees</SelectItem>
                    {projectAssignees.map((assignee) => (
                      <SelectItem key={assignee} value={assignee}>
                        {assignee}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
            <Button variant="outline" onClick={() => copyCsv("project")}>Copy CSV</Button>
          </div>

          <div className="mb-4 grid gap-4 lg:grid-cols-[320px,1fr]">
            <div className="h-56 rounded-xl border border-white/70 bg-white p-2 shadow-sm">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={projectPieData} dataKey="value" nameKey="name" outerRadius={78} innerRadius={40}>
                    {projectPieData.map((entry) => (
                      <Cell key={`project-${entry.name}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="h-56 rounded-xl border border-white/70 bg-white p-2 shadow-sm">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={projectAssigneeChartData}>
                  <XAxis dataKey="assignee" tick={{ fontSize: 10 }} interval={0} angle={-25} textAnchor="end" height={70} />
                  <YAxis allowDecimals={false} />
                  <Tooltip />
                  <Bar dataKey="count" fill="#0ea5e9" name="Task count" radius={[4, 4, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>

          <VirtualTable
            rows={projectRows}
            height={560}
            getRowKey={(row, index) => `${row.module}-${row.assignee}-${row.taskKey}-${index}`}
            columns={[
              {
                key: "module",
                label: "Module Name",
                cellClassName: "p-2 align-top font-medium",
                render: (row) => (row.showModule ? row.module : "")
              },
              {
                key: "assignee",
                label: "Assignee",
                cellClassName: "p-2 align-top",
                render: (row) => (row.showAssignee ? row.assignee : "")
              },
              {
                key: "taskId",
                label: "Task ID",
                render: (row) =>
                  row.taskUrl ? (
                    <a href={row.taskUrl} target="_blank" rel="noreferrer" className="text-primary underline underline-offset-2">
                      {row.taskId}
                    </a>
                  ) : (
                    row.taskId
                  )
              },
              { key: "taskName", label: "Task Name", render: (row) => row.taskName },
              { key: "status", label: "Status", render: (row) => <StatusBadge status={row.status} /> },
              { key: "storyPoint", label: "Story Point", render: (row) => row.storyPoint },
              { key: "weeks", label: "Weeks", render: (row) => row.weeks },
              {
                key: "warning",
                label: "Warning",
                render: (row) =>
                  row.prevDoneStillAppear ? (
                    <span
                      title={`Task was done in previous week(s) (${row.prevWeek}) but still appears in selected weeks`}
                      className="inline-flex items-center gap-1 text-amber-600"
                    >
                      <AlertTriangle className="h-4 w-4" />
                      done last week
                    </span>
                  ) : (
                    "-"
                  )
              }
            ]}
          />
        </AccordionSection>

        <AccordionSection
          title="2. Assignee Workload by Difficulty"
          description="Select multiple weeks or all weeks to count tasks by difficulty per assignee"
        >
          <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
            <div className="w-full max-w-xs space-y-2">
              <Label>Weeks (multi-select)</Label>
              <details className="group relative">
                <summary className="flex h-10 cursor-pointer list-none items-center justify-between rounded-md border border-input bg-background px-3 py-2 text-sm">
                  <span className="truncate">{assigneeWeekSummary}</span>
                </summary>
                <div className="absolute z-20 mt-2 w-full rounded-md border bg-card p-3 shadow-md">
                  <label className="mb-2 flex items-center gap-2 text-sm">
                    <input
                      type="checkbox"
                      checked={assigneeAllWeeks}
                      onChange={(e) => {
                        if (e.target.checked) {
                          setAssigneeAllWeeks(true);
                          setAssigneeWeekFilters([]);
                        } else {
                          setAssigneeAllWeeks(false);
                        }
                      }}
                    />
                    All weeks
                  </label>
                  <div className="max-h-44 space-y-1 overflow-auto border-t pt-2">
                    {allWeeks.map((week) => (
                      <label key={week} className="flex items-center gap-2 text-sm">
                        <input
                          type="checkbox"
                          checked={assigneeWeekFilters.includes(week)}
                          onChange={(e) => {
                            setAssigneeAllWeeks(false);
                            setAssigneeWeekFilters((prev) => (e.target.checked ? [...prev, week] : prev.filter((w) => w !== week)));
                          }}
                        />
                        {week}
                      </label>
                    ))}
                  </div>
                </div>
              </details>
            </div>
            <Button variant="outline" onClick={() => copyCsv("assignee")}>Copy CSV</Button>
          </div>

          <div className="grid gap-4 lg:grid-cols-[260px,1fr]">
            <div className="h-64 rounded-xl border border-white/70 bg-white p-2 shadow-sm">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={assigneePieData} dataKey="value" nameKey="name" outerRadius={78} innerRadius={40}>
                    {assigneePieData.map((entry) => (
                      <Cell key={`assignee-${entry.name}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>

            <VirtualTable
              rows={assigneeRows}
              height={460}
              getRowKey={(row) => row.assignee}
              columns={[
                { key: "assignee", label: "Assignee", render: (row) => row.assignee },
                { key: "point1", label: "C1", render: (row) => displayCount(row.point1) },
                { key: "point2", label: "C2", render: (row) => displayCount(row.point2) },
                { key: "point3", label: "C3", render: (row) => displayCount(row.point3) },
                { key: "point4", label: "C4", render: (row) => displayCount(row.point4) },
                { key: "point5", label: "C5", render: (row) => displayCount(row.point5) },
                { key: "total", label: "Total", render: (row) => displayCount(row.total) }
              ]}
            />
          </div>
        </AccordionSection>

        <AccordionSection
          title="3. Status Change Between 2 Weeks"
          description="Compare task status by Task ID across 2 nearby weeks and detect anomalies"
        >
          <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
            <div className="grid w-full max-w-xl grid-cols-1 gap-3 sm:grid-cols-2">
              <div className="space-y-2">
                <Label>Week A</Label>
                <Select value={compareWeekA || "__none__"} onValueChange={(v) => setCompareWeekA(v === "__none__" ? "" : v)}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select week A" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">Not selected</SelectItem>
                    {allWeeks.map((week) => (
                      <SelectItem key={`a-${week}`} value={week}>
                        {week}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
              <div className="space-y-2">
                <Label>Week B</Label>
                <Select value={compareWeekB || "__none__"} onValueChange={(v) => setCompareWeekB(v === "__none__" ? "" : v)}>
                  <SelectTrigger>
                    <SelectValue placeholder="Select week B" />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="__none__">Not selected</SelectItem>
                    {allWeeks.map((week) => (
                      <SelectItem key={`b-${week}`} value={week}>
                        {week}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              </div>
            </div>
          </div>

          <div className="mb-3 text-sm text-muted-foreground">
            Total compared tasks: {compareSummary.total} | Changed tasks: {compareSummary.changed} | Done-done warnings: {compareSummary.invalidDoneBoth}
          </div>

          <div className="mb-5 grid gap-4 lg:grid-cols-[260px,1fr]">
            <div className="h-64 rounded-xl border border-white/70 bg-white p-2 shadow-sm">
              <ResponsiveContainer width="100%" height="100%">
                <PieChart>
                  <Pie data={comparePieData} dataKey="value" nameKey="name" outerRadius={78} innerRadius={40}>
                    {comparePieData.map((entry) => (
                      <Cell key={`compare-${entry.name}`} fill={entry.color} />
                    ))}
                  </Pie>
                  <Tooltip />
                  <Legend />
                </PieChart>
              </ResponsiveContainer>
            </div>
            <div className="grid content-start gap-2 rounded-xl border border-white/70 bg-white p-4 text-sm shadow-sm">
              <div className="font-medium">Comparison Summary</div>
              <div>Status changed: <span className="font-semibold">{compareSummary.changed}</span></div>
              <div>No change: <span className="font-semibold">{Math.max(compareSummary.total - compareSummary.changed, 0)}</span></div>
              <div>Done in both weeks: <span className="font-semibold text-red-600">{compareSummary.invalidDoneBoth}</span></div>
            </div>
          </div>

          {compareWeekA && compareWeekB && compareWeekA === compareWeekB && (
            <p className="mb-3 text-sm text-amber-700">Please select 2 different weeks for a valid comparison.</p>
          )}

          <VirtualTable
            rows={compareGroupedRows}
            height={520}
            getRowKey={(row, index) => `${row.taskKey}-${index}`}
            columns={[
              {
                key: "module",
                label: "Module",
                cellClassName: "p-2 align-top font-medium",
                render: (row) => (row.showModule ? row.module : "")
              },
              {
                key: "assignee",
                label: "Assignee",
                cellClassName: "p-2 align-top",
                render: (row) => (row.showAssignee ? row.assignee : "")
              },
              { key: "taskId", label: "Task ID", render: (row) => row.taskId },
              { key: "taskName", label: "Task Name", render: (row) => row.taskName },
              { key: "statusA", label: `Status ${compareWeekA || "A"}`, render: (row) => <StatusBadge status={row.statusA} /> },
              { key: "statusB", label: `Status ${compareWeekB || "B"}`, render: (row) => <StatusBadge status={row.statusB} /> },
              { key: "transition", label: "Transition", render: (row) => row.transition },
              {
                key: "warning",
                label: "Warning",
                render: (row) =>
                  row.invalidDoneBoth
                    ? "Task done in both weeks (invalid)"
                    : row.missingNextWeekLabelNeedUpdate
                      ? "Not done, missing next-week label"
                      : "-"
              }
            ]}
          />

          {!compareRows.length && <div className="py-5 text-center text-sm text-muted-foreground">No comparison data for selected weeks.</div>}
        </AccordionSection>
      </section>
      ) : (
      <section className="space-y-5">
        <div className="flex flex-wrap items-end justify-between gap-3 rounded-xl border border-white/70 bg-white p-4 shadow-sm">
          <div>
            <div className="text-sm font-medium">Manager week scope</div>
            <div className="text-xs text-muted-foreground">All manager metrics below follow this selected week.</div>
          </div>
          <div className="w-full max-w-xs space-y-1">
            <Label>Week</Label>
            <Select value={managerWeekCode || "__none__"} onValueChange={(v) => setManagerWeek(v === "__none__" ? "" : v)}>
              <SelectTrigger>
                <SelectValue placeholder="Select week" />
              </SelectTrigger>
              <SelectContent>
                {allWeeks.length === 0 && <SelectItem value="__none__">No week data</SelectItem>}
                {allWeeks.map((week) => (
                  <SelectItem key={`manager-${week}`} value={week}>
                    {week}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        </div>

        <div className="grid gap-4 sm:grid-cols-3">
          <Card className="border-0 bg-gradient-to-br from-rose-500 to-red-600 text-white">
            <CardHeader>
              <CardDescription className="text-rose-100">Need Follow-up ({managerWeekCode || "N/A"})</CardDescription>
              <CardTitle className="text-white">{managerReminderRows.length}</CardTitle>
            </CardHeader>
          </Card>
          <Card className="border-0 bg-gradient-to-br from-sky-500 to-blue-600 text-white">
            <CardHeader>
              <CardDescription className="text-sky-100">In Progress 2+ weeks ({managerWeekCode || "N/A"})</CardDescription>
              <CardTitle className="text-white">{managerInProgressMultiWeek.length}</CardTitle>
            </CardHeader>
          </Card>
          <Card className="border-0 bg-gradient-to-br from-violet-500 to-indigo-600 text-white">
            <CardHeader>
              <CardDescription className="text-violet-100">Tasks in Current Week</CardDescription>
              <CardTitle className="text-white">{managerCurrentWeekTasks.length}</CardTitle>
            </CardHeader>
          </Card>
        </div>

        <AccordionSection
          title="Follow-up Board"
          description={`Warnings using previous week (${managerPrevWeek || "N/A"}) vs selected week (${managerWeekCode || "N/A"})`}
        >
          <div className="max-h-[420px] overflow-y-auto rounded-xl border border-white/70 bg-white shadow-sm">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Task ID</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Assignee</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Module</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Warning</th>
                </tr>
              </thead>
              <tbody>
                {managerReminderRows.map((row) => (
                  <tr key={`reminder-${row.taskKey}`} className="border-b hover:bg-slate-50">
                    <td className="p-2">{row.taskId}</td>
                    <td className="p-2">{row.assignee}</td>
                    <td className="p-2">{row.module}</td>
                    <td className="p-2">
                      {row.doneBoth ? "Task done in both weeks (invalid)" : "Not done, missing current-week label"}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </AccordionSection>

        <AccordionSection
          title="Workload by Assignee"
          description="Task load split by difficulty group"
        >
          <div className="max-h-[420px] overflow-y-auto rounded-xl border border-white/70 bg-white shadow-sm">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Assignee</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Light (C1-C2)</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Medium (C3)</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Heavy (C4-C5)</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Total</th>
                </tr>
              </thead>
              <tbody>
                {managerWorkloadRows.map((row) => (
                  <tr key={`work-${row.assignee}`} className="border-b hover:bg-slate-50">
                    <td className="p-2">{row.assignee}</td>
                    <td className="p-2">{displayCount(row.light)}</td>
                    <td className="p-2">{displayCount(row.medium)}</td>
                    <td className="p-2">{displayCount(row.heavy)}</td>
                    <td className="p-2">{displayCount(row.total)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </AccordionSection>

        <AccordionSection
          title="In Progress Over Multiple Weeks"
          description="Potential stalled tasks for direct reminder"
        >
          <div className="max-h-[420px] overflow-y-auto rounded-xl border border-white/70 bg-white shadow-sm">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Task ID</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Task</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Assignee</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Module</th>
                  <th className="sticky top-0 z-20 bg-slate-100 px-2 py-2 text-left font-medium text-slate-700">Weeks</th>
                </tr>
              </thead>
              <tbody>
                {managerInProgressMultiWeek.map((row) => (
                  <tr key={`ip-${row.taskId}-${row.taskName}`} className="border-b hover:bg-slate-50">
                    <td className="p-2">{row.taskId}</td>
                    <td className="p-2">{row.taskName}</td>
                    <td className="p-2">{row.assignee}</td>
                    <td className="p-2">{row.module}</td>
                    <td className="p-2">{row.weeks}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </AccordionSection>
      </section>
      )}

      {!rowCount && !error && <div className="py-8 text-center text-sm text-muted-foreground">Upload a file to start analysis.</div>}

      {toast && (
        <div className={`fixed bottom-6 right-6 z-[100] rounded-lg px-4 py-2 text-sm text-white shadow-lg ${toast.type === "success" ? "bg-emerald-600" : "bg-red-600"}`}>
          {toast.message}
        </div>
      )}
    </main>
  );
}
