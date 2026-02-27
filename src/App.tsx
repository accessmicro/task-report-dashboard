import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { AlertTriangle, Circle, Upload } from "lucide-react";
import { Bar, BarChart, Cell, Legend, Pie, PieChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";

type RawRow = Record<string, string | number>;
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

const COLUMN_ALIASES = {
  taskId: ["key", "taskid", "task id", "id", "ticketid", "jiraid"],
  issueType: ["issuetype", "issue type", "type"],
  assignee: ["assignee"],
  storyPoint: ["storypoints", "story points", "storypoint", "story point"],
  name: ["summary", "name", "taskname", "task name"],
  labels: ["labels", "label"],
  epicLink: ["epiclink", "epic link", "epic"],
  module: ["modulename", "module name", "module", "project", "projectname"],
  status: ["status", "state"]
} as const;

function normalizeText(value: string) {
  return value.toLowerCase().replace(/[_\s-]+/g, "").trim();
}

function detectColumn(columns: string[], aliases: readonly string[]) {
  const normalizedAliases = aliases.map(normalizeText);
  return columns.find((column) => normalizedAliases.includes(normalizeText(column))) ?? "";
}

function normalizeStatus(value: string | number | undefined): Status {
  const v = String(value ?? "").toLowerCase().replace(/\s+/g, "");
  if (v === "open" || v === "todo" || v === "to-do" || v === "new" || v === "backlog") return "open";
  if (v === "inprogress" || v === "in-progress" || v === "doing" || v === "progress") return "inprogress";
  if (v === "done" || v === "closed") return "done";
  return "other";
}

function parseWeekLabels(value: string | number | undefined) {
  const raw = String(value ?? "").replace(/[\[\]"]/g, " ");
  const weeks = raw
    .split(/[;,|\s]+/)
    .map((item) => item.trim().toUpperCase())
    .filter((item) => /^W\d{1,2}$/.test(item));
  return Array.from(new Set(weeks));
}

function storyPointValue(value: string | number | undefined) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(1, Math.min(5, Math.round(n)));
}

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

function parseTaskKeyCell(value: string | number | undefined) {
  const raw = String(value ?? "").trim();
  const markdown = raw.match(/^\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)$/i);
  if (markdown) return { taskId: markdown[1], taskUrl: markdown[2] };

  const urlMatch = raw.match(/https?:\/\/[^\s|,;]+/i);
  const keyMatch = raw.match(/[A-Z][A-Z0-9]+-\d+/i);

  if (urlMatch) {
    const taskUrl = urlMatch[0];
    const keyFromUrl = taskUrl.match(/[A-Z][A-Z0-9]+-\d+/i)?.[0];
    return { taskId: (keyMatch?.[0] || keyFromUrl || raw).toUpperCase(), taskUrl };
  }

  if (keyMatch) return { taskId: keyMatch[0].toUpperCase(), taskUrl: "" };
  return { taskId: raw || "-", taskUrl: "" };
}

function escapeCsv(value: string | number) {
  const text = String(value ?? "");
  if (text.includes(",") || text.includes("\n") || text.includes("\"")) {
    return `"${text.replace(/\"/g, '""')}"`;
  }
  return text;
}

function toCsv(headers: string[], rows: Array<Array<string | number>>) {
  const headerLine = headers.map(escapeCsv).join(",");
  const rowLines = rows.map((row) => row.map(escapeCsv).join(","));
  return [headerLine, ...rowLines].join("\n");
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

export default function App() {
  const [rawRows, setRawRows] = useState<RawRow[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState("");
  const [now, setNow] = useState(new Date());

  const [projectWeekFilter, setProjectWeekFilter] = useState("all");
  const [projectModuleFilter, setProjectModuleFilter] = useState("all");
  const [projectAssigneeFilter, setProjectAssigneeFilter] = useState("all");
  const [assigneeAllWeeks, setAssigneeAllWeeks] = useState(true);
  const [assigneeWeekFilters, setAssigneeWeekFilters] = useState<string[]>([]);
  const [compareWeekA, setCompareWeekA] = useState("");
  const [compareWeekB, setCompareWeekB] = useState("");

  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 1000);
    return () => window.clearInterval(timer);
  }, []);

  const handleFileUpload = async (file: File | null) => {
    if (!file) return;

    try {
      setError("");
      setCopied("");
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const parsed = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "" });

      const rows: RawRow[] = parsed.map((row) => {
        const obj: RawRow = {};
        Object.entries(row).forEach(([k, v]) => {
          obj[String(k).trim()] = typeof v === "number" ? v : String(v ?? "").trim();
        });
        return obj;
      });

      if (!rows.length) {
        setError("File không có dữ liệu.");
        setRawRows([]);
        setTasks([]);
        return;
      }

      const columns = Array.from(new Set(rows.flatMap((r) => Object.keys(r)).filter(Boolean)));

      const colTaskId = detectColumn(columns, COLUMN_ALIASES.taskId);
      const colIssueType = detectColumn(columns, COLUMN_ALIASES.issueType);
      const colAssignee = detectColumn(columns, COLUMN_ALIASES.assignee);
      const colStoryPoint = detectColumn(columns, COLUMN_ALIASES.storyPoint);
      const colName = detectColumn(columns, COLUMN_ALIASES.name);
      const colLabels = detectColumn(columns, COLUMN_ALIASES.labels);
      const colEpicLink = detectColumn(columns, COLUMN_ALIASES.epicLink);
      const colModule = detectColumn(columns, COLUMN_ALIASES.module);
      const colStatus = detectColumn(columns, COLUMN_ALIASES.status);

      const missing: string[] = [];
      if (!colTaskId) missing.push("Key");
      if (!colIssueType) missing.push("Issue Type");
      if (!colAssignee) missing.push("Assignee");
      if (!colStoryPoint) missing.push("Story Points");
      if (!colName) missing.push("Summary");
      if (!colLabels) missing.push("Labels");
      if (!colEpicLink) missing.push("Epic Link");
      if (!colModule) missing.push("ModuleName");
      if (!colStatus) missing.push("Status");

      if (missing.length) {
        setError(`Thiếu cột bắt buộc: ${missing.join(", ")}`);
        setRawRows(rows);
        setTasks([]);
        return;
      }

      const parsedTasks: Task[] = rows.map((row) => {
        const keyInfo = parseTaskKeyCell(row[colTaskId]);
        return {
          taskId: keyInfo.taskId,
          taskUrl: keyInfo.taskUrl,
          issueType: String(row[colIssueType] || "-"),
          assignee: String(row[colAssignee] || "Unknown"),
          storyPoint: storyPointValue(row[colStoryPoint]),
          name: String(row[colName] || ""),
          weeks: parseWeekLabels(row[colLabels]),
          epicLink: String(row[colEpicLink] || "-"),
          module: String(row[colModule] || "Unknown"),
          status: normalizeStatus(row[colStatus])
        };
      });

      setRawRows(rows);
      setTasks(parsedTasks);
      setProjectWeekFilter("all");
      setProjectModuleFilter("all");
      setProjectAssigneeFilter("all");
      setAssigneeAllWeeks(true);
      setAssigneeWeekFilters([]);
      const weekSet = new Set<string>();
      parsedTasks.forEach((task) => task.weeks.forEach((week) => weekSet.add(week)));
      const sortedWeeks = Array.from(weekSet).sort((a, b) => weekIndex(a) - weekIndex(b));
      setCompareWeekA(sortedWeeks.length >= 2 ? sortedWeeks[sortedWeeks.length - 2] : sortedWeeks[0] ?? "");
      setCompareWeekB(sortedWeeks.length >= 1 ? sortedWeeks[sortedWeeks.length - 1] : "");
    } catch {
      setError("Không thể đọc file. Vui lòng kiểm tra lại định dạng Excel/CSV.");
      setRawRows([]);
      setTasks([]);
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
      const matchWeek = projectWeekFilter === "all" || task.weeks.includes(projectWeekFilter);
      if (matchWeek) set.add(task.module);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [tasks, projectWeekFilter]);

  const projectAssignees = useMemo(() => {
    const set = new Set<string>();
    tasks.forEach((task) => {
      const matchWeek = projectWeekFilter === "all" || task.weeks.includes(projectWeekFilter);
      const matchModule = projectModuleFilter === "all" || task.module === projectModuleFilter;
      if (matchWeek && matchModule) set.add(task.assignee);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [tasks, projectWeekFilter, projectModuleFilter]);

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

    const prevWeek = projectWeekFilter !== "all" ? previousWeek(projectWeekFilter) : "";
    const filtered = tasks
      .filter((task) => {
        const matchWeek = projectWeekFilter === "all" || task.weeks.includes(projectWeekFilter);
        const matchModule = projectModuleFilter === "all" || task.module === projectModuleFilter;
        const matchAssignee = projectAssigneeFilter === "all" || task.assignee === projectAssigneeFilter;
        return matchWeek && matchModule && matchAssignee;
      })
      .map((task) => ({
        task,
        warning:
          Boolean(prevWeek) &&
          (taskStatusByWeek.get(task.taskId || `${task.module}||${task.name}`)?.get(prevWeek) ?? "-") === "done"
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
        prevWeek,
        showModule,
        showAssignee,
        moduleRowSpan: moduleCount.get(task.module) ?? 1,
        assigneeRowSpan: assigneeCount.get(assigneeKey) ?? 1
      };
    });
  }, [tasks, projectWeekFilter, projectModuleFilter, projectAssigneeFilter]);

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
    if (assigneeAllWeeks) return "Tất cả tuần";
    if (!assigneeWeekFilters.length) return "Chọn tuần";
    const ordered = allWeeks.filter((week) => assigneeWeekFilters.includes(week));
    return ordered.join(", ");
  }, [assigneeAllWeeks, assigneeWeekFilters, allWeeks]);

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
        const transition =
          statusA === "-" && statusB === "-"
            ? "Không có dữ liệu ở 2 tuần"
            : statusA === "-"
              ? `Mới xuất hiện ở ${compareWeekB}`
              : statusB === "-"
                ? `Không còn ở ${compareWeekB}`
                : statusA === statusB
                  ? "Không đổi"
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
          invalidDoneBoth
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
      { name: "Đổi trạng thái", value: compareSummary.changed, color: "#0ea5e9" },
      { name: "Không đổi", value: Math.max(compareSummary.total - compareSummary.changed, 0), color: "#64748b" },
      { name: "Done ở cả 2 tuần", value: compareSummary.invalidDoneBoth, color: "#dc2626" }
    ],
    [compareSummary]
  );

  const copyCsv = async (kind: "project" | "assignee" | "compare") => {
    try {
      let csv = "";
      if (kind === "project") {
        csv = toCsv(
          ["Module Name", "Assignee", "Task ID", "Issue Type", "Summary", "Epic Link", "Status", "Story Points", "Labels", "Prev Week Done But Still Appears"],
          projectRows.map((r) => [
            r.module,
            r.assignee,
            r.taskId,
            r.issueType,
            r.taskName,
            r.epicLink,
            r.status,
            r.storyPoint,
            r.weeks,
            r.prevDoneStillAppear ? "YES" : "NO"
          ])
        );
      }

      if (kind === "assignee") {
        csv = toCsv(
          ["Assignee", "C1", "C2", "C3", "C4", "C5", "Total"],
          assigneeRows.map((r) => [r.assignee, r.point1, r.point2, r.point3, r.point4, r.point5, r.total])
        );
      }

      if (kind === "compare") {
        csv = toCsv(
          ["Task ID", "Task Name", "Module Name", "Assignee", `Status ${compareWeekA}`, `Status ${compareWeekB}`, "Transition", "Invalid Done Both Weeks"],
          compareRows.map((r) => [r.taskId, r.taskName, r.module, r.assignee, r.statusA, r.statusB, r.transition, r.invalidDoneBoth ? "YES" : "NO"])
        );
      }

      await navigator.clipboard.writeText(csv);
      setCopied(
        kind === "project"
          ? "Đã copy CSV: thống kê theo dự án"
          : kind === "assignee"
            ? "Đã copy CSV: thống kê theo người dùng"
            : "Đã copy CSV: so sánh 2 tuần"
      );
    } catch {
      setCopied("Copy CSV thất bại. Hãy kiểm tra quyền clipboard của trình duyệt.");
    }
  };

  const totalTasks = tasks.length;
  const withWeek = tasks.filter((task) => task.weeks.length > 0).length;
  const currentWeekInfo = isoWeekInfo(now);
  const currentWeekCode = `W${twoDigits(currentWeekInfo.week)}`;
  const currentRange = weekRangeMonToFri(now);

  return (
    <main className="mx-auto max-w-7xl p-4 md:p-8">
      <section className="mb-6 grid gap-4 lg:grid-cols-4">
        <Card className="lg:col-span-2 border-2 border-primary/40 bg-primary/5">
          <CardHeader>
            <CardDescription>Tuần làm việc hiện tại</CardDescription>
            <CardTitle className="text-4xl text-primary">{currentWeekCode}</CardTitle>
          </CardHeader>
          <CardContent className="space-y-1 text-sm">
            <div>
              Thời gian làm việc: {formatDate(currentRange.monday)} - {formatDate(currentRange.friday)} (Thứ 2 - Thứ 6)
            </div>
            <div>
              Đồng hồ hiện tại: <span className="font-semibold">{formatClock(now)}</span> | {formatDate(now)}
            </div>
          </CardContent>
        </Card>
        <Card>
          <CardHeader>
            <CardDescription>Total Task</CardDescription>
            <CardTitle>{totalTasks}</CardTitle>
          </CardHeader>
        </Card>
        <Card>
          <CardHeader>
            <CardDescription>Có tuần hợp lệ</CardDescription>
            <CardTitle>{withWeek}</CardTitle>
          </CardHeader>
        </Card>
      </section>

      <section className="mb-6 grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
        <Card><CardHeader><CardDescription>Open</CardDescription><CardTitle className="text-sky-600">{overviewStatus.open}</CardTitle></CardHeader></Card>
        <Card><CardHeader><CardDescription>In Progress</CardDescription><CardTitle className="text-amber-500">{overviewStatus.inprogress}</CardTitle></CardHeader></Card>
        <Card><CardHeader><CardDescription>Done</CardDescription><CardTitle className="text-emerald-600">{overviewStatus.done}</CardTitle></CardHeader></Card>
        <Card><CardHeader><CardDescription>Other</CardDescription><CardTitle className="text-slate-500">{overviewStatus.other}</CardTitle></CardHeader></Card>
      </section>

      <section className="mb-6 rounded-2xl border bg-card/80 p-6 backdrop-blur-sm">
        <div className="mb-2 flex items-center gap-2">
          <Upload className="h-5 w-5 text-primary" />
          <h1 className="text-2xl font-semibold tracking-tight">Task Statistics Dashboard</h1>
        </div>
        <p className="mb-4 text-sm text-muted-foreground">
          Bộ lọc đơn giản theo tuần để thống kê theo dự án, người dùng, và độ khó task.
        </p>

        <div className="grid gap-4 md:grid-cols-1">
          <div className="space-y-2">
            <Label htmlFor="file">File Excel/CSV</Label>
            <Input id="file" type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleFileUpload(e.target.files?.[0] ?? null)} />
          </div>
        </div>

        {error && <p className="mt-3 text-sm text-red-600">{error}</p>}
        {copied && <p className="mt-3 text-sm text-emerald-700">{copied}</p>}
      </section>

      <section className="mb-6">
        <Card>
          <CardHeader>
            <CardTitle>1. Lọc theo dự án theo từng tuần</CardTitle>
            <CardDescription>Nhóm theo dự án - người dùng - từng task (rowspan/colspan)</CardDescription>
          </CardHeader>
          <CardContent>
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div className="grid w-full max-w-4xl grid-cols-1 gap-3 sm:grid-cols-3">
                <div className="space-y-2">
                  <Label>Tuần</Label>
                  <Select value={projectWeekFilter} onValueChange={setProjectWeekFilter}>
                    <SelectTrigger>
                      <SelectValue placeholder="Chọn tuần" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Tất cả tuần</SelectItem>
                      {allWeeks.map((week) => (
                        <SelectItem key={week} value={week}>
                          {week}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label>Dự án (optional)</Label>
                  <Select
                    value={projectModuleFilter}
                    onValueChange={(value) => {
                      setProjectModuleFilter(value);
                      setProjectAssigneeFilter("all");
                    }}
                  >
                    <SelectTrigger>
                      <SelectValue placeholder="Tất cả dự án" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Tất cả dự án</SelectItem>
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
                      <SelectValue placeholder="Tất cả assignee" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="all">Tất cả assignee</SelectItem>
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
              <div className="h-56 rounded-md border p-2">
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
              <div className="h-56 rounded-md border p-2">
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

            <div className="max-h-[560px] overflow-y-auto rounded-md border">
              <table className="w-full border-collapse text-sm">
                <thead>
                  <tr>
                    <th rowSpan={2} className="sticky top-0 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Module Name</th>
                    <th rowSpan={2} className="sticky top-0 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Assignee</th>
                    <th colSpan={6} className="sticky top-0 z-30 bg-muted/80 px-2 py-2 text-center font-medium text-muted-foreground">Task Details</th>
                  </tr>
                  <tr>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Task ID</th>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Task Name</th>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Status</th>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Story Point</th>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Weeks</th>
                    <th className="sticky top-9 z-30 bg-muted/80 px-2 py-2 text-left font-medium text-muted-foreground">Cảnh báo</th>
                  </tr>
                </thead>
                <tbody>
                {projectRows.map((row) => (
                  <tr key={`${row.module}-${row.assignee}-${row.taskKey}`} className="border-b transition-colors hover:bg-muted/50">
                    {row.showModule && (
                      <td rowSpan={row.moduleRowSpan} className="p-2 align-top font-medium">
                        {row.module}
                      </td>
                    )}
                    {row.showAssignee && (
                      <td rowSpan={row.assigneeRowSpan} className="p-2 align-top">
                        {row.assignee}
                      </td>
                    )}
                    <td className="p-2">
                      {row.taskUrl ? (
                        <a
                          href={row.taskUrl}
                          target="_blank"
                          rel="noreferrer"
                          className="text-primary underline underline-offset-2"
                        >
                          {row.taskId}
                        </a>
                      ) : (
                        row.taskId
                      )}
                    </td>
                    <td className="p-2">{row.taskName}</td>
                    <td className="p-2">
                      <StatusBadge status={row.status} />
                    </td>
                    <td className="p-2">{row.storyPoint}</td>
                    <td className="p-2">{row.weeks}</td>
                    <td className="p-2">
                      {row.prevDoneStillAppear ? (
                        <span
                          title={`Task đã done ở ${row.prevWeek} nhưng vẫn xuất hiện ở ${projectWeekFilter}`}
                          className="inline-flex items-center gap-1 text-amber-600"
                        >
                          <AlertTriangle className="h-4 w-4" />
                          done tuần trước
                        </span>
                      ) : (
                        "-"
                      )}
                    </td>
                  </tr>
                ))}
                </tbody>
              </table>
            </div>
          </CardContent>
        </Card>
      </section>

      <section className="mb-6">
        <Card>
          <CardHeader>
            <CardTitle>2. Người dùng làm bao nhiêu task theo độ khó</CardTitle>
            <CardDescription>Chọn nhiều tuần hoặc tất cả, mỗi người có bao nhiêu task ở từng độ khó</CardDescription>
          </CardHeader>
          <CardContent>
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div className="w-full max-w-xs space-y-2">
                <Label>Tuần (multi-select)</Label>
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
                      Tất cả tuần
                    </label>
                    <div className="max-h-44 space-y-1 overflow-auto border-t pt-2">
                      {allWeeks.map((week) => (
                        <label key={week} className="flex items-center gap-2 text-sm">
                          <input
                            type="checkbox"
                            checked={assigneeWeekFilters.includes(week)}
                            onChange={(e) => {
                              setAssigneeAllWeeks(false);
                              setAssigneeWeekFilters((prev) =>
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
              <Button variant="outline" onClick={() => copyCsv("assignee")}>Copy CSV</Button>
            </div>

            <div className="grid gap-4 lg:grid-cols-[260px,1fr]">
              <div className="h-64 rounded-md border p-2">
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

              <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Assignee</TableHead>
                  <TableHead>C1</TableHead>
                  <TableHead>C2</TableHead>
                  <TableHead>C3</TableHead>
                  <TableHead>C4</TableHead>
                  <TableHead>C5</TableHead>
                  <TableHead>Total</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {assigneeRows.map((row) => (
                  <TableRow key={row.assignee}>
                    <TableCell>{row.assignee}</TableCell>
                    <TableCell>{displayCount(row.point1)}</TableCell>
                    <TableCell>{displayCount(row.point2)}</TableCell>
                    <TableCell>{displayCount(row.point3)}</TableCell>
                    <TableCell>{displayCount(row.point4)}</TableCell>
                    <TableCell>{displayCount(row.point5)}</TableCell>
                    <TableCell>{displayCount(row.total)}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
              </Table>
            </div>
          </CardContent>
        </Card>
      </section>

      <section>
        <Card>
          <CardHeader>
            <CardTitle>3. So sánh chuyển trạng thái giữa 2 tuần</CardTitle>
            <CardDescription>
              So sánh theo Task ID giữa 2 tuần gần nhau và cảnh báo task bị done ở cả 2 tuần
            </CardDescription>
          </CardHeader>
          <CardContent>
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div className="grid w-full max-w-xl grid-cols-1 gap-3 sm:grid-cols-2">
                <div className="space-y-2">
                  <Label>Tuần A</Label>
                  <Select value={compareWeekA || "__none__"} onValueChange={(v) => setCompareWeekA(v === "__none__" ? "" : v)}>
                    <SelectTrigger>
                      <SelectValue placeholder="Chọn tuần A" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Chưa chọn</SelectItem>
                      {allWeeks.map((week) => (
                        <SelectItem key={`a-${week}`} value={week}>
                          {week}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                <div className="space-y-2">
                  <Label>Tuần B</Label>
                  <Select value={compareWeekB || "__none__"} onValueChange={(v) => setCompareWeekB(v === "__none__" ? "" : v)}>
                    <SelectTrigger>
                      <SelectValue placeholder="Chọn tuần B" />
                    </SelectTrigger>
                    <SelectContent>
                      <SelectItem value="__none__">Chưa chọn</SelectItem>
                      {allWeeks.map((week) => (
                        <SelectItem key={`b-${week}`} value={week}>
                          {week}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
              </div>
              <Button variant="outline" onClick={() => copyCsv("compare")}>Copy CSV</Button>
            </div>

            <div className="mb-3 text-sm text-muted-foreground">
              Tổng task so sánh: {compareSummary.total} | Task có thay đổi: {compareSummary.changed} | Cảnh báo done-done: {compareSummary.invalidDoneBoth}
            </div>

            <div className="mb-5 grid gap-4 lg:grid-cols-[260px,1fr]">
              <div className="h-64 rounded-md border p-2">
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
              <div className="grid content-start gap-2 rounded-md border p-4 text-sm">
                <div className="font-medium">Tổng quan so sánh</div>
                <div>Đổi trạng thái: <span className="font-semibold">{compareSummary.changed}</span></div>
                <div>Không đổi: <span className="font-semibold">{Math.max(compareSummary.total - compareSummary.changed, 0)}</span></div>
                <div>Done ở cả 2 tuần: <span className="font-semibold text-red-600">{compareSummary.invalidDoneBoth}</span></div>
              </div>
            </div>

            {compareWeekA && compareWeekB && compareWeekA === compareWeekB && (
              <p className="mb-3 text-sm text-amber-700">Hãy chọn 2 tuần khác nhau để so sánh đúng.</p>
            )}

            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Module</TableHead>
                  <TableHead>Assignee</TableHead>
                  <TableHead>Task ID</TableHead>
                  <TableHead>Task Name</TableHead>
                  <TableHead>Status {compareWeekA || "A"}</TableHead>
                  <TableHead>Status {compareWeekB || "B"}</TableHead>
                  <TableHead>Chuyển biến</TableHead>
                  <TableHead>Cảnh báo</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {compareGroupedRows.map((row) => (
                  <TableRow key={row.taskKey}>
                    {row.showModule && (
                      <TableCell rowSpan={row.moduleRowSpan} className="align-top font-medium">
                        {row.module}
                      </TableCell>
                    )}
                    {row.showAssignee && (
                      <TableCell rowSpan={row.assigneeRowSpan} className="align-top">
                        {row.assignee}
                      </TableCell>
                    )}
                    <TableCell>{row.taskId}</TableCell>
                    <TableCell>{row.taskName}</TableCell>
                    <TableCell>
                      <StatusBadge status={row.statusA} />
                    </TableCell>
                    <TableCell>
                      <StatusBadge status={row.statusB} />
                    </TableCell>
                    <TableCell>{row.transition}</TableCell>
                    <TableCell>{row.invalidDoneBoth ? "Task done ở cả 2 tuần (không hợp lệ)" : "-"}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>

            {!compareRows.length && (
              <div className="py-5 text-center text-sm text-muted-foreground">
                Chưa có dữ liệu so sánh cho 2 tuần đã chọn.
              </div>
            )}
          </CardContent>
        </Card>
      </section>

      {!rawRows.length && !error && (
        <div className="py-8 text-center text-sm text-muted-foreground">Upload file để bắt đầu thống kê.</div>
      )}
    </main>
  );
}
