import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { Upload } from "lucide-react";
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
  assignee: string;
  storyPoint: number;
  name: string;
  weeks: string[];
  module: string;
  status: Status;
};

const COLUMN_ALIASES = {
  taskId: ["taskid", "task id", "id", "ticketid", "jiraid"],
  assignee: ["assignee"],
  storyPoint: ["storypoint", "story point"],
  name: ["name", "taskname", "task name"],
  labels: ["labels", "label"],
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
  if (v === "open") return "open";
  if (v === "inprogress" || v === "in-progress" || v === "doing") return "inprogress";
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

function statusIndex(status: Status) {
  if (status === "open") return 0;
  if (status === "inprogress") return 1;
  if (status === "done") return 2;
  return 3;
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

export default function App() {
  const [rawRows, setRawRows] = useState<RawRow[]>([]);
  const [tasks, setTasks] = useState<Task[]>([]);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState("");

  const [projectWeekFilter, setProjectWeekFilter] = useState("all");
  const [assigneeWeekFilter, setAssigneeWeekFilter] = useState("all");
  const [compareWeekA, setCompareWeekA] = useState("");
  const [compareWeekB, setCompareWeekB] = useState("");

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
      const colAssignee = detectColumn(columns, COLUMN_ALIASES.assignee);
      const colStoryPoint = detectColumn(columns, COLUMN_ALIASES.storyPoint);
      const colName = detectColumn(columns, COLUMN_ALIASES.name);
      const colLabels = detectColumn(columns, COLUMN_ALIASES.labels);
      const colModule = detectColumn(columns, COLUMN_ALIASES.module);
      const colStatus = detectColumn(columns, COLUMN_ALIASES.status);

      const missing: string[] = [];
      if (!colAssignee) missing.push("Assignee");
      if (!colStoryPoint) missing.push("Story point");
      if (!colName) missing.push("Name");
      if (!colLabels) missing.push("Labels");
      if (!colModule) missing.push("Module Name");
      if (!colStatus) missing.push("Status");

      if (missing.length) {
        setError(`Thiếu cột bắt buộc: ${missing.join(", ")}`);
        setRawRows(rows);
        setTasks([]);
        return;
      }

      const parsedTasks: Task[] = rows.map((row) => ({
        taskId: colTaskId ? String(row[colTaskId] || "") : "",
        assignee: String(row[colAssignee] || "Unknown"),
        storyPoint: storyPointValue(row[colStoryPoint]),
        name: String(row[colName] || ""),
        weeks: parseWeekLabels(row[colLabels]),
        module: String(row[colModule] || "Unknown"),
        status: normalizeStatus(row[colStatus])
      }));

      setRawRows(rows);
      setTasks(parsedTasks);
      setProjectWeekFilter("all");
      setAssigneeWeekFilter("all");
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

  const projectRows = useMemo(() => {
    const filtered = tasks
      .filter((task) => projectWeekFilter === "all" || task.weeks.includes(projectWeekFilter))
      .sort(
        (a, b) =>
          a.module.localeCompare(b.module) ||
          a.assignee.localeCompare(b.assignee) ||
          statusIndex(a.status) - statusIndex(b.status) ||
          a.name.localeCompare(b.name)
      );

    const moduleCount = new Map<string, number>();
    const assigneeCount = new Map<string, number>();
    filtered.forEach((task) => {
      moduleCount.set(task.module, (moduleCount.get(task.module) ?? 0) + 1);
      const key = `${task.module}||${task.assignee}`;
      assigneeCount.set(key, (assigneeCount.get(key) ?? 0) + 1);
    });

    const seenModule = new Set<string>();
    const seenAssignee = new Set<string>();

    return filtered.map((task) => {
      const assigneeKey = `${task.module}||${task.assignee}`;
      const showModule = !seenModule.has(task.module);
      const showAssignee = !seenAssignee.has(assigneeKey);
      if (showModule) seenModule.add(task.module);
      if (showAssignee) seenAssignee.add(assigneeKey);

      return {
        module: task.module,
        assignee: task.assignee,
        taskId: task.taskId || "-",
        taskName: task.name || "-",
        status: task.status,
        storyPoint: task.storyPoint,
        weeks: task.weeks.join(","),
        showModule,
        showAssignee,
        moduleRowSpan: moduleCount.get(task.module) ?? 1,
        assigneeRowSpan: assigneeCount.get(assigneeKey) ?? 1
      };
    });
  }, [tasks, projectWeekFilter]);

  const assigneeRows = useMemo(() => {
    const map = new Map<
      string,
      { assignee: string; point1: number; point2: number; point3: number; point4: number; point5: number; total: number }
    >();

    tasks.forEach((task) => {
      const matched = assigneeWeekFilter === "all" || task.weeks.includes(assigneeWeekFilter);
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
  }, [tasks, assigneeWeekFilter]);

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
          Number(b.invalidDoneBoth) - Number(a.invalidDoneBoth) ||
          a.module.localeCompare(b.module) ||
          a.assignee.localeCompare(b.assignee) ||
          a.taskId.localeCompare(b.taskId)
      );
  }, [tasks, compareWeekA, compareWeekB]);

  const compareSummary = useMemo(() => {
    const changed = compareRows.filter((row) => row.statusA !== row.statusB).length;
    const invalidDoneBoth = compareRows.filter((row) => row.invalidDoneBoth).length;
    return { changed, invalidDoneBoth, total: compareRows.length };
  }, [compareRows]);

  const copyCsv = async (kind: "project" | "assignee" | "compare") => {
    try {
      let csv = "";
      if (kind === "project") {
        csv = toCsv(
          ["Module Name", "Assignee", "Task ID", "Task Name", "Status", "Story Point", "Weeks"],
          projectRows.map((r) => [r.module, r.assignee, r.taskId, r.taskName, r.status, r.storyPoint, r.weeks])
        );
      }

      if (kind === "assignee") {
        csv = toCsv(
          ["Assignee", "SP1", "SP2", "SP3", "SP4", "SP5", "Total"],
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

  return (
    <main className="mx-auto max-w-7xl p-4 md:p-8">
      <section className="mb-6 rounded-2xl border bg-card/80 p-6 backdrop-blur-sm">
        <div className="mb-2 flex items-center gap-2">
          <Upload className="h-5 w-5 text-primary" />
          <h1 className="text-2xl font-semibold tracking-tight">Task Statistics Dashboard</h1>
        </div>
        <p className="mb-4 text-sm text-muted-foreground">
          Bộ lọc đơn giản theo tuần để thống kê theo dự án, người dùng, và độ khó task.
        </p>

        <div className="grid gap-4 md:grid-cols-2">
          <div className="space-y-2">
            <Label htmlFor="file">File Excel/CSV</Label>
            <Input id="file" type="file" accept=".xlsx,.xls,.csv" onChange={(e) => handleFileUpload(e.target.files?.[0] ?? null)} />
          </div>
          <div className="grid grid-cols-2 gap-3">
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
              <div className="w-full max-w-xs space-y-2">
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
              <Button variant="outline" onClick={() => copyCsv("project")}>Copy CSV</Button>
            </div>

            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead rowSpan={2}>Module Name</TableHead>
                  <TableHead rowSpan={2}>Assignee</TableHead>
                  <TableHead colSpan={5} className="text-center">
                    Task Details
                  </TableHead>
                </TableRow>
                <TableRow>
                  <TableHead>Task ID</TableHead>
                  <TableHead>Task Name</TableHead>
                  <TableHead>Status</TableHead>
                  <TableHead>Story Point</TableHead>
                  <TableHead>Weeks</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {projectRows.map((row) => (
                  <TableRow key={`${row.module}-${row.assignee}-${row.taskId}-${row.taskName}`}>
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
                    <TableCell>{row.status}</TableCell>
                    <TableCell>{row.storyPoint}</TableCell>
                    <TableCell>{row.weeks}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </CardContent>
        </Card>
      </section>

      <section className="mb-6">
        <Card>
          <CardHeader>
            <CardTitle>2. Người dùng làm bao nhiêu task theo độ khó</CardTitle>
            <CardDescription>Trong tuần đã chọn, mỗi người có bao nhiêu task ở từng Story Point</CardDescription>
          </CardHeader>
          <CardContent>
            <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-end sm:justify-between">
              <div className="w-full max-w-xs space-y-2">
                <Label>Tuần</Label>
                <Select value={assigneeWeekFilter} onValueChange={setAssigneeWeekFilter}>
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
              <Button variant="outline" onClick={() => copyCsv("assignee")}>Copy CSV</Button>
            </div>

            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Assignee</TableHead>
                  <TableHead>SP1</TableHead>
                  <TableHead>SP2</TableHead>
                  <TableHead>SP3</TableHead>
                  <TableHead>SP4</TableHead>
                  <TableHead>SP5</TableHead>
                  <TableHead>Total</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {assigneeRows.map((row) => (
                  <TableRow key={row.assignee}>
                    <TableCell>{row.assignee}</TableCell>
                    <TableCell>{row.point1}</TableCell>
                    <TableCell>{row.point2}</TableCell>
                    <TableCell>{row.point3}</TableCell>
                    <TableCell>{row.point4}</TableCell>
                    <TableCell>{row.point5}</TableCell>
                    <TableCell>{row.total}</TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
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

            {compareWeekA && compareWeekB && compareWeekA === compareWeekB && (
              <p className="mb-3 text-sm text-amber-700">Hãy chọn 2 tuần khác nhau để so sánh đúng.</p>
            )}

            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Task ID</TableHead>
                  <TableHead>Task Name</TableHead>
                  <TableHead>Module</TableHead>
                  <TableHead>Assignee</TableHead>
                  <TableHead>Status {compareWeekA || "A"}</TableHead>
                  <TableHead>Status {compareWeekB || "B"}</TableHead>
                  <TableHead>Chuyển biến</TableHead>
                  <TableHead>Cảnh báo</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {compareRows.map((row) => (
                  <TableRow key={row.taskKey}>
                    <TableCell>{row.taskId}</TableCell>
                    <TableCell>{row.taskName}</TableCell>
                    <TableCell>{row.module}</TableCell>
                    <TableCell>{row.assignee}</TableCell>
                    <TableCell>{row.statusA}</TableCell>
                    <TableCell>{row.statusB}</TableCell>
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
