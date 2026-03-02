import * as XLSX from "xlsx";

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

type RawRow = Record<string, string | number>;

type WorkerRequest = {
  type: "parse";
  requestId: number;
  buffer: ArrayBuffer;
};

type WorkerResponse =
  | { type: "success"; requestId: number; rowCount: number; tasks: Task[] }
  | { type: "error"; requestId: number; rowCount: number; error: string };

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

function parseExcelBuffer(buffer: ArrayBuffer): WorkerResponse {
  try {
    const workbook = XLSX.read(buffer, { type: "array" });
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
      return { type: "error", requestId: -1, rowCount: 0, error: "The file has no data." };
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
      return {
        type: "error",
        requestId: -1,
        rowCount: rows.length,
        error: `Missing required columns: ${missing.join(", ")}`
      };
    }

    const tasks: Task[] = rows.map((row) => {
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

    return { type: "success", requestId: -1, rowCount: rows.length, tasks };
  } catch {
    return { type: "error", requestId: -1, rowCount: 0, error: "Unable to read file. Please check Excel/CSV format." };
  }
}

self.onmessage = (event: MessageEvent<WorkerRequest>) => {
  if (event.data?.type !== "parse") return;
  const { requestId, buffer } = event.data;
  const parsed = parseExcelBuffer(buffer);
  const response: WorkerResponse =
    parsed.type === "success"
      ? { ...parsed, requestId }
      : { ...parsed, requestId };
  self.postMessage(response);
};

export {};
