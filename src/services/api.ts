import type { ToolCall, TaskItem } from "@/types";

export interface StreamCallbacks {
  onThinkingToken: (token: string) => void;
  onContentToken: (token: string) => void;
  onToolStart: (call: ToolCall) => void;
  onToolResult: (callId: string, patch: Pick<ToolCall, "status" | "result" | "duration_ms" | "args">) => void;
  onTasksUpdate: (tasks: TaskItem[]) => void;
  onDone: (error?: string) => void;
}

function getToken(): string {
  return (
    new URLSearchParams(window.location.search).get("token") ??
    localStorage.getItem("app_token") ??
    ""
  );
}

function authHeaders(): Record<string, string> {
  const token = getToken();
  return token
    ? { "X-App-Token": token, "Content-Type": "application/json" }
    : { "Content-Type": "application/json" };
}

export function streamMessage(
  message: string,
  callbacks: StreamCallbacks,
  threadId?: string,
): () => void {
  const controller = new AbortController();

  (async () => {
    let res: Response;
    try {
      res = await fetch("/api/stream", {
        method: "POST",
        headers: authHeaders(),
        body: JSON.stringify({ message, thread_id: threadId }),
        signal: controller.signal,
      });
    } catch (err) {
      if ((err as Error).name !== "AbortError") callbacks.onDone(String(err));
      return;
    }

    if (!res.ok) {
      callbacks.onDone(`HTTP ${res.status}`);
      return;
    }

    const reader = res.body!.getReader();
    const decoder = new TextDecoder();
    let buf = "";

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });

      const parts = buf.split("\n\n");
      buf = parts.pop() ?? "";

      for (const part of parts) {
        const line = part.trim();
        if (!line.startsWith("data:")) continue;
        const json = line.slice(5).trim();
        if (!json) continue;

        let msg: Record<string, unknown>;
        try {
          msg = JSON.parse(json);
        } catch {
          continue;
        }

        switch (msg.type) {
          case "stream:text":
            callbacks.onContentToken(msg.token as string);
            break;
          case "stream:thinking":
            callbacks.onThinkingToken(msg.token as string);
            break;
          case "tool:start":
            callbacks.onToolStart({
              id: msg.id as string,
              name: msg.name as string,
              status: "running",
              args: (msg.args ?? {}) as Record<string, unknown>,
            });
            break;
          case "tool:result":
            callbacks.onToolResult(msg.id as string, {
              status: (msg.status as "success" | "error") ?? "success",
              result: msg.result as string | undefined,
              duration_ms: msg.duration_ms as number | undefined,
              args: (msg.args as Record<string, unknown>) || undefined,
            });
            break;
          case "tasks:update":
            callbacks.onTasksUpdate(msg.tasks as TaskItem[]);
            break;
          case "stream:done":
            callbacks.onDone(msg.error as string | undefined);
            return;
        }
      }
    }
  })().catch((err) => {
    if ((err as Error).name !== "AbortError") callbacks.onDone(String(err));
  });

  return () => controller.abort();
}

export async function cancelStream(): Promise<void> {
  await fetch("/api/stream/cancel", { method: "POST", headers: authHeaders() });
}

export const fileDialog = {
  open: async (options?: { filters?: string[]; multiple?: boolean }) => {
    const res = await fetch("/api/dialog/open", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ multiple: options?.multiple ?? false, filters: options?.filters ?? [] }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.paths as string[] | null;
  },
  save: async (options?: { default_path?: string; filters?: string[] }) => {
    const res = await fetch("/api/dialog/save", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ default_path: options?.default_path ?? "", filters: options?.filters ?? [] }),
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.path as string | null;
  },
};

export interface FileEntry {
  name: string;
  path: string;
  type: "file" | "dir";
  size: number | null;
}

export const fileTree = {
  list: async (path?: string): Promise<{ path: string; entries: FileEntry[] }> => {
    const url = path ? `/api/files?path=${encodeURIComponent(path)}` : "/api/files";
    const res = await fetch(url, { headers: authHeaders() });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  },
  rename: async (path: string, newName: string): Promise<{ path: string }> => {
    const res = await fetch("/api/files/rename", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ path, new_name: newName }),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  },
  delete: async (path: string): Promise<void> => {
    const res = await fetch("/api/files", {
      method: "DELETE",
      headers: authHeaders(),
      body: JSON.stringify({ path }),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
  },
  move: async (src: string, dstDir: string): Promise<{ path: string }> => {
    const res = await fetch("/api/files/move", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ src, dst_dir: dstDir }),
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    return res.json();
  },
  openFolderDialog: async (): Promise<string | null> => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const api = (window as any).pywebview?.api;
    if (api?.open_folder) {
      const result = await api.open_folder();
      return result as string | null;
    }
    const res = await fetch("/api/dialog/folder", {
      method: "POST",
      headers: authHeaders(),
    });
    if (!res.ok) return null;
    const data = await res.json();
    return data.path as string | null;
  },
};

export const appInfo = {
  getVersion: async (): Promise<string> => {
    const res = await fetch("/api/version", { headers: authHeaders() });
    if (!res.ok) return "unknown";
    const data = await res.json();
    return data.version;
  },
};

export const projects = {
  getRecent: async (): Promise<{ recent_projects: import("@/types/project").RecentProject[] }> => {
    const res = await fetch("/api/projects/recent", { headers: authHeaders() });
    if (!res.ok) return { recent_projects: [] };
    return res.json();
  },
  create: async (body: {
    project_path: string;
    name: string;
    data_sources?: import("@/types/project").DataSource[];
    options?: import("@/types/project").ProjectOptions;
  }): Promise<{ project: import("@/types/project").Project }> => {
    const res = await fetch("/api/projects/create", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify(body),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { detail?: string }).detail ?? `HTTP ${res.status}`);
    }
    return res.json();
  },
  open: async (project_path: string): Promise<{ project: import("@/types/project").Project }> => {
    const res = await fetch("/api/projects/open", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ project_path }),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { detail?: string }).detail ?? `HTTP ${res.status}`);
    }
    return res.json();
  },
  getCurrent: async (): Promise<{ project: import("@/types/project").Project | null }> => {
    const res = await fetch("/api/projects/current", { headers: authHeaders() });
    if (!res.ok) return { project: null };
    return res.json();
  },
  close: async (): Promise<void> => {
    await fetch("/api/projects/close", { method: "POST", headers: authHeaders() });
  },
  addDataSource: async (source: { type: string; path: string; name: string }): Promise<{ project: import("@/types/project").Project }> => {
    const res = await fetch("/api/projects/data-sources", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify(source),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { detail?: string }).detail ?? `HTTP ${res.status}`);
    }
    return res.json();
  },
  removeDataSource: async (sourceId: string): Promise<{ project: import("@/types/project").Project }> => {
    const res = await fetch(`/api/projects/data-sources/${encodeURIComponent(sourceId)}`, {
      method: "DELETE",
      headers: authHeaders(),
    });
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      throw new Error((err as { detail?: string }).detail ?? `HTTP ${res.status}`);
    }
    return res.json();
  },
  triggerAnalysis: async (): Promise<void> => {
    await fetch("/api/projects/analyze", { method: "POST", headers: authHeaders() });
  },
  getAnalysisStatus: async (): Promise<{ status: import("@/types/project").AnalysisStatus; project_root: string | null }> => {
    const res = await fetch("/api/projects/analysis-status", { headers: authHeaders() });
    if (!res.ok) return { status: "idle", project_root: null };
    return res.json();
  },
};

export interface ModelEntry {
  id: string;
  provider: string;
  model_name: string;
  api_key: string;
  base_url: string | null;
  display_name: string;
}

export interface AppSettings {
  models: ModelEntry[];
  main_model_id: string | null;
  subagents_model_id: string | null;
  analysis_model_id: string | null;
  language: string;
}

export const settings = {
  get: async (): Promise<AppSettings> => {
    const res = await fetch("/api/settings", { headers: authHeaders() });
    return res.json();
  },
  save: async (s: AppSettings): Promise<void> => {
    await fetch("/api/settings", {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify(s),
    });
  },
  getProviders: async (): Promise<{ name: string; default_model: string }[]> => {
    const res = await fetch("/api/settings/providers", { headers: authHeaders() });
    return res.json();
  },
};
