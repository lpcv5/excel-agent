import { useEffect, useRef, useCallback } from "react";

function getToken(): string {
  return (
    new URLSearchParams(window.location.search).get("token") ??
    localStorage.getItem("app_token") ??
    ""
  );
}

export interface FileChangeEvent {
  type: "added" | "modified" | "deleted";
  path: string;
  dir: string;
}

interface UseFileWatcherOptions {
  path: string;
  onChanges: (affectedDirs: string[], changes: FileChangeEvent[]) => void;
  enabled?: boolean;
}

export function useFileWatcher({ path, onChanges, enabled = true }: UseFileWatcherOptions) {
  const wsRef = useRef<WebSocket | null>(null);
  const reconnectTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const mountedRef = useRef(true);
  const pathRef = useRef(path);
  const onChangesRef = useRef(onChanges);

  onChangesRef.current = onChanges;
  pathRef.current = path;

  const connect = useCallback(() => {
    if (!mountedRef.current || !enabled) return;
    const token = getToken();
    const wsUrl = `ws://localhost:8765/ws/files/watch?path=${encodeURIComponent(pathRef.current)}&token=${encodeURIComponent(token)}`;
    const ws = new WebSocket(wsUrl);
    wsRef.current = ws;

    ws.onmessage = (evt) => {
      try {
        const msg = JSON.parse(evt.data);
        if (msg.event === "changes" && msg.changes && msg.affected_dirs) {
          onChangesRef.current(msg.affected_dirs, msg.changes);
        }
      } catch { /* ignore */ }
    };

    ws.onclose = () => {
      if (!mountedRef.current) return;
      reconnectTimerRef.current = setTimeout(connect, 2000);
    };

    ws.onerror = () => ws.close();
  }, [enabled]);

  const switchPath = useCallback((newPath: string) => {
    const ws = wsRef.current;
    if (ws?.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify({ action: "watch", path: newPath }));
    }
  }, []);

  useEffect(() => {
    mountedRef.current = true;
    if (enabled && path) connect();
    return () => {
      mountedRef.current = false;
      if (reconnectTimerRef.current) clearTimeout(reconnectTimerRef.current);
      wsRef.current?.close();
      wsRef.current = null;
    };
  }, [enabled, connect]);

  // Path changes after mount: send switch message instead of reconnecting
  useEffect(() => {
    if (enabled && path) switchPath(path);
  }, [path, enabled, switchPath]);

  return { switchPath };
}
