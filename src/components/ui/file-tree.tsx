import { useEffect, useRef, useState } from "react";
import { useFileWatcher } from "@/hooks/useFileWatcher";
import { ChevronRight, File, FileSpreadsheet, Folder, FolderOpen, FolderPlus } from "lucide-react";
import { fileTree, type FileEntry } from "@/services/api";
import { useFileTreeStore } from "@/stores/fileTreeStore";
import { useProjectStore } from "@/stores/projectStore";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import {
  ContextMenu,
  ContextMenuContent,
  ContextMenuItem,
  ContextMenuTrigger,
} from "@/components/ui/context-menu";
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
} from "@/components/ui/alert-dialog";

// ── helpers ──────────────────────────────────────────────────────────────────

function formatSize(bytes: number | null): string {
  if (bytes === null) return "";
  if (bytes < 1024) return `${bytes}B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)}KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)}MB`;
}

function FileIcon({ entry }: { entry: FileEntry; isOpen?: boolean }) {
  if (entry.type === "dir") return null; // handled by caller
  const ext = entry.name.split(".").pop()?.toLowerCase();
  if (ext === "xlsx" || ext === "xls" || ext === "csv") {
    return <FileSpreadsheet className="size-3.5 shrink-0 text-ok" />;
  }
  return <File className="size-3.5 shrink-0 text-muted-foreground" />;
}

// ── FileTreeNode ─────────────────────────────────────────────────────────────

interface NodeProps {
  entry: FileEntry;
  depth: number;
}

function FileTreeNode({ entry, depth }: NodeProps) {
  const {
    expanded,
    loading,
    renamingPath,
    children,
    toggleExpanded,
    setChildren,
    setLoading,
    setRenamingPath,
    setDeletingPath,
    applyRename,
  } = useFileTreeStore();

  const isDir = entry.type === "dir";
  const isExpanded = expanded.has(entry.path);
  const isLoading = loading.has(entry.path);
  const isRenaming = renamingPath === entry.path;

  const inputRef = useRef<HTMLInputElement>(null);
  const committingRef = useRef(false);
  const [renameValue, setRenameValue] = useState(entry.name);

  // Focus input when rename starts
  useEffect(() => {
    if (isRenaming) {
      setRenameValue(entry.name);
      committingRef.current = false;
      setTimeout(() => inputRef.current?.select(), 0);
    }
  }, [isRenaming, entry.name]);

  async function handleExpand() {
    if (!isDir) return;
    toggleExpanded(entry.path);
    if (!expanded.has(entry.path) && children[entry.path] === undefined) {
      setLoading(entry.path, true);
      try {
        const data = await fileTree.list(entry.path);
        setChildren(entry.path, data.entries);
      } catch {
        setChildren(entry.path, []);
      } finally {
        setLoading(entry.path, false);
      }
    }
  }

  async function commitRename() {
    if (committingRef.current) return;
    const trimmed = renameValue.trim();
    if (!trimmed || trimmed === entry.name) {
      setRenamingPath(null);
      return;
    }
    committingRef.current = true;
    try {
      const result = await fileTree.rename(entry.path, trimmed);
      applyRename(entry.path, result.path, trimmed);
    } catch {
      // revert silently
    } finally {
      setRenamingPath(null);
    }
  }

  function handleRenameKeyDown(e: React.KeyboardEvent<HTMLInputElement>) {
    if (e.key === "Enter") { e.preventDefault(); commitRename(); }
    if (e.key === "Escape") { setRenamingPath(null); }
  }

  function handleDragStart(e: React.DragEvent) {
    e.dataTransfer.setData("text/plain", entry.path);
    e.dataTransfer.effectAllowed = "move";
  }

  function handleDragOver(e: React.DragEvent) {
    if (!isDir) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "move";
  }

  async function handleDrop(e: React.DragEvent) {
    if (!isDir) return;
    e.preventDefault();
    const srcPath = e.dataTransfer.getData("text/plain");
    if (!srcPath || srcPath === entry.path) return;
    try {
      const result = await fileTree.move(srcPath, entry.path);
      useFileTreeStore.getState().applyMove(srcPath, entry.path, result.path);
    } catch {
      // ignore
    }
  }

  const paddingLeft = depth * 12 + 8;

  return (
    <>
      <ContextMenu>
        <ContextMenuTrigger asChild>
          <div
            className={cn(
              "flex items-center gap-1.5 py-1 pr-3 text-xs-plus cursor-pointer transition-colors hover:bg-surface-hover select-none",
            )}
            style={{ paddingLeft }}
            draggable
            onDragStart={handleDragStart}
            onDragOver={handleDragOver}
            onDrop={handleDrop}
            onClick={handleExpand}
          >
            {isDir ? (
              <>
                <ChevronRight
                  className={cn(
                    "size-3 shrink-0 text-muted-foreground transition-transform duration-150",
                    isExpanded && "rotate-90",
                  )}
                />
                {isExpanded
                  ? <FolderOpen className="size-3.5 shrink-0 text-warning" />
                  : <Folder className="size-3.5 shrink-0 text-warning" />}
              </>
            ) : (
              <>
                <span className="size-3 shrink-0" />
                <FileIcon entry={entry} />
              </>
            )}

            {isRenaming ? (
              <input
                ref={inputRef}
                className="flex-1 min-w-0 bg-background border border-border rounded px-1 text-xs-plus outline-none"
                value={renameValue}
                onChange={(e) => setRenameValue(e.target.value)}
                onKeyDown={handleRenameKeyDown}
                onBlur={commitRename}
                onClick={(e) => e.stopPropagation()}
              />
            ) : (
              <span className="flex-1 truncate text-muted-foreground">{entry.name}</span>
            )}

            {!isRenaming && entry.size !== null && (
              <span className="text-2xs text-muted-foreground shrink-0">{formatSize(entry.size)}</span>
            )}
            {isLoading && <span className="text-2xs text-muted-foreground shrink-0">…</span>}
          </div>
        </ContextMenuTrigger>
        <ContextMenuContent>
          <ContextMenuItem onSelect={() => setRenamingPath(entry.path)}>重命名</ContextMenuItem>
          <ContextMenuItem
            className="text-destructive focus:text-destructive"
            onSelect={() => setDeletingPath(entry.path)}
          >
            删除
          </ContextMenuItem>
        </ContextMenuContent>
      </ContextMenu>

      {isDir && isExpanded && (
        <FileTreeList dirPath={entry.path} depth={depth + 1} />
      )}
    </>
  );
}

// ── FileTreeList ─────────────────────────────────────────────────────────────

function FileTreeList({ dirPath, depth }: { dirPath: string; depth: number }) {
  const entries = useFileTreeStore((s) => s.children[dirPath]);

  if (entries === undefined || entries === null) {
    return (
      <div className="py-1 text-2xs text-muted-foreground" style={{ paddingLeft: depth * 12 + 8 }}>
        加载中…
      </div>
    );
  }

  if (entries.length === 0) {
    return (
      <div className="py-1 text-2xs text-muted-foreground" style={{ paddingLeft: depth * 12 + 8 }}>
        空文件夹
      </div>
    );
  }

  return (
    <>
      {entries.map((entry) => (
        <FileTreeNode key={entry.path} entry={entry} depth={depth} />
      ))}
    </>
  );
}

// ── FileTree (root) ──────────────────────────────────────────────────────────

export function FileTree() {
  const { rootPath, setRootPath, setChildren, deletingPath, setDeletingPath, applyDelete } =
    useFileTreeStore();
  const projectPath = useProjectStore((s) => s.currentProject?.path);

  async function loadRoot(path?: string) {
    try {
      const data = await fileTree.list(path);
      setRootPath(data.path);
      setChildren(data.path, data.entries);
    } catch {
      // ignore
    }
  }

  useEffect(() => {
    loadRoot(projectPath);
  }, [projectPath]); // eslint-disable-line react-hooks/exhaustive-deps

  useFileWatcher({
    path: rootPath ?? "",
    enabled: !!rootPath,
    onChanges: (affectedDirs) => {
      const store = useFileTreeStore.getState();
      for (const dir of affectedDirs) {
        if (store.children[dir] !== undefined) {
          store.setChildren(dir, []);
          fileTree.list(dir)
            .then((data) => store.setChildren(dir, data.entries))
            .catch(() => {});
        }
      }
    },
  });

  async function handleOpenFolder() {
    const path = await fileTree.openFolderDialog();
    if (path) {
      setRootPath(path);
      loadRoot(path);
    }
  }

  async function confirmDelete() {
    if (!deletingPath) return;
    try {
      await fileTree.delete(deletingPath);
      applyDelete(deletingPath);
    } catch {
      // ignore
    } finally {
      setDeletingPath(null);
    }
  }

  return (
    <>
      <div className="flex items-center justify-between px-3 py-1">
        <span className="text-2xs text-muted-foreground truncate flex-1" title={rootPath}>
          {rootPath ? rootPath.split(/[/\\]/).pop() : "主目录"}
        </span>
        <Button
          variant="ghost"
          size="icon"
          className="size-5 shrink-0"
          title="打开文件夹"
          onClick={handleOpenFolder}
        >
          <FolderPlus className="size-3" />
        </Button>
      </div>

      {rootPath && <FileTreeList dirPath={rootPath} depth={0} />}

      <AlertDialog open={!!deletingPath} onOpenChange={(open) => !open && setDeletingPath(null)}>
        <AlertDialogContent>
          <AlertDialogHeader>
            <AlertDialogTitle>确认删除</AlertDialogTitle>
            <AlertDialogDescription>
              确定要删除 <strong>{deletingPath?.split(/[/\\]/).pop()}</strong> 吗？此操作不可撤销。
            </AlertDialogDescription>
          </AlertDialogHeader>
          <AlertDialogFooter>
            <AlertDialogCancel>取消</AlertDialogCancel>
            <AlertDialogAction onClick={confirmDelete}>删除</AlertDialogAction>
          </AlertDialogFooter>
        </AlertDialogContent>
      </AlertDialog>
    </>
  );
}
