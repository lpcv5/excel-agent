import { create } from "zustand";
import type { FileEntry } from "@/services/api";

interface FileTreeState {
  rootPath: string;
  children: Record<string, FileEntry[] | null>;
  expanded: Set<string>;
  loading: Set<string>;
  renamingPath: string | null;
  deletingPath: string | null;

  setRootPath: (path: string) => void;
  setChildren: (dirPath: string, entries: FileEntry[]) => void;
  setLoading: (path: string, on: boolean) => void;
  toggleExpanded: (path: string) => void;
  setRenamingPath: (path: string | null) => void;
  setDeletingPath: (path: string | null) => void;
  applyRename: (oldPath: string, newPath: string, newName: string) => void;
  applyDelete: (path: string) => void;
  applyMove: (srcPath: string, dstDirPath: string, newPath: string) => void;
}

export const useFileTreeStore = create<FileTreeState>((set) => ({
  rootPath: "",
  children: {},
  expanded: new Set(),
  loading: new Set(),
  renamingPath: null,
  deletingPath: null,

  setRootPath: (path) =>
    set({ rootPath: path, children: {}, expanded: new Set(), loading: new Set() }),

  setChildren: (dirPath, entries) =>
    set((s) => ({ children: { ...s.children, [dirPath]: entries } })),

  setLoading: (path, on) =>
    set((s) => {
      const next = new Set(s.loading);
      if (on) next.add(path);
      else next.delete(path);
      return { loading: next };
    }),

  toggleExpanded: (path) =>
    set((s) => {
      const next = new Set(s.expanded);
      if (next.has(path)) next.delete(path);
      else next.add(path);
      return { expanded: next };
    }),

  setRenamingPath: (path) => set({ renamingPath: path }),
  setDeletingPath: (path) => set({ deletingPath: path }),

  applyRename: (oldPath, newPath, newName) =>
    set((s) => {
      // Update entry in parent's children list
      const parentPath = oldPath.substring(0, Math.max(oldPath.lastIndexOf("/"), oldPath.lastIndexOf("\\")));
      const newChildren = { ...s.children };

      if (newChildren[parentPath]) {
        newChildren[parentPath] = newChildren[parentPath]!.map((e) =>
          e.path === oldPath ? { ...e, path: newPath, name: newName } : e
        );
      }

      // Re-key any cached children under the old path
      for (const key of Object.keys(newChildren)) {
        if (key === oldPath || key.startsWith(oldPath + "/") || key.startsWith(oldPath + "\\")) {
          const newKey = newPath + key.slice(oldPath.length);
          newChildren[newKey] = newChildren[key];
          delete newChildren[key];
        }
      }

      // Update expanded set
      const newExpanded = new Set<string>();
      for (const p of s.expanded) {
        if (p === oldPath || p.startsWith(oldPath + "/") || p.startsWith(oldPath + "\\")) {
          newExpanded.add(newPath + p.slice(oldPath.length));
        } else {
          newExpanded.add(p);
        }
      }

      return { children: newChildren, expanded: newExpanded };
    }),

  applyDelete: (path) =>
    set((s) => {
      const parentPath = path.substring(0, Math.max(path.lastIndexOf("/"), path.lastIndexOf("\\")));
      const newChildren = { ...s.children };

      if (newChildren[parentPath]) {
        newChildren[parentPath] = newChildren[parentPath]!.filter((e) => e.path !== path);
      }

      // Remove cached children for deleted path
      for (const key of Object.keys(newChildren)) {
        if (key === path || key.startsWith(path + "/") || key.startsWith(path + "\\")) {
          delete newChildren[key];
        }
      }

      const newExpanded = new Set(s.expanded);
      newExpanded.delete(path);

      return { children: newChildren, expanded: newExpanded };
    }),

  applyMove: (srcPath, dstDirPath, newPath) =>
    set((s) => {
      const srcParent = srcPath.substring(0, Math.max(srcPath.lastIndexOf("/"), srcPath.lastIndexOf("\\")));
      const newChildren = { ...s.children };

      // Remove from old parent
      if (newChildren[srcParent]) {
        newChildren[srcParent] = newChildren[srcParent]!.filter((e) => e.path !== srcPath);
      }

      // Add to new parent if loaded
      if (newChildren[dstDirPath]) {
        const movedEntry = s.children[srcParent]?.find((e) => e.path === srcPath);
        if (movedEntry) {
          const name = srcPath.split(/[/\\]/).pop() ?? movedEntry.name;
          newChildren[dstDirPath] = [...newChildren[dstDirPath]!, { ...movedEntry, path: newPath, name }];
        }
      }

      // Clear cached children for moved dir
      for (const key of Object.keys(newChildren)) {
        if (key === srcPath || key.startsWith(srcPath + "/") || key.startsWith(srcPath + "\\")) {
          delete newChildren[key];
        }
      }

      return { children: newChildren };
    }),
}));
