import { create } from "zustand";
import type { Project, RecentProject, DataSource, AnalysisStatus } from "@/types/project";

interface ProjectState {
  currentProject: Project | null;
  recentProjects: RecentProject[];
  analysisStatus: AnalysisStatus;
  setCurrentProject: (project: Project | null) => void;
  setRecentProjects: (projects: RecentProject[]) => void;
  setAnalysisStatus: (s: AnalysisStatus) => void;
  openProject: (project: Project) => void;
  closeProject: () => void;
  addDataSource: (source: DataSource) => void;
  removeDataSource: (sourceId: string) => void;
  startAnalysisPolling: () => void;
  stopAnalysisPolling: () => void;
}

let _pollingInterval: ReturnType<typeof setInterval> | null = null;

export const useProjectStore = create<ProjectState>((set, get) => ({
  currentProject: null,
  recentProjects: [],
  analysisStatus: "idle",
  setCurrentProject: (project) => set({ currentProject: project }),
  setRecentProjects: (projects) => set({ recentProjects: projects }),
  setAnalysisStatus: (s) => set({ analysisStatus: s }),
  openProject: (project) => {
    import("@/stores/fileTreeStore").then(({ useFileTreeStore }) => {
      useFileTreeStore.getState().setRootPath(project.path);
    });
    set({ currentProject: project, analysisStatus: project.analysis_status ?? "idle" });
    if (project.analysis_status === "running") {
      get().startAnalysisPolling();
    }
  },
  closeProject: () => {
    get().stopAnalysisPolling();
    import("@/stores/fileTreeStore").then(({ useFileTreeStore }) => {
      useFileTreeStore.getState().setRootPath("");
    });
    set({ currentProject: null, analysisStatus: "idle" });
  },
  addDataSource: (source) =>
    set((state) =>
      state.currentProject
        ? { currentProject: { ...state.currentProject, data_sources: [...state.currentProject.data_sources, source] } }
        : {}
    ),
  removeDataSource: (sourceId) =>
    set((state) =>
      state.currentProject
        ? { currentProject: { ...state.currentProject, data_sources: state.currentProject.data_sources.filter((s) => s.id !== sourceId) } }
        : {}
    ),
  startAnalysisPolling: () => {
    if (_pollingInterval) return;
    _pollingInterval = setInterval(async () => {
      try {
        const { projects } = await import("@/services/api");
        const data = await projects.getAnalysisStatus();
        set({ analysisStatus: data.status });
        if (data.status !== "running") {
          get().stopAnalysisPolling();
          if (data.status === "done") {
            // Refresh project to get updated analysis_status from server
            const { getCurrent } = await import("@/services/api").then(m => ({ getCurrent: m.projects.getCurrent }));
            const result = await getCurrent();
            if (result.project) {
              set({ currentProject: result.project });
            }
          }
        }
      } catch {
        // ignore polling errors
      }
    }, 2000);
  },
  stopAnalysisPolling: () => {
    if (_pollingInterval) {
      clearInterval(_pollingInterval);
      _pollingInterval = null;
    }
  },
}));
