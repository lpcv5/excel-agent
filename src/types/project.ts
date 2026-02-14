export interface DataSource {
  id: string;
  type: "excel" | "csv" | "folder";
  path: string;
  name: string;
}

export interface ProjectOptions {
  data_cleaning_enabled: boolean;
  auto_save_memory: boolean;
  show_hidden_files: boolean;
}

export type AnalysisStatus = "idle" | "running" | "done" | "error";

export interface Project {
  name: string;
  path: string;
  created_at: string;
  modified_at: string;
  data_sources: DataSource[];
  options: ProjectOptions;
  analysis_status: AnalysisStatus;
  analysis_completed_at: string | null;
  data_sources_hash: string | null;
}

export interface RecentProject {
  path: string;
  name: string;
  last_opened: string;
}
