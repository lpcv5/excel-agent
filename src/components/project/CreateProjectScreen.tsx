import { useState } from "react";
import { useNavigate } from "react-router";
import { ArrowLeft, FolderOpen, Plus, X, FileSpreadsheet, Folder, ChevronDown } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Separator } from "@/components/ui/separator";
import { ScrollArea } from "@/components/ui/scroll-area";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { useProjectStore } from "@/stores/projectStore";
import { projects, fileDialog, fileTree } from "@/services/api";
import type { DataSource, ProjectOptions } from "@/types/project";

function Toggle({ checked, onChange, label }: { checked: boolean; onChange: (v: boolean) => void; label: string }) {
  return (
    <label className="flex items-center justify-between cursor-pointer py-2">
      <span className="text-sm">{label}</span>
      <button
        role="switch"
        aria-checked={checked}
        onClick={() => onChange(!checked)}
        className={`relative inline-flex h-5 w-9 shrink-0 rounded-full border-2 border-transparent transition-colors focus-visible:outline-none ${checked ? "bg-primary" : "bg-input"}`}
      >
        <span className={`pointer-events-none block h-4 w-4 rounded-full bg-background shadow-lg transition-transform ${checked ? "translate-x-4" : "translate-x-0"}`} />
      </button>
    </label>
  );
}

export function CreateProjectScreen() {
  const { t } = useTranslation();
  const { openProject } = useProjectStore();
  const navigate = useNavigate();
  const [projectPath, setProjectPath] = useState("");
  const [projectName, setProjectName] = useState("");
  const [dataSources, setDataSources] = useState<DataSource[]>([]);
  const [options, setOptions] = useState<ProjectOptions>({
    data_cleaning_enabled: true,
    auto_save_memory: true,
    show_hidden_files: false,
  });
  const [creating, setCreating] = useState(false);
  const [error, setError] = useState<string | null>(null);

  async function handleBrowseFolder() {
    const path = await fileTree.openFolderDialog();
    if (!path) return;
    setProjectPath(path);
    if (!projectName) {
      const parts = path.replace(/\\/g, "/").split("/");
      setProjectName(parts[parts.length - 1] || "");
    }
  }

  async function handleAddSources() {
    const paths = await fileDialog.open({ multiple: true });
    if (!paths) return;
    const newSources: DataSource[] = paths.map((p) => {
      const name = p.replace(/\\/g, "/").split("/").pop() ?? p;
      const ext = name.split(".").pop()?.toLowerCase() ?? "";
      const type = ext === "csv" ? "csv" : "excel";
      return { id: "", type, path: p, name } as DataSource;
    });
    setDataSources((prev) => [...prev, ...newSources.filter((s) => !prev.some((e) => e.path === s.path))]);
  }

  async function handleAddFolder() {
    const path = await fileTree.openFolderDialog();
    if (!path) return;
    const name = path.replace(/\\/g, "/").split("/").pop() ?? path;
    setDataSources((prev) => prev.some((s) => s.path === path) ? prev : [...prev, { id: "", type: "folder", path, name }]);
  }

  function removeSource(id: string) {
    setDataSources((prev) => prev.filter((s) => s.id !== id));
  }

  async function handleCreate() {
    if (!projectPath) return;
    setCreating(true);
    setError(null);
    try {
      const { project } = await projects.create({
        project_path: projectPath,
        name: projectName || projectPath.split(/[\\/]/).pop() || "Project",
        data_sources: dataSources,
        options,
      });
      openProject(project);
      navigate("/project");
    } catch (err) {
      setError(String(err));
      setCreating(false);
    }
  }

  return (
    <div className="flex h-full flex-col bg-background">
      <div className="flex items-center gap-3 border-b px-6 py-3">
        <Button variant="ghost" size="sm" className="gap-1.5" onClick={() => navigate("/")}>
          <ArrowLeft className="size-4" />
          {t("createProject.back")}
        </Button>
        <span className="text-sm font-semibold">{t("createProject.title")}</span>
      </div>

      <ScrollArea className="flex-1">
        <div className="mx-auto max-w-[600px] space-y-6 px-6 py-6">

          {/* Section 1 — Project Folder */}
          <section className="space-y-3">
            <h2 className="text-sm font-semibold">{t("createProject.folderSection")}</h2>
            <div className="flex gap-2">
              <Input
                placeholder={t("createProject.folderPlaceholder")}
                value={projectPath}
                onChange={(e) => setProjectPath(e.target.value)}
                className="flex-1 font-mono text-xs"
              />
              <Button variant="outline" size="sm" className="gap-1.5 shrink-0" onClick={handleBrowseFolder}>
                <FolderOpen className="size-4" />
                {t("createProject.browse")}
              </Button>
            </div>
            <div className="space-y-1">
              <label className="text-xs text-muted-foreground">{t("createProject.projectName")}</label>
              <Input
                placeholder={t("createProject.projectNamePlaceholder")}
                value={projectName}
                onChange={(e) => setProjectName(e.target.value)}
              />
            </div>
          </section>

          <Separator />

          {/* Section 2 — Data Sources */}
          <section className="space-y-3">
            <h2 className="text-sm font-semibold">{t("createProject.dataSourcesSection")}</h2>
            {dataSources.length > 0 && (
              <div className="rounded-md border divide-y">
                {dataSources.map((s) => (
                  <div key={s.id} className="flex items-center gap-2 px-3 py-2">
                    {s.type === "folder"
                      ? <Folder className="size-4 shrink-0 text-muted-foreground" />
                      : <FileSpreadsheet className="size-4 shrink-0 text-muted-foreground" />
                    }
                    <div className="flex-1 min-w-0">
                      <div className="text-sm truncate">{s.name}</div>
                      <div className="text-xs text-muted-foreground truncate">{s.path}</div>
                    </div>
                    <span className="text-xs text-muted-foreground uppercase shrink-0">{s.type}</span>
                    <button
                      onClick={() => removeSource(s.id)}
                      className="p-0.5 rounded hover:bg-surface-hover"
                      aria-label="Remove"
                    >
                      <X className="size-3.5 text-muted-foreground" />
                    </button>
                  </div>
                ))}
              </div>
            )}
            <div className="flex items-center gap-3">
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" size="sm" className="gap-1.5">
                    <Plus className="size-3.5" />
                    {t("createProject.addSource")}
                    <ChevronDown className="size-3.5" />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="start">
                  <DropdownMenuItem onClick={handleAddSources}>
                    <FileSpreadsheet className="size-4" />
                    {t("createProject.sourceFiles")}
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={handleAddFolder}>
                    <FolderOpen className="size-4" />
                    {t("createProject.sourceFolder")}
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
              <span className="text-xs text-muted-foreground">{t("createProject.sourceHint")}</span>
            </div>
          </section>

          <Separator />

          {/* Section 3 — Options */}
          <section className="space-y-1">
            <h2 className="text-sm font-semibold mb-2">{t("createProject.optionsSection")}</h2>
            <Toggle
              label={t("createProject.dataCleaning")}
              checked={options.data_cleaning_enabled}
              onChange={(v) => setOptions((o) => ({ ...o, data_cleaning_enabled: v }))}
            />
            <Toggle
              label={t("createProject.autoSave")}
              checked={options.auto_save_memory}
              onChange={(v) => setOptions((o) => ({ ...o, auto_save_memory: v }))}
            />
            <Toggle
              label={t("createProject.showHidden")}
              checked={options.show_hidden_files}
              onChange={(v) => setOptions((o) => ({ ...o, show_hidden_files: v }))}
            />
          </section>

          {error && (
            <p className="text-sm text-destructive rounded-md border border-destructive/30 bg-destructive/10 px-3 py-2">{error}</p>
          )}
        </div>
      </ScrollArea>

      <div className="border-t px-6 py-3 flex justify-end">
        <Button disabled={!projectPath || creating} onClick={handleCreate}>
          {creating ? t("createProject.creating") : t("createProject.create")}
        </Button>
      </div>
    </div>
  );
}
