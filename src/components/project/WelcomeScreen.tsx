import { useEffect } from "react";
import { useNavigate } from "react-router";
import { FolderOpen, Plus, X, Clock } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";
import { useProjectStore } from "@/stores/projectStore";
import { projects, fileTree } from "@/services/api";
import type { RecentProject } from "@/types/project";

function formatDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, { month: "short", day: "numeric", year: "numeric" });
  } catch {
    return iso;
  }
}

function RecentProjectRow({ project, onOpen, onRemove }: {
  project: RecentProject;
  onOpen: (p: RecentProject) => void;
  onRemove: (p: RecentProject) => void;
}) {
  return (
    <div
      className="group flex items-center gap-3 rounded-md px-3 py-2.5 cursor-pointer hover:bg-surface-hover transition-colors"
      onClick={() => onOpen(project)}
    >
      <FolderOpen className="size-4 shrink-0 text-muted-foreground" />
      <div className="flex-1 min-w-0">
        <div className="text-sm font-medium truncate">{project.name}</div>
        <div className="text-xs text-muted-foreground truncate">{project.path}</div>
      </div>
      <div className="flex items-center gap-2 shrink-0">
        <span className="text-xs text-muted-foreground flex items-center gap-1">
          <Clock className="size-3" />
          {formatDate(project.last_opened)}
        </span>
        <button
          className="opacity-0 group-hover:opacity-100 transition-opacity p-0.5 rounded hover:bg-surface-active"
          onClick={(e) => { e.stopPropagation(); onRemove(project); }}
          aria-label="Remove from recent"
        >
          <X className="size-3.5 text-muted-foreground" />
        </button>
      </div>
    </div>
  );
}

export function WelcomeScreen() {
  const { t } = useTranslation();
  const { openProject, recentProjects, setRecentProjects } = useProjectStore();
  const navigate = useNavigate();

  useEffect(() => {
    projects.getRecent().then(({ recent_projects }) => setRecentProjects(recent_projects));
  }, [setRecentProjects]);

  async function handleOpenProject() {
    const path = await fileTree.openFolderDialog();
    if (!path) return;
    try {
      const { project } = await projects.open(path);
      openProject(project);
      navigate("/project");
    } catch (err) {
      console.error("Failed to open project:", err);
    }
  }

  async function handleOpenRecent(recent: RecentProject) {
    try {
      const { project } = await projects.open(recent.path);
      openProject(project);
      navigate("/project");
    } catch (err) {
      console.error("Failed to open recent project:", err);
    }
  }

  function handleRemoveRecent(recent: RecentProject) {
    setRecentProjects(recentProjects.filter((r) => r.path !== recent.path));
  }

  return (
    <div className="flex h-full items-center justify-center bg-background p-8">
      <div className="w-full max-w-[600px] space-y-6">
        <div className="text-center space-y-1">
          <h1 className="text-2xl font-semibold tracking-tight">{t("welcome.title")}</h1>
          <p className="text-sm text-muted-foreground">{t("welcome.subtitle")}</p>
        </div>

        <div className="flex gap-3">
          <Button className="flex-1 gap-2" onClick={() => navigate("/project/new")}>
            <Plus className="size-4" />
            {t("welcome.newProject")}
          </Button>
          <Button variant="outline" className="flex-1 gap-2" onClick={handleOpenProject}>
            <FolderOpen className="size-4" />
            {t("welcome.openProject")}
          </Button>
        </div>

        {recentProjects.length > 0 && (
          <div className="rounded-lg border bg-surface">
            <div className="px-3 py-2">
              <span className="text-xs font-semibold uppercase tracking-wide text-muted-foreground">{t("welcome.recentProjects")}</span>
            </div>
            <Separator />
            <ScrollArea className="max-h-72">
              <div className="p-1.5 space-y-0.5">
                {recentProjects.map((p) => (
                  <RecentProjectRow
                    key={p.path}
                    project={p}
                    onOpen={handleOpenRecent}
                    onRemove={handleRemoveRecent}
                  />
                ))}
              </div>
            </ScrollArea>
          </div>
        )}
      </div>
    </div>
  );
}
