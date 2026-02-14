import { ChevronDown, Plus, FileSpreadsheet, Folder, X, RefreshCw } from "lucide-react";
import { useTranslation } from "react-i18next";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import { Badge } from "@/components/ui/badge";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from "@/components/ui/tooltip";
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from "@/components/ui/dropdown-menu";
import { FileTree } from "@/components/ui/file-tree";
import { useProjectStore } from "@/stores/projectStore";
import { projects, fileDialog, fileTree } from "@/services/api";
import type { DataSource } from "@/types/project";

function Section({ title, count, action, children }: { title: string; count?: number; action?: React.ReactNode; children: React.ReactNode }) {
  return (
    <Collapsible defaultOpen className="border-b last:border-b-0">
      <CollapsibleTrigger className="flex w-full items-center gap-1.5 px-3 py-2 hover:bg-surface-hover transition-colors [&[data-state=closed]>svg:first-child]:-rotate-90">
        <ChevronDown className="size-3 text-muted-foreground transition-transform duration-200" />
        <span className="flex-1 text-left text-xs-plus font-semibold text-muted-foreground">{title}</span>
        {count !== undefined && <Badge className="rounded-full bg-surface px-1.5 text-2xs text-muted-foreground border-0 h-auto">{count}</Badge>}
        {action && <span onClick={(e) => e.stopPropagation()}>{action}</span>}
      </CollapsibleTrigger>
      <CollapsibleContent className="pb-1">{children}</CollapsibleContent>
    </Collapsible>
  );
}

export function LeftSidebar() {
  const { t } = useTranslation();
  const currentProject = useProjectStore((s) => s.currentProject);
  const analysisStatus = useProjectStore((s) => s.analysisStatus);
  const dataSources = currentProject?.data_sources ?? [];

  const showUpdateMemory = dataSources.length > 0 && (analysisStatus === "idle" || analysisStatus === "error");

  async function handleUpdateMemory() {
    await projects.triggerAnalysis();
    useProjectStore.getState().setAnalysisStatus("running");
    useProjectStore.getState().startAnalysisPolling();
  }

  async function handleAddFiles() {
    const paths = await fileDialog.open({ multiple: true });
    if (!paths) return;
    for (const p of paths) {
      if (dataSources.some((s) => s.path === p)) continue;
      const name = p.replace(/\\/g, "/").split("/").pop() ?? p;
      const ext = name.split(".").pop()?.toLowerCase() ?? "";
      const type: DataSource["type"] = ext === "csv" ? "csv" : "excel";
      try {
        const { project } = await projects.addDataSource({ type, path: p, name });
        useProjectStore.getState().setCurrentProject(project);
      } catch (err) {
        console.error("Failed to add data source:", err);
      }
    }
  }

  async function handleAddFolder() {
    const path = await fileTree.openFolderDialog();
    if (!path) return;
    if (dataSources.some((s) => s.path === path)) return;
    const name = path.replace(/\\/g, "/").split("/").pop() ?? path;
    try {
      const { project } = await projects.addDataSource({ type: "folder", path, name });
      useProjectStore.getState().setCurrentProject(project);
    } catch (err) {
      console.error("Failed to add folder source:", err);
    }
  }

  async function handleRemove(id: string) {
    try {
      const { project } = await projects.removeDataSource(id);
      useProjectStore.getState().setCurrentProject(project);
    } catch (err) {
      console.error("Failed to remove data source:", err);
    }
  }

  return (
    <aside className="h-full bg-surface overflow-hidden">
      <div className="flex h-full flex-col">
        <div className="flex items-center justify-between px-3 py-2.5">
          <span className="text-xs-plus font-semibold uppercase tracking-wide text-muted-foreground">{t("leftSidebar.header")}</span>
        </div>
        <Separator />

        <ScrollArea className="flex-1">
          <Section
            title={t("leftSidebar.dataSources")}
            count={dataSources.length}
            action={showUpdateMemory ? (
              <TooltipProvider delayDuration={500}>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <button
                      className="p-0.5 rounded hover:bg-surface-active transition-colors"
                      onClick={handleUpdateMemory}
                      aria-label={t("leftSidebar.updateMemory")}
                    >
                      <RefreshCw className="size-3 text-muted-foreground" />
                    </button>
                  </TooltipTrigger>
                  <TooltipContent side="right">{t("leftSidebar.updateMemoryTooltip")}</TooltipContent>
                </Tooltip>
              </TooltipProvider>
            ) : undefined}
          >
            {dataSources.length === 0 ? (
              <p className="px-3 py-2 text-xs text-muted-foreground">{t("leftSidebar.noDataSources")}</p>
            ) : (
              <TooltipProvider delayDuration={500}>
                {dataSources.map((d) => (
                  <Tooltip key={d.id}>
                    <TooltipTrigger asChild>
                      <div className="group flex items-center gap-2 px-3 py-1.5 text-xs-plus cursor-pointer transition-colors hover:bg-surface-hover">
                        {d.type === "folder"
                          ? <Folder className="size-3.5 shrink-0 text-muted-foreground" />
                          : <FileSpreadsheet className="size-3.5 shrink-0 text-muted-foreground" />
                        }
                        <span className="flex-1 truncate text-muted-foreground">{d.name}</span>
                        <span className="text-2xs text-muted-foreground uppercase">{d.type}</span>
                        <button
                          className="opacity-0 group-hover:opacity-100 transition-opacity p-0.5 rounded hover:bg-surface-active"
                          onClick={(e) => { e.stopPropagation(); handleRemove(d.id); }}
                          aria-label="Remove"
                        >
                          <X className="size-3 text-muted-foreground" />
                        </button>
                      </div>
                    </TooltipTrigger>
                    <TooltipContent side="right">{d.path}</TooltipContent>
                  </Tooltip>
                ))}
              </TooltipProvider>
            )}
            <DropdownMenu>
              <DropdownMenuTrigger asChild>
                <Button variant="ghost" size="sm" className="w-full justify-start gap-1.5 px-3 h-auto py-1.5 text-xs-plus text-primary">
                  <Plus className="size-3" />{t("leftSidebar.addDataSource")}
                </Button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="start">
                <DropdownMenuItem onClick={handleAddFiles}>
                  <FileSpreadsheet className="size-4" />
                  {t("createProject.sourceFiles")}
                </DropdownMenuItem>
                <DropdownMenuItem onClick={handleAddFolder}>
                  <Folder className="size-4" />
                  {t("createProject.sourceFolder")}
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          </Section>

          <Section title={t("leftSidebar.projectFiles")}>
            <FileTree />
          </Section>
        </ScrollArea>

      </div>
    </aside>
  );
}
