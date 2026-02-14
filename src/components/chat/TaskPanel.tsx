import { useState } from "react";
import { ChevronDown, ClipboardList, Timer, CheckCircle2, Circle } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { cn } from "@/lib/utils";
import type { TaskItem } from "@/types";

interface Props {
  tasks: TaskItem[];
}

export function TaskPanel({ tasks }: Props) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(true);
  if (!tasks.length) return null;

  const doneCount = tasks.filter((task) => task.status === "completed").length;
  const pct = Math.round((doneCount / tasks.length) * 100);

  return (
    <Collapsible open={open} onOpenChange={setOpen} className="border-t bg-surface">
      <div className="flex items-center">
        <CollapsibleTrigger className="flex flex-1 items-center gap-2 px-6 py-2 cursor-pointer hover:bg-surface-hover transition-colors">
          <ChevronDown className={cn("size-3 text-muted-foreground transition-transform duration-200 shrink-0", !open && "-rotate-90")} />
          <ClipboardList className="size-3.5 shrink-0 text-muted-foreground" />
          <span className="text-xs font-semibold text-foreground">{t("chat.taskPlanHeader")}</span>
          <span className="rounded-full bg-primary px-1.5 py-0.5 text-2xs font-semibold text-primary-foreground">{doneCount}/{tasks.length}</span>
          <span className="flex items-center gap-1 text-2xs text-muted-foreground ml-auto"><Timer className="size-3" />{t("chat.taskRunning")}</span>
        </CollapsibleTrigger>
      </div>
      <CollapsibleContent>
        <div className="px-6 pb-3">
          <div className="mb-2.5 h-[3px] overflow-hidden rounded-full bg-border">
            <div className="h-full rounded-full bg-primary transition-all duration-500" style={{ width: `${pct}%` }} />
          </div>
          <div className="flex flex-col gap-1">
            {tasks.map((task) => (
              <div key={task.id} className="flex items-center gap-2 rounded px-2 py-1 text-xs-plus hover:bg-surface-hover transition-colors">
                <span className="flex w-[18px] shrink-0 items-center justify-center">
                  {task.status === "completed"
                    ? <CheckCircle2 className="size-3.5 text-ok" />
                    : task.status === "in_progress"
                    ? <span className="inline-block size-3.5 rounded-full border-2 border-primary border-t-transparent animate-spin" />
                    : <Circle className="size-3.5 text-muted-foreground/40" />}
                </span>
                <span className={cn(
                  "flex-1",
                  task.status === "completed" && "text-muted-foreground line-through",
                  task.status === "in_progress" && "font-medium text-foreground",
                  task.status === "pending" && "text-muted-foreground",
                )}>
                  {task.label}
                </span>
                <span className="font-mono text-2xs">
                  {task.status === "in_progress"
                    ? <span className="text-info animate-pulse">{t("chat.taskRunning")}</span>
                    : <span className="text-muted-foreground">—</span>}
                </span>
              </div>
            ))}
          </div>
        </div>
      </CollapsibleContent>
    </Collapsible>
  );
}
