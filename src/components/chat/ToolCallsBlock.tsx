import { useState } from "react";
import { ChevronDown, Wrench, Check, X } from "lucide-react";
import { useTranslation } from "react-i18next";
import { cn } from "@/lib/utils";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import type { ToolCall } from "@/types";

interface Props {
  calls: ToolCall[];
  onInspect: (call: ToolCall) => void;
}

export function ToolCallsBlock({ calls, onInspect }: Props) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(false);
  const running = calls.some((c) => c.status === "running");

  return (
    <Card className={cn("my-2 gap-0 overflow-hidden rounded-lg py-0", running && "border-info/40")}>
      <button
        className="flex w-full items-center gap-2 px-3 py-2 text-left hover:bg-surface-hover transition-colors"
        onClick={() => setOpen((p) => !p)}
      >
        <ChevronDown className={cn("size-3 text-muted-foreground transition-transform duration-200 shrink-0", open && "rotate-180")} />
        <Wrench className="size-3.5 text-brand-muted shrink-0" />
        <span className="text-xs-plus font-semibold text-muted-foreground">{t("chat.toolCallsHeader")}</span>
        <span className="rounded-full bg-surface px-1.5 py-0.5 text-2xs font-semibold text-muted-foreground border-0">×{calls.length}</span>
        <div className="flex-1" />
        {running ? (
          <span className="text-2xs text-info animate-pulse shrink-0">{t("chat.toolRunning")}</span>
        ) : (
          <span className="text-2xs text-muted-foreground shrink-0">
            {t("chat.toolCompleted", { done: calls.filter((c) => c.status === "success").length, total: calls.length })}
          </span>
        )}
      </button>

      {open && (
        <div className="border-t">
          {calls.map((c) => (
            <div key={c.id} className="flex items-center gap-2 border-b px-3 py-2 text-xs-plus last:border-b-0 hover:bg-surface-hover">
              {c.status === "running" ? (
                <span className="size-4 rounded-full border-2 border-info border-t-transparent animate-spin shrink-0" />
              ) : (
                <span className={cn(
                  "flex size-4 items-center justify-center rounded-full text-2xs shrink-0",
                  c.status === "success" && "bg-ok/10 text-ok",
                  c.status === "error" && "bg-err/10 text-err",
                )}>
                  {c.status === "success" ? <Check className="size-2.5" /> : <X className="size-2.5" />}
                </span>
              )}
              <span className="font-mono font-semibold text-foreground">{c.name}</span>
              <span className="flex-1 truncate text-muted-foreground">
                {c.status === "running" ? t("chat.toolStatusRunning") : c.result ? t("chat.toolStatusDone") : t("chat.toolStatusWaiting")}
              </span>
              {c.duration_ms != null && (
                <span className="font-mono text-2xs text-muted-foreground shrink-0">
                  {(c.duration_ms / 1000).toFixed(1)}s
                </span>
              )}
              <Button variant="outline" size="pill" onClick={() => onInspect(c)}>
                详情
              </Button>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}
