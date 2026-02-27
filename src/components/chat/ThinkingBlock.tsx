import { useState, useEffect } from "react";
import { ChevronDown, Brain } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from "@/components/ui/collapsible";
import { cn } from "@/lib/utils";

interface Props {
  thinking: string;
  isStreaming: boolean;
}

export function ThinkingBlock({ thinking, isStreaming }: Props) {
  const { t } = useTranslation();
  const [open, setOpen] = useState(isStreaming);

  useEffect(() => { setOpen(isStreaming); }, [isStreaming]);

  if (!thinking) return null;

  return (
    <Collapsible open={open} onOpenChange={setOpen} className="my-2 rounded-md border border-think-border bg-think overflow-hidden">
      <CollapsibleTrigger className="flex w-full items-center gap-2 px-3 py-[7px] cursor-pointer hover:bg-surface-hover transition-colors">
        <ChevronDown className={cn("size-3 text-muted-foreground transition-transform duration-200 shrink-0", !open && "-rotate-90")} />
        <Brain className="size-3.5 shrink-0 text-think-foreground/70" />
        <span className="text-xs font-medium italic text-think-foreground">
          {isStreaming ? t("chat.thinkingInProgress") : t("chat.thinkingDone")}
        </span>
      </CollapsibleTrigger>
      <CollapsibleContent>
        <div className="border-t border-think-border px-3.5 py-2.5 max-h-40 overflow-y-auto">
          <p className="whitespace-pre-wrap text-xs-plus leading-relaxed italic text-think-foreground">
            {thinking}
            {isStreaming && <span className="inline-block size-1.5 ml-0.5 rounded-full bg-think-foreground/50 animate-pulse" />}
          </p>
        </div>
      </CollapsibleContent>
    </Collapsible>
  );
}
