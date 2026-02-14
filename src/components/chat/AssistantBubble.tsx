import { Bot, Copy, Check, RotateCcw, AlertCircle } from "lucide-react";
import { useState } from "react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Tooltip, TooltipContent, TooltipTrigger } from "@/components/ui/tooltip";
import { MarkdownRenderer } from "./MarkdownRenderer";
import { ThinkingBlock } from "./ThinkingBlock";
import { ToolCallsBlock } from "./ToolCallsBlock";
import { ToolInspectModal } from "./ToolInspectModal";
import type { Message, ToolCall } from "@/types";
import { formatMessageForCopy } from "@/lib/formatMessage";

interface Props {
  message: Message;
  isStreaming: boolean;
}

type GroupedBlock =
  | { type: "thinking"; content: string; key: number }
  | { type: "text"; content: string; key: number }
  | { type: "tool_group"; calls: ToolCall[]; key: number };

function groupBlocks(message: Message): GroupedBlock[] {
  const result: GroupedBlock[] = [];
  for (const block of message.blocks ?? []) {
    if (block.type === "tool_call") {
      const last = result[result.length - 1];
      if (last?.type === "tool_group") {
        last.calls.push(block.call);
      } else {
        result.push({ type: "tool_group", calls: [block.call], key: result.length });
      }
    } else {
      result.push({ ...block, key: result.length });
    }
  }
  return result;
}

export function AssistantBubble({ message, isStreaming }: Props) {
  const { t, i18n } = useTranslation();
  const [copied, setCopied] = useState(false);
  const [inspecting, setInspecting] = useState<ToolCall | null>(null);

  const grouped = groupBlocks(message);
  const isEmpty = grouped.length === 0;

  return (
    <div className="flex gap-2.5">
      <div className="flex size-8 shrink-0 items-center justify-center rounded-md bg-gradient-to-br from-primary/30 to-primary text-primary-foreground">
        <Bot className="size-4" />
      </div>
      <div className="min-w-0 max-w-prose flex-1">
        <div className="flex items-center gap-2 mb-1">
          <span className="text-xs font-semibold text-muted-foreground">ExcelAgent</span>
          <span className="text-2xs text-muted-foreground">
            {new Date(message.createdAt).toLocaleTimeString(i18n.language, { hour: "2-digit", minute: "2-digit" })}
          </span>
          {isStreaming && <span className="rounded-full bg-info/10 px-1.5 py-0.5 text-2xs font-medium text-info">{t("chat.processing")}</span>}
        </div>

        <Card className="rounded-lg rounded-tl-sm border-agent-bubble-border bg-agent-bubble px-4_5 py-3.5 shadow-none gap-0">
          {grouped.map((block, i) => {
            const isLast = i === grouped.length - 1;
            if (block.type === "thinking") {
              return <ThinkingBlock key={block.key} thinking={block.content} isStreaming={isStreaming && isLast} />;
            }
            if (block.type === "text") {
              return (
                <div key={block.key} className="msg-text">
                  <MarkdownRenderer content={block.content} />
                  {isStreaming && isLast && <span className="ml-0.5 inline-block size-2 rounded-full bg-primary animate-pulse" />}
                </div>
              );
            }
            return <ToolCallsBlock key={block.key} calls={block.calls} onInspect={setInspecting} />;
          })}

          {isStreaming && isEmpty && !message.error && (
            <span className="inline-block size-2 rounded-full bg-primary animate-pulse" />
          )}

          {message.error && (
            <div className="flex items-start gap-2 text-sm text-destructive">
              <AlertCircle className="mt-0.5 size-4 shrink-0" />
              <span>{message.error}</span>
            </div>
          )}
        </Card>

        {!isStreaming && (!!message.content || grouped.length > 0) && (
          <div className="mt-1.5 flex items-center gap-1">
            <Tooltip>
              <TooltipTrigger asChild>
                <Button variant="ghost" size="icon" className="size-6 text-muted-foreground hover:text-foreground"
                  onClick={async () => { await navigator.clipboard.writeText(formatMessageForCopy(message)); setCopied(true); setTimeout(() => setCopied(false), 2000); }}>
                  {copied ? <Check className="size-3 text-ok" /> : <Copy className="size-3" />}
                </Button>
              </TooltipTrigger>
              <TooltipContent>{t("chat.copy")}</TooltipContent>
            </Tooltip>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button variant="ghost" size="icon" className="size-6 text-muted-foreground hover:text-foreground">
                  <RotateCcw className="size-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent>{t("chat.regenerate")}</TooltipContent>
            </Tooltip>
          </div>
        )}
      </div>

      <ToolInspectModal call={inspecting} onClose={() => setInspecting(null)} />
    </div>
  );
}

