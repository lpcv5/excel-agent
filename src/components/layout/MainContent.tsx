import { ArrowUp, Paperclip, ArrowDown, Zap, AtSign, BarChart2, Square } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import {
  InputGroup,
  InputGroupAddon,
  InputGroupButton,
  InputGroupTextarea,
} from "@/components/ui/input-group";
import { useState, useRef, useEffect } from "react";
import { useChatStore } from "@/stores/chatStore";
import { useProjectStore } from "@/stores/projectStore";
import { useAutoScroll } from "@/hooks/useAutoScroll";
import { streamMessage, cancelStream, projects } from "@/services/api";
import { UserBubble } from "@/components/chat/UserBubble";
import { AssistantBubble } from "@/components/chat/AssistantBubble";
import { TaskPanel } from "@/components/chat/TaskPanel";
import { useShallow } from "zustand/react/shallow";

export function MainContent() {
  const { t } = useTranslation();
  const suggestions = t("mainContent.suggestions", { returnObjects: true }) as string[];
  const [value, setValue] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const {
    streaming, streamingMessageId, activeConversationId,
    addUserMessage, addAssistantMessage,
    appendToThinking, appendToContent,
    addToolCall, updateToolCall, setTasks,
    setStreamingDone,
  } = useChatStore();

  const messages = useChatStore(
    useShallow((s) => s.activeConversationId ? (s.messages[s.activeConversationId] ?? []) : [])
  );

  // Show tasks from the last assistant message that has tasks — persists after streaming ends
  const activeTasks = useChatStore((s) => {
    if (!s.activeConversationId) return undefined;
    const msgs = s.messages[s.activeConversationId] ?? [];
    for (let i = msgs.length - 1; i >= 0; i--) {
      if (msgs[i].role === "assistant" && msgs[i].tasks?.length) return msgs[i].tasks;
    }
    return undefined;
  });

  const { containerRef, bottomRef, isNearBottom, scrollToBottom, handleScroll } = useAutoScroll([messages.length, streamingMessageId]);

  useEffect(() => {
    const el = textareaRef.current;
    if (!el) return;
    el.style.height = "auto";
    el.style.height = Math.min(el.scrollHeight, 128) + "px";
  }, [value]);

  const handleSend = async (text?: string) => {
    const content = (text ?? value).trim();
    if (!content || streaming) return;

    // Intercept /update-memory slash command
    if (content === "/update-memory") {
      setValue("");
      await projects.triggerAnalysis();
      useProjectStore.getState().setAnalysisStatus("running");
      useProjectStore.getState().startAnalysisPolling();
      return;
    }

    setValue("");
    addUserMessage(content);
    const msg = addAssistantMessage();
    if (!msg) return;

    streamMessage(content, {
      onThinkingToken: (t) => appendToThinking(msg.id, t),
      onContentToken: (t) => appendToContent(msg.id, t),
      onToolStart: (call) => addToolCall(msg.id, call),
      onToolResult: (callId, patch) => updateToolCall(msg.id, callId, patch),
      onTasksUpdate: (tasks) => setTasks(msg.id, tasks),
      onDone: setStreamingDone,
    }, activeConversationId ?? undefined);
  };

  return (
    <main className="relative flex h-full flex-col overflow-hidden">
      <div ref={containerRef} className="flex-1 overflow-y-auto" onScroll={handleScroll}>
        <div className="px-6 py-5">
          {messages.length === 0 ? (
            <Card className="mx-auto max-w-chat flex flex-col items-center justify-center py-20 text-center border-dashed shadow-none bg-transparent">
              <div className="mb-4 flex size-16 items-center justify-center rounded-2xl bg-primary/10">
                <BarChart2 className="size-8 text-primary/60" />
              </div>
              <h2 className="mb-2 text-lg font-semibold">{t("mainContent.welcomeTitle")}</h2>
              <p className="max-w-sm text-sm-minus text-muted-foreground">
                {t("mainContent.welcomeSubtitle")}
              </p>
              <div className="mt-6 flex max-w-md flex-wrap justify-center gap-2">
                {suggestions.map((s) => (
                  <Button key={s} variant="outline" size="sm" className="h-auto rounded-full px-3 py-1.5 text-xs-plus text-muted-foreground hover:border-primary/30 hover:bg-primary/5 hover:text-foreground" onClick={() => handleSend(s)}>
                    {s}
                  </Button>
                ))}
              </div>
            </Card>
          ) : (
            <div className="flex flex-col gap-4">
              {messages.map((msg) =>
                msg.role === "user"
                  ? <UserBubble key={msg.id} message={msg} />
                  : <AssistantBubble key={msg.id} message={msg} isStreaming={streaming && streamingMessageId === msg.id} />
              )}
            </div>
          )}
          <div ref={bottomRef} className="h-1" />
        </div>
      </div>

      {!isNearBottom && messages.length > 0 && (
        <Button variant="outline" size="icon-sm" onClick={() => scrollToBottom()} className="absolute bottom-24 right-6 z-10 rounded-full shadow-md">
          <ArrowDown className="size-4 text-muted-foreground" />
        </Button>
      )}

      {/* Task Panel */}
      {activeTasks && <TaskPanel tasks={activeTasks} />}

      {/* Input */}
      <div className="shrink-0 bg-background px-6 pt-2 pb-4">
        <div className="mx-auto max-w-chat">
          <InputGroup className="rounded-xl bg-card shadow-md focus-within:border-primary/50 focus-within:ring-1 focus-within:ring-primary/20">
            <InputGroupTextarea
              ref={textareaRef}
              className="min-h-[7rem] max-h-64 text-body"
              placeholder={streaming ? t("mainContent.placeholderStreaming") : t("mainContent.placeholder")}
              rows={3}
              value={value}
              onChange={(e) => setValue(e.target.value)}
              onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); handleSend(); } }}
            />
            <InputGroupAddon align="block-end">
              <InputGroupButton size="xs" className="text-muted-foreground">
                <Zap className="size-3" />{t("mainContent.shortcut")}
              </InputGroupButton>
              <InputGroupButton size="xs" className="text-muted-foreground">
                <Paperclip className="size-3" />{t("mainContent.attachment")}
              </InputGroupButton>
              <InputGroupButton size="xs" className="text-muted-foreground">
                <AtSign className="size-3" />{t("mainContent.reference")}
              </InputGroupButton>
              <div className="flex-1" />
              <Select defaultValue="claude-3-5-sonnet">
                <SelectTrigger className="h-6 w-auto gap-1 border-0 bg-transparent px-2 py-0.5 text-2xs text-muted-foreground shadow-none focus:ring-0 [&>svg]:size-3">
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  <SelectItem value="claude-3-5-sonnet" className="text-xs">Claude 3.5 Sonnet</SelectItem>
                  <SelectItem value="claude-3-5-haiku" className="text-xs">Claude 3.5 Haiku</SelectItem>
                  <SelectItem value="claude-opus-4" className="text-xs">Claude Opus 4</SelectItem>
                </SelectContent>
              </Select>
              {streaming ? (
                <InputGroupButton size="icon-xs" variant="outline" className="rounded-full" onClick={() => cancelStream()}>
                  <Square className="size-3.5" />
                </InputGroupButton>
              ) : (
                <InputGroupButton size="icon-xs" variant="default" className="rounded-full" disabled={!value.trim()} onClick={() => handleSend()}>
                  <ArrowUp className="size-3.5" />
                </InputGroupButton>
              )}
            </InputGroupAddon>
          </InputGroup>
          <div className="mt-1 flex gap-3 text-[10px] text-muted-foreground/60">
            <span><kbd className="rounded border px-0.5 font-mono">Enter</kbd> 发送</span>
            <span><kbd className="rounded border px-0.5 font-mono">Shift+Enter</kbd> 换行</span>
            <span><kbd className="rounded border px-0.5 font-mono">/</kbd> 命令</span>
          </div>
        </div>
      </div>
    </main>
  );
}
