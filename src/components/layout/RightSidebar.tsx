import { Plus, Search, MessageSquare } from "lucide-react";
import { useTranslation } from "react-i18next";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Separator } from "@/components/ui/separator";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { useChatStore } from "@/stores/chatStore";
import { cn } from "@/lib/utils";

import type { TFunction } from "i18next";

function formatTime(iso: string, t: TFunction, locale: string) {
  const now = Date.now();
  const diffMin = Math.floor((now - new Date(iso).getTime()) / 60000);
  if (diffMin < 1) return t("rightSidebar.justNow");
  if (diffMin < 60) return t("rightSidebar.minutesAgo", { count: diffMin });
  const diffHr = Math.floor(diffMin / 60);
  if (diffHr < 24) return t("rightSidebar.hoursAgo", { count: diffHr });
  return new Date(iso).toLocaleDateString(locale);
}

export function RightSidebar() {
  const { t, i18n } = useTranslation();
  const { conversations, activeConversationId, setActiveConversation, createConversation, streaming } = useChatStore();

  return (
    <aside className="h-full bg-surface overflow-hidden">
      <div className="flex h-full flex-col">
        <div className="flex items-center justify-between px-3 py-2.5">
          <span className="text-xs-plus font-semibold uppercase tracking-wide text-muted-foreground">{t("rightSidebar.header")}</span>
          <Button
            variant="ghost"
            size="icon"
            className="size-6 text-muted-foreground hover:text-foreground"
            onClick={() => createConversation(t("conversation.defaultTitle"))}
            disabled={streaming}
          >
            <Plus className="size-3.5" />
          </Button>
        </div>
        <Separator />

        {/* Search */}
        <div className="px-2.5 py-2">
          <div className="flex items-center gap-2 rounded-md border bg-surface px-2.5 py-1.5">
            <Search className="size-3 text-muted-foreground shrink-0" />
            <Input className="h-auto flex-1 border-0 bg-transparent p-0 text-xs-plus shadow-none placeholder:text-muted-foreground focus-visible:ring-0" placeholder={t("rightSidebar.searchPlaceholder")} />
          </div>
        </div>
        <Separator />

        <ScrollArea className="flex-1">
          <div className="p-2">
            {conversations.length === 0 && (
              <p className="px-3 py-6 text-center text-xs-plus text-muted-foreground">{t("rightSidebar.noConversations")}</p>
            )}
            {conversations.map((c) => {
              const isActive = c.id === activeConversationId;
              return (
                <div
                  key={c.id}
                  className={cn(
                    "mb-1 cursor-pointer rounded-md border p-2.5 transition-all",
                    isActive ? "border-primary/40 bg-surface-active" : "border-transparent hover:bg-surface-hover"
                  )}
                  onClick={() => !streaming && setActiveConversation(c.id)}
                >
                  <div className="flex items-center gap-1.5 mb-1">
                  <MessageSquare className="size-3.5 shrink-0 text-muted-foreground" />
                    <span className="flex-1 truncate text-xs-plus font-semibold text-foreground">{c.title}</span>
                  </div>
                  <div className="flex items-center justify-between text-2xs text-muted-foreground">
                    <span>{t("rightSidebar.messageCount", { count: c.messageCount })} · {formatTime(c.updatedAt, t, i18n.language)}</span>
                    <span className={cn("rounded-full px-1.5 py-0.5 text-2xs font-medium", isActive ? "bg-ok/10 text-ok" : "bg-surface text-muted-foreground")}>
                      {isActive ? t("rightSidebar.statusActive") : t("rightSidebar.statusDone")}
                    </span>
                  </div>
                </div>
              );
            })}
          </div>
        </ScrollArea>
      </div>
    </aside>
  );
}
