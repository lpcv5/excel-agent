import { useTranslation } from "react-i18next";
import { useChatStore } from "@/stores/chatStore";
import { useProjectStore } from "@/stores/projectStore";

export function Footer() {
  const { t } = useTranslation();
  const streaming = useChatStore((s) => s.streaming);
  const analysisStatus = useProjectStore((s) => s.analysisStatus);
  const estimatedTokens = useChatStore((s) => {
    if (!s.activeConversationId) return 0;
    return (s.messages[s.activeConversationId] ?? []).reduce(
      (acc, m) => acc + Math.ceil(m.content.length / 2), 0
    );
  });

  const statusText = streaming
    ? t("footer.generating")
    : analysisStatus === "running"
    ? t("footer.analyzingDataSources")
    : t("footer.connected");

  const dotClass = streaming
    ? "bg-warn animate-pulse"
    : analysisStatus === "running"
    ? "bg-blue-400 animate-pulse"
    : "bg-ok animate-pulse-dot";

  return (
    <footer className="flex h-footer shrink-0 items-center border-t bg-background px-4 text-2xs text-muted-foreground">
      <div className="flex items-center gap-1.5">
        <span className={`size-1.5 rounded-full ${dotClass}`} />
        <span>{statusText}</span>
      </div>
      <div className="flex-1" />
      <div className="flex items-center gap-3">
        <span>Token: ~{estimatedTokens.toLocaleString()} / 200K</span>
        <span>v0.1.0</span>
      </div>
    </footer>
  );
}
