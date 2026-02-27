import { PanelLeft, PanelRight, Plus, Search } from "lucide-react";
import { Button } from "@/components/ui/button";
import { ThemeToggle } from "@/components/theme/ThemeToggle";
import { useLayout } from "@/hooks/useLayout";
import { useChatStore } from "@/stores/chatStore";

export function Header() {
  const { toggleLeftSidebar, toggleRightSidebar } = useLayout();
  const { createConversation, streaming } = useChatStore();
  const activeConv = useChatStore((s) => s.conversations.find((c) => c.id === s.activeConversationId));

  return (
    <header className="flex h-header shrink-0 items-center border-b bg-background px-3 gap-1">
      <Button variant="ghost" size="icon-xs" onClick={toggleLeftSidebar}>
        <PanelLeft className="size-4" />
      </Button>
      <Button variant="ghost" size="sm" className="h-7 gap-1.5 text-xs-minus" onClick={() => createConversation()} disabled={streaming}>
        <Plus className="size-3.5" />新对话
      </Button>
      <div className="flex flex-1 items-center justify-center gap-2">
        <span className="text-sm-minus font-bold tracking-tight text-primary select-none">ExcelAgent</span>
        {activeConv && (
          <>
            <span className="text-muted-foreground/40">·</span>
            <span className="max-w-48 truncate text-xs-plus text-muted-foreground">{activeConv.title}</span>
          </>
        )}
      </div>
      <Button variant="ghost" size="icon-xs">
        <Search className="size-3.5" />
      </Button>
      <ThemeToggle />
      <Button variant="ghost" size="icon-xs" onClick={toggleRightSidebar}>
        <PanelRight className="size-4" />
      </Button>
    </header>
  );
}
