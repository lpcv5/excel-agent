import { User } from "lucide-react";
import { Card } from "@/components/ui/card";
import type { Message } from "@/types";

export function UserBubble({ message }: { message: Message }) {
  return (
    <div className="flex justify-end gap-2.5">
      <div className="max-w-prose">
        <Card className="rounded-lg rounded-tr-sm border-user-bubble-border bg-user-bubble px-4_5 py-3.5 text-user-bubble-foreground shadow-none gap-0">
          <p className="whitespace-pre-wrap text-body leading-relaxed">{message.content}</p>
        </Card>
        <div className="mt-1 text-right">
          <time className="text-2xs text-muted-foreground">
            {new Date(message.createdAt).toLocaleTimeString("zh-CN", { hour: "2-digit", minute: "2-digit" })}
          </time>
        </div>
      </div>
      <div className="flex size-8 shrink-0 items-center justify-center rounded-md bg-secondary text-muted-foreground">
        <User className="size-3.5" />
      </div>
    </div>
  );
}
