import { Copy, Check, ArrowDownToLine, ArrowUpFromLine, Loader2 } from "lucide-react";
import { useState } from "react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "@/components/ui/dialog";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import type { ToolCall } from "@/types";

interface Props {
  call: ToolCall | null;
  onClose: () => void;
}

function syntaxHL(json: string) {
  return json
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"([^"]+)"(?=\s*:)/g, '<span class="text-primary">"$1"</span>')
    .replace(/:\s*"([^"]*?)"/g, ': <span class="text-warn">"$1"</span>')
    .replace(/:\s*(-?\d+\.?\d*)/g, ': <span class="text-info">$1</span>')
    .replace(/:\s*(true|false|null)/g, ': <span class="text-muted-foreground">$1</span>');
}

export function ToolInspectModal({ call, onClose }: Props) {
  const { t } = useTranslation();
  const [copied, setCopied] = useState(false);

  const handleCopy = async (content: string) => {
    await navigator.clipboard.writeText(content);
    setCopied(true);
    setTimeout(() => setCopied(false), 1500);
  };

  const argsJson = call?.args ? JSON.stringify(call.args, null, 2) : null;

  return (
    <Dialog open={!!call} onOpenChange={(open) => !open && onClose()}>
      <DialogContent className="flex max-h-[75vh] w-[560px] max-w-[90vw] flex-col gap-0 p-0">
        <DialogHeader className="border-b px-4 py-3">
          <DialogTitle className="flex items-center gap-2.5">
            {call && (
              <span className={`flex size-5 items-center justify-center rounded-full ${call.status === "success" ? "bg-ok/10 text-ok" : call.status === "running" ? "bg-info/10 text-info" : "bg-err/10 text-err"}`}>
                {call.status === "running" ? <Loader2 className="size-3 animate-spin" /> : call.status === "success" ? <Check className="size-3" /> : <span className="text-2xs font-bold">✕</span>}
              </span>
            )}
            <span className="font-mono text-sm">{call?.name}</span>
          </DialogTitle>
        </DialogHeader>

        {call && (
          <Tabs defaultValue="params" className="flex min-h-0 flex-1 flex-col">
            <TabsList className="h-auto w-full justify-start rounded-none border-b bg-surface px-4 py-0">
              <TabsTrigger value="params" className="gap-1.5 rounded-none border-b-2 border-transparent px-3 py-2 text-xs-plus data-[state=active]:border-primary data-[state=active]:bg-transparent data-[state=active]:shadow-none">
                <ArrowDownToLine className="size-3" />{t("chat.tabParams")}
              </TabsTrigger>
              <TabsTrigger value="result" disabled={!call.result} className="gap-1.5 rounded-none border-b-2 border-transparent px-3 py-2 text-xs-plus data-[state=active]:border-primary data-[state=active]:bg-transparent data-[state=active]:shadow-none">
                <ArrowUpFromLine className="size-3" />{t("chat.tabResult")}
              </TabsTrigger>
            </TabsList>

            {(["params", "result"] as const).map((tab) => {
              const content = tab === "params" ? argsJson : call.result;
              return (
                <TabsContent key={tab} value={tab} className="mt-0 flex-1 overflow-y-auto p-4">
                  <Card className="mb-3 flex flex-wrap gap-4 rounded-md py-3 px-3 text-2xs">
                    {[
                      [t("chat.fieldName"), call.name],
                      [t("chat.fieldStatus"), call.status],
                      [t("chat.fieldDuration"), call.duration_ms != null ? `${(call.duration_ms / 1000).toFixed(2)}s` : "—"],
                    ].map(([l, v]) => (
                      <div key={l} className="flex flex-col gap-0.5">
                        <span className="text-muted-foreground">{l}</span>
                        <span className="font-medium text-foreground">{v}</span>
                      </div>
                    ))}
                  </Card>
                  {content ? (
                    <pre className="rounded-md bg-surface p-3 font-mono text-xs leading-relaxed overflow-x-auto" dangerouslySetInnerHTML={{ __html: syntaxHL(content) }} />
                  ) : (
                    <div className="flex items-center justify-center gap-1.5 py-8 text-xs-plus text-muted-foreground">
                      <Loader2 className="size-3.5 animate-spin" />{t("chat.waitingResult")}
                    </div>
                  )}
                </TabsContent>
              );
            })}

            <div className="flex justify-end gap-2 border-t px-4 py-2.5">
              <Button variant="outline" size="sm" className="h-7 text-xs-plus" onClick={onClose}>{t("chat.close")}</Button>
              <Button size="sm" className="h-7 text-xs-plus" onClick={() => argsJson && handleCopy(argsJson)} disabled={!argsJson}>
                {copied ? <><Check className="mr-1.5 size-3" />{t("chat.copied")}</> : <><Copy className="mr-1.5 size-3" />{t("chat.copyJson")}</>}
              </Button>
            </div>
          </Tabs>
        )}
      </DialogContent>
    </Dialog>
  );
}
