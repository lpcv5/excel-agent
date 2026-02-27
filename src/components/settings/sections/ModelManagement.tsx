import { useState } from "react";
import { Trash2 } from "lucide-react";
import { useTranslation } from "react-i18next";
import { Button } from "@/components/ui/button";
import { useSettingsStore } from "@/stores/settingsStore";
import { AddModelDialog } from "@/components/settings/dialogs/AddModelDialog";
import type { ModelEntry } from "@/services/api";

export function ModelManagement() {
  const { t } = useTranslation();
  const { settings, addModel, removeModel, saveSettings } = useSettingsStore();
  const [dialogOpen, setDialogOpen] = useState(false);

  const handleAdd = (entry: Omit<ModelEntry, "id">) => {
    addModel(entry);
    // saveSettings is called after state update via the store
    setTimeout(() => saveSettings(), 0);
  };

  const handleRemove = (id: string) => {
    removeModel(id);
    setTimeout(() => saveSettings(), 0);
  };

  const models = settings?.models ?? [];

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <h2 className="text-base font-semibold">{t("settings.models.title")}</h2>
        <Button size="sm" onClick={() => setDialogOpen(true)}>
          {t("settings.models.addModel")}
        </Button>
      </div>

      {models.length === 0 ? (
        <p className="text-sm text-muted-foreground">{t("settings.models.empty")}</p>
      ) : (
        <div className="rounded-lg border bg-card divide-y">
          {models.map((m) => (
            <div key={m.id} className="flex items-center justify-between px-4 py-3">
              <div className="min-w-0">
                <p className="text-sm font-medium truncate">
                  {m.display_name || `${m.provider}:${m.model_name}`}
                </p>
                {m.display_name && (
                  <p className="text-xs text-muted-foreground mt-0.5">{m.provider}:{m.model_name}</p>
                )}
                {m.base_url && (
                  <p className="text-xs text-muted-foreground truncate">{m.base_url}</p>
                )}
              </div>
              <Button
                variant="ghost"
                size="icon"
                className="size-7 shrink-0 text-muted-foreground hover:text-destructive"
                onClick={() => handleRemove(m.id)}
              >
                <Trash2 className="size-3.5" />
              </Button>
            </div>
          ))}
        </div>
      )}

      <AddModelDialog
        open={dialogOpen}
        onOpenChange={setDialogOpen}
        onAdd={handleAdd}
      />
    </div>
  );
}
