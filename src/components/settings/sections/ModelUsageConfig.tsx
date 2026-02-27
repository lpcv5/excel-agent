import { useTranslation } from "react-i18next";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useSettingsStore } from "@/stores/settingsStore";

const NONE = "__none__";

export function ModelUsageConfig() {
  const { t } = useTranslation();
  const { settings, setMainModelId, setSubagentsModelId, setAnalysisModelId, saveSettings } = useSettingsStore();

  const models = settings?.models ?? [];

  const handleChange = (setter: (id: string | null) => void) => (val: string) => {
    setter(val === NONE ? null : val);
    setTimeout(() => saveSettings(), 0);
  };

  const modelLabel = (id: string | null) => {
    if (!id) return NONE;
    const m = models.find((m) => m.id === id);
    return m ? m.id : NONE;
  };

  const rows = [
    {
      label: t("settings.usage.mainModel"),
      desc: t("settings.usage.mainModelDesc"),
      value: modelLabel(settings?.main_model_id ?? null),
      onChange: handleChange(setMainModelId),
    },
    {
      label: t("settings.usage.subagentsModel"),
      desc: t("settings.usage.subagentsModelDesc"),
      value: modelLabel(settings?.subagents_model_id ?? null),
      onChange: handleChange(setSubagentsModelId),
    },
    {
      label: t("settings.usage.analysisModel"),
      desc: t("settings.usage.analysisModelDesc"),
      value: modelLabel(settings?.analysis_model_id ?? null),
      onChange: handleChange(setAnalysisModelId),
    },
  ];

  return (
    <div>
      <h2 className="text-base font-semibold mb-6">{t("settings.usage.title")}</h2>
      <div className="rounded-lg border bg-card divide-y">
        {rows.map((row) => (
          <div key={row.label} className="flex items-center justify-between px-4 py-3">
            <div>
              <p className="text-sm font-medium">{row.label}</p>
              <p className="text-xs text-muted-foreground mt-0.5">{row.desc}</p>
            </div>
            <Select value={row.value} onValueChange={row.onChange}>
              <SelectTrigger className="w-52 h-8 text-sm">
                <SelectValue placeholder={t("settings.usage.notSelected")} />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value={NONE} className="text-sm text-muted-foreground">
                  {t("settings.usage.notSelected")}
                </SelectItem>
                {models.map((m) => (
                  <SelectItem key={m.id} value={m.id} className="text-sm">
                    {m.display_name || `${m.provider}:${m.model_name}`}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
        ))}
      </div>
    </div>
  );
}
