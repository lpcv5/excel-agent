import { useTranslation } from "react-i18next";

export function AboutSection() {
  const { t } = useTranslation();
  const version = (import.meta as Record<string, unknown> & { env?: Record<string, string> }).env?.VITE_APP_VERSION ?? "0.1.0";

  return (
    <div>
      <h2 className="text-base font-semibold mb-6">{t("settings.about.title")}</h2>
      <div className="rounded-lg border bg-card divide-y">
        <div className="flex items-center justify-between px-4 py-3">
          <p className="text-sm font-medium">{t("settings.about.appName")}</p>
          <span className="text-sm text-muted-foreground">{version}</span>
        </div>
      </div>
    </div>
  );
}
