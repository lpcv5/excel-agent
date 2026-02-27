import { useEffect, useState } from "react";
import { useTranslation } from "react-i18next";
import { useSettingsStore } from "@/stores/settingsStore";
import { ModelManagement } from "./sections/ModelManagement";
import { ModelUsageConfig } from "./sections/ModelUsageConfig";
import { UILanguage } from "./sections/UILanguage";
import { AboutSection } from "./sections/AboutSection";

type NavSection = "models" | "usage" | "language" | "about";

interface SettingsViewProps {
  onClose?: () => void;
}

export function SettingsView({ onClose }: SettingsViewProps) {
  void onClose;
  const { t } = useTranslation();
  const { fetchSettings, loading } = useSettingsStore();
  const [active, setActive] = useState<NavSection>("models");

  useEffect(() => {
    fetchSettings();
  }, [fetchSettings]);

  const navItems: { key: NavSection; label: string }[] = [
    { key: "models", label: t("settings.nav.models") },
    { key: "usage", label: t("settings.nav.usage") },
    { key: "language", label: t("settings.nav.language") },
    { key: "about", label: t("settings.nav.about") },
  ];

  if (loading) {
    return (
      <main className="flex h-full items-center justify-center text-muted-foreground text-sm">
        {t("settings.loading")}
      </main>
    );
  }

  return (
    <main className="flex h-full overflow-hidden">
      <nav className="w-44 shrink-0 border-r p-3 flex flex-col gap-1">
        {navItems.map((item) => (
          <button
            key={item.key}
            onClick={() => setActive(item.key)}
            className={[
              "w-full text-left px-3 py-2 rounded-md text-sm transition-colors",
              active === item.key
                ? "bg-accent text-accent-foreground font-medium"
                : "text-muted-foreground hover:text-foreground hover:bg-accent/50",
            ].join(" ")}
          >
            {item.label}
          </button>
        ))}
      </nav>

      <div className="flex-1 overflow-y-auto p-8">
        {active === "models" && <ModelManagement />}
        {active === "usage" && <ModelUsageConfig />}
        {active === "language" && <UILanguage />}
        {active === "about" && <AboutSection />}
      </div>
    </main>
  );
}
