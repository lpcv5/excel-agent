import { useTranslation } from "react-i18next";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";

export function UILanguage() {
  const { t, i18n } = useTranslation();

  return (
    <div>
      <h2 className="text-base font-semibold mb-6">{t("settings.language")}</h2>
      <div className="rounded-lg border bg-card divide-y">
        <div className="flex items-center justify-between px-4 py-3">
          <div>
            <p className="text-sm font-medium">{t("settings.language")}</p>
            <p className="text-xs text-muted-foreground mt-0.5">{t("settings.languageDesc")}</p>
          </div>
          <Select value={i18n.language} onValueChange={(lang) => i18n.changeLanguage(lang)}>
            <SelectTrigger className="w-44 h-8 text-sm"><SelectValue /></SelectTrigger>
            <SelectContent>
              <SelectItem value="zh-CN" className="text-sm">中文（简体）</SelectItem>
              <SelectItem value="en-US" className="text-sm">English</SelectItem>
            </SelectContent>
          </Select>
        </div>
      </div>
    </div>
  );
}
