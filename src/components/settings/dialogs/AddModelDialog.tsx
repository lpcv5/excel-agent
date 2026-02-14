import { useState, useEffect } from "react";
import { useTranslation } from "react-i18next";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogFooter } from "@/components/ui/dialog";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { settings as settingsApi } from "@/services/api";
import type { ModelEntry } from "@/services/api";

interface AddModelDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
  onAdd: (entry: Omit<ModelEntry, "id">) => void;
}

const CUSTOM_PROVIDER = "__custom__";

export function AddModelDialog({ open, onOpenChange, onAdd }: AddModelDialogProps) {
  const { t } = useTranslation();
  const [providers, setProviders] = useState<{ name: string; default_model: string }[]>([]);
  const [provider, setProvider] = useState("zhipu");
  const [modelName, setModelName] = useState("glm-4.7");
  const [apiKey, setApiKey] = useState("");
  const [baseUrl, setBaseUrl] = useState("");
  const [displayName, setDisplayName] = useState("");
  const [customProvider, setCustomProvider] = useState("");

  useEffect(() => {
    settingsApi.getProviders().then(setProviders).catch(() => {});
  }, []);

  const isCustom = provider === CUSTOM_PROVIDER;

  const handleProviderChange = (val: string) => {
    setProvider(val);
    if (val !== CUSTOM_PROVIDER) {
      const found = providers.find((p) => p.name === val);
      if (found) setModelName(found.default_model);
    } else {
      setModelName("");
    }
  };

  const handleSubmit = () => {
    const resolvedProvider = isCustom ? customProvider.trim() : provider;
    if (!resolvedProvider || !modelName.trim()) return;
    onAdd({
      provider: resolvedProvider,
      model_name: modelName.trim(),
      api_key: apiKey,
      base_url: isCustom && baseUrl.trim() ? baseUrl.trim() : null,
      display_name: displayName.trim(),
    });
    // Reset
    setProvider("zhipu");
    setModelName("glm-4.7");
    setApiKey("");
    setBaseUrl("");
    setDisplayName("");
    setCustomProvider("");
    onOpenChange(false);
  };

  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-md">
        <DialogHeader>
          <DialogTitle>{t("settings.models.addModel")}</DialogTitle>
        </DialogHeader>

        <div className="space-y-4 py-2">
          {/* Provider */}
          <div className="space-y-1.5">
            <label className="text-sm font-medium">{t("settings.models.provider")}</label>
            <Select value={provider} onValueChange={handleProviderChange}>
              <SelectTrigger className="h-8 text-sm">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {providers.map((p) => (
                  <SelectItem key={p.name} value={p.name} className="text-sm">{p.name}</SelectItem>
                ))}
                <SelectItem value={CUSTOM_PROVIDER} className="text-sm">{t("settings.models.customProvider")}</SelectItem>
              </SelectContent>
            </Select>
          </div>

          {/* Custom provider name */}
          {isCustom && (
            <div className="space-y-1.5">
              <label className="text-sm font-medium">{t("settings.models.customProvider")}</label>
              <Input
                className="h-8 text-sm"
                value={customProvider}
                onChange={(e) => setCustomProvider(e.target.value)}
                placeholder="my-provider"
              />
            </div>
          )}

          {/* Model name */}
          <div className="space-y-1.5">
            <label className="text-sm font-medium">{t("settings.models.modelName")}</label>
            <Input
              className="h-8 text-sm"
              value={modelName}
              onChange={(e) => setModelName(e.target.value)}
              placeholder="model-id"
            />
          </div>

          {/* API Key */}
          <div className="space-y-1.5">
            <label className="text-sm font-medium">{t("settings.models.apiKey")}</label>
            <Input
              className="h-8 text-sm font-mono"
              type="password"
              value={apiKey}
              onChange={(e) => setApiKey(e.target.value)}
              placeholder="sk-..."
            />
            <p className="text-xs text-muted-foreground">{t("settings.models.apiKeyHint")}</p>
          </div>

          {/* Base URL (custom only) */}
          {isCustom && (
            <div className="space-y-1.5">
              <label className="text-sm font-medium">{t("settings.models.baseUrl")}</label>
              <Input
                className="h-8 text-sm"
                value={baseUrl}
                onChange={(e) => setBaseUrl(e.target.value)}
                placeholder="https://api.example.com/v1"
              />
              <p className="text-xs text-muted-foreground">{t("settings.models.baseUrlDesc")}</p>
            </div>
          )}

          {/* Display name */}
          <div className="space-y-1.5">
            <label className="text-sm font-medium">{t("settings.models.displayName")}</label>
            <Input
              className="h-8 text-sm"
              value={displayName}
              onChange={(e) => setDisplayName(e.target.value)}
              placeholder="My Model"
            />
          </div>
        </div>

        <DialogFooter>
          <Button variant="outline" size="sm" onClick={() => onOpenChange(false)}>
            Cancel
          </Button>
          <Button size="sm" onClick={handleSubmit} disabled={!modelName.trim() || (isCustom && !customProvider.trim())}>
            {t("settings.models.addModel")}
          </Button>
        </DialogFooter>
      </DialogContent>
    </Dialog>
  );
}
