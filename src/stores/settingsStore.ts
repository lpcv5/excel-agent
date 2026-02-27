import { create } from "zustand";
import { settings as settingsApi, type AppSettings, type ModelEntry } from "@/services/api";

interface SettingsStore {
  settings: AppSettings | null;
  loading: boolean;
  saving: boolean;
  fetchSettings: () => Promise<void>;
  saveSettings: (s?: AppSettings) => Promise<void>;
  addModel: (entry: Omit<ModelEntry, "id">) => void;
  removeModel: (id: string) => void;
  updateModel: (id: string, patch: Partial<Omit<ModelEntry, "id">>) => void;
  setMainModelId: (id: string | null) => void;
  setSubagentsModelId: (id: string | null) => void;
  setAnalysisModelId: (id: string | null) => void;
}

export const useSettingsStore = create<SettingsStore>((set, get) => ({
  settings: null,
  loading: false,
  saving: false,

  fetchSettings: async () => {
    set({ loading: true });
    try {
      const s = await settingsApi.get();
      set({ settings: s });
    } finally {
      set({ loading: false });
    }
  },

  saveSettings: async (s?: AppSettings) => {
    const current = s ?? get().settings;
    if (!current) return;
    set({ saving: true });
    try {
      await settingsApi.save(current);
      set({ settings: current });
    } finally {
      set({ saving: false });
    }
  },

  addModel: (entry) => {
    const current = get().settings;
    if (!current) return;
    const newEntry: ModelEntry = { ...entry, id: crypto.randomUUID() };
    const updated: AppSettings = {
      ...current,
      models: [...current.models, newEntry],
      main_model_id: current.main_model_id ?? newEntry.id,
    };
    set({ settings: updated });
  },

  removeModel: (id) => {
    const current = get().settings;
    if (!current) return;
    const updated: AppSettings = {
      ...current,
      models: current.models.filter((m) => m.id !== id),
      main_model_id: current.main_model_id === id ? null : current.main_model_id,
      subagents_model_id: current.subagents_model_id === id ? null : current.subagents_model_id,
      analysis_model_id: current.analysis_model_id === id ? null : current.analysis_model_id,
    };
    set({ settings: updated });
  },

  updateModel: (id, patch) => {
    const current = get().settings;
    if (!current) return;
    set({
      settings: {
        ...current,
        models: current.models.map((m) => (m.id === id ? { ...m, ...patch } : m)),
      },
    });
  },

  setMainModelId: (id) => {
    const current = get().settings;
    if (!current) return;
    set({ settings: { ...current, main_model_id: id } });
  },

  setSubagentsModelId: (id) => {
    const current = get().settings;
    if (!current) return;
    set({ settings: { ...current, subagents_model_id: id } });
  },

  setAnalysisModelId: (id) => {
    const current = get().settings;
    if (!current) return;
    set({ settings: { ...current, analysis_model_id: id } });
  },
}));
