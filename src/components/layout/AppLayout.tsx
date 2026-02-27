import { Outlet } from "react-router"
import { TitleBar } from "./TitleBar"
import { Footer } from "./Footer"
import { useLayout } from "@/hooks/useLayout"
import { SettingsView } from "@/components/settings/SettingsView"

export function AppLayout() {
  const { settingsOpen, closeSettings } = useLayout()

  return (
    <div
      className="flex h-screen flex-col bg-background text-foreground"
      onContextMenu={(e) => e.preventDefault()}
    >
      <TitleBar />
      <div className="relative flex-1 overflow-hidden">
        <Outlet />
        {settingsOpen && (
          <div className="absolute inset-0 z-50 bg-background">
            <SettingsView onClose={closeSettings} />
          </div>
        )}
      </div>
      <Footer />
    </div>
  )
}
