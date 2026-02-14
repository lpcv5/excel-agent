import { createContext, useContext, useState, useCallback, type ReactNode } from "react"

interface LayoutState {
  leftSidebarOpen: boolean
  rightSidebarOpen: boolean
  settingsOpen: boolean
  toggleLeftSidebar: () => void
  toggleRightSidebar: () => void
  setLeftSidebarOpen: (open: boolean) => void
  setRightSidebarOpen: (open: boolean) => void
  openSettings: () => void
  closeSettings: () => void
}

const LayoutContext = createContext<LayoutState | undefined>(undefined)

export function LayoutProvider({ children }: { children: ReactNode }) {
  const [leftSidebarOpen, setLeftSidebarOpen] = useState(true)
  const [rightSidebarOpen, setRightSidebarOpen] = useState(true)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const toggleLeftSidebar = useCallback(() => setLeftSidebarOpen((p) => !p), [])
  const toggleRightSidebar = useCallback(() => setRightSidebarOpen((p) => !p), [])
  const openSettings = useCallback(() => setSettingsOpen(true), [])
  const closeSettings = useCallback(() => setSettingsOpen(false), [])

  return (
    <LayoutContext.Provider value={{
      leftSidebarOpen, rightSidebarOpen, settingsOpen,
      toggleLeftSidebar, toggleRightSidebar,
      setLeftSidebarOpen, setRightSidebarOpen,
      openSettings, closeSettings,
    }}>
      {children}
    </LayoutContext.Provider>
  )
}

// eslint-disable-next-line react-refresh/only-export-components
export function useLayout() {
  const ctx = useContext(LayoutContext)
  if (!ctx) throw new Error("useLayout must be used within LayoutProvider")
  return ctx
}
