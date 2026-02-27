import { Component, type ReactNode } from "react"
import { RouterProvider } from "react-router"
import { ThemeProvider } from "@/components/theme/ThemeProvider"
import { LayoutProvider } from "@/hooks/useLayout"
import { TooltipProvider } from "@/components/ui/tooltip"
import { router } from "@/router"

class ErrorBoundary extends Component<{ children: ReactNode }, { error: string | null }> {
  state = { error: null }
  static getDerivedStateFromError(e: Error) { return { error: e.message + "\n" + e.stack } }
  render() {
    if (this.state.error) return <pre style={{ padding: 20, color: "red", whiteSpace: "pre-wrap", fontSize: 12 }}>{this.state.error}</pre>
    return this.props.children
  }
}

function App() {
  return (
    <ErrorBoundary>
      <ThemeProvider defaultTheme="light">
        <LayoutProvider>
          <TooltipProvider>
            <RouterProvider router={router} />
          </TooltipProvider>
        </LayoutProvider>
      </ThemeProvider>
    </ErrorBoundary>
  )
}

export default App
