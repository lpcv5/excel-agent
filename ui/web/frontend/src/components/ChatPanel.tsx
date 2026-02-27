import { useEffect, useState, useCallback } from 'react'
import { RefreshCw, AlertCircle, AlertTriangle, Save, X } from 'lucide-react'
import { Button } from '@/components/ui/button'
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogFooter,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'
import { Alert, AlertDescription } from '@/components/ui/alert'
import { MessageList } from './MessageList'
import { InputBar } from './InputBar'
import { TodoPanel } from './TodoPanel'
import { useChat } from '@/hooks/useChat'

interface UnsavedWorkbook {
  name: string
  path: string | null
  saved: boolean
}

export function ChatPanel() {
  const {
    messages,
    isReady,
    isStreaming,
    error,
    messageEndRef,
    scrollViewportRef,
    handleSend,
    handleNewSession,
    clearError,
    todos,
  } = useChat()

  const [showCloseDialog, setShowCloseDialog] = useState(false)
  const [unsavedWorkbooks, setUnsavedWorkbooks] = useState<UnsavedWorkbook[]>([])
  const [isClosing, setIsClosing] = useState(false)
  const [quickQuery, setQuickQuery] = useState<string | undefined>()

  const handleExampleClick = useCallback((text: string) => {
    setQuickQuery(text)
  }, [])

  // Handle window close event
  useEffect(() => {
    const handleBeforeUnload = async (e: BeforeUnloadEvent) => {
      if (isStreaming) {
        e.preventDefault()
        e.returnValue = 'A query is in progress. Are you sure you want to close?'
        return e.returnValue
      }
    }

    window.addEventListener('beforeunload', handleBeforeUnload)
    return () => window.removeEventListener('beforeunload', handleBeforeUnload)
  }, [isStreaming])

  // Check for unsaved workbooks and show dialog
  const handleCloseClick = async () => {
    if (!window.pywebview?.api) {
      window.close()
      return
    }

    try {
      const result = await window.pywebview.api.prepareClose()
      if (result.has_unsaved) {
        setUnsavedWorkbooks(result.unsaved_workbooks)
        setShowCloseDialog(true)
      } else {
        // No unsaved workbooks, close directly
        await forceClose(false)
      }
    } catch {
      window.close()
    }
  }

  const forceClose = async (save: boolean) => {
    setIsClosing(true)
    try {
      if (window.pywebview?.api) {
        await window.pywebview.api.forceClose(save)
      }
    } catch {
      // Ignore errors
    }
    // Close the window
    window.close()
  }

  const handleSaveAndClose = async () => {
    await forceClose(true)
  }

  const handleDiscardAndClose = async () => {
    await forceClose(false)
  }

  const handleCancelClose = () => {
    setShowCloseDialog(false)
    setUnsavedWorkbooks([])
  }

  return (
    <div className="flex h-full flex-col">
      {/* Header */}
      <div className="flex items-center justify-between border-b px-4 py-3">
        <h1 className="text-lg font-semibold">Excel Agent</h1>
        <div className="flex items-center gap-2">
          <Button
            variant="outline"
            size="sm"
            onClick={handleNewSession}
            disabled={isStreaming}
          >
            <RefreshCw className="mr-2 h-4 w-4" />
            New Session
          </Button>
          <Button
            variant="ghost"
            size="sm"
            onClick={handleCloseClick}
            disabled={isClosing}
          >
            <X className="h-4 w-4" />
          </Button>
        </div>
      </div>

      {/* Error Banner */}
      {error && (
        <Alert variant="destructive" className="rounded-none border-x-0 border-t-0">
          <AlertCircle className="h-4 w-4" />
          <AlertDescription className="flex items-center gap-2">
            <span className="flex-1">{error}</span>
            <Button variant="ghost" size="sm" onClick={clearError}>
              Dismiss
            </Button>
          </AlertDescription>
        </Alert>
      )}

      {/* Not Ready Banner */}
      {!isReady && (
        <div className="flex items-center justify-center gap-2 border-b bg-muted px-4 py-3 text-sm text-muted-foreground">
          <div className="h-4 w-4 animate-spin rounded-full border-2 border-current border-t-transparent" />
          Connecting to Python backend...
        </div>
      )}

      {/* Close Confirmation Dialog */}
      <Dialog open={showCloseDialog} onOpenChange={(open) => !open && handleCancelClose()}>
        <DialogContent className="sm:max-w-md">
          <DialogHeader>
            <div className="flex items-center gap-3">
              <div className="flex h-10 w-10 items-center justify-center rounded-full bg-warning/15">
                <AlertTriangle className="h-5 w-5 text-warning" />
              </div>
              <div>
                <DialogTitle>Unsaved Workbooks</DialogTitle>
                <DialogDescription>
                  The following workbooks have unsaved changes:
                </DialogDescription>
              </div>
            </div>
          </DialogHeader>

          <ul className="my-2 max-h-40 space-y-1 overflow-auto rounded border bg-muted/50 p-2">
            {unsavedWorkbooks.map((wb, i) => (
              <li key={i} className="text-sm">
                <span className="font-medium">{wb.name}</span>
                {wb.path && (
                  <span className="ml-2 text-muted-foreground text-xs">
                    {wb.path}
                  </span>
                )}
              </li>
            ))}
          </ul>

          <p className="text-sm text-muted-foreground">
            Do you want to save changes before closing?
          </p>

          <DialogFooter>
            <Button variant="outline" onClick={handleCancelClose} disabled={isClosing}>
              Cancel
            </Button>
            <Button
              variant="destructive"
              onClick={handleDiscardAndClose}
              disabled={isClosing}
            >
              Don't Save
            </Button>
            <Button onClick={handleSaveAndClose} disabled={isClosing}>
              <Save className="mr-2 h-4 w-4" />
              Save All
            </Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Messages */}
      <div className="min-w-0 flex-1 overflow-hidden">
        <MessageList
          messages={messages}
          messageEndRef={messageEndRef}
          viewportRef={scrollViewportRef}
          onExampleClick={handleExampleClick}
        />
      </div>

      {/* Todo Panel - above InputBar */}
      <TodoPanel todos={todos} />

      {/* Input */}
      <InputBar
        onSend={handleSend}
        isStreaming={isStreaming}
        disabled={!isReady}
        quickQuery={quickQuery}
      />
    </div>
  )
}
