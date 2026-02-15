import { useEffect, useState, useCallback } from 'react'
import type { AgentEvent } from '@/types/events'

/**
 * pywebview API type definition
 */
interface UnsavedWorkbook {
  name: string
  path: string | null
  saved: boolean
}

interface PywebviewAPI {
  sendQuery: (query: string) => Promise<{ success: boolean; error?: string }>
  newSession: () => Promise<{ success: boolean; thread_id?: string; error?: string }>
  getExcelStatus: () => Promise<{
    success: boolean
    excel_running?: boolean
    workbook_count?: number
    open_workbooks?: string[]
    error?: string
  }>
  getConfig: () => Promise<{
    success: boolean
    model?: string
    working_dir?: string
    thread_id?: string
    streaming_enabled?: boolean
    error?: string
  }>
  isStreaming: () => Promise<{ success: boolean; is_streaming: boolean }>
  stopStreaming: () => Promise<{ success: boolean; message?: string }>
  // Excel cleanup methods
  getUnsavedWorkbooks: () => Promise<{
    success: boolean
    unsaved_workbooks: UnsavedWorkbook[]
    error?: string
  }>
  saveAllWorkbooks: () => Promise<{
    success: boolean
    saved?: number
    errors?: string[]
    error?: string
  }>
  closeAllWorkbooks: (save: boolean) => Promise<{
    success: boolean
    closed?: number
    errors?: string[]
    error?: string
  }>
  quitExcel: () => Promise<{ success: boolean; error?: string }>
  prepareClose: () => Promise<{
    success: boolean
    has_unsaved: boolean
    unsaved_workbooks: UnsavedWorkbook[]
    can_close: boolean
    error?: string
  }>
  forceClose: (save: boolean) => Promise<{ success: boolean; error?: string }>
}

declare global {
  interface Window {
    pywebview?: {
      api: PywebviewAPI
    }
    __onAgentEvent?: (event: AgentEvent) => void
  }
}

/**
 * Hook to access pywebview API
 */
export function usePywebview() {
  const [isReady, setIsReady] = useState(false)
  const [api, setApi] = useState<PywebviewAPI | null>(null)

  useEffect(() => {
    const checkReady = () => {
      if (window.pywebview?.api) {
        setApi(window.pywebview.api)
        setIsReady(true)
      }
    }

    // Check immediately
    checkReady()

    // Also listen for pywebview ready event
    window.addEventListener('pywebviewready', checkReady)

    return () => {
      window.removeEventListener('pywebviewready', checkReady)
    }
  }, [])

  return { api, isReady }
}

/**
 * Event listener registry for supporting multiple listeners
 */
const eventListeners: Set<(event: AgentEvent) => void> = new Set()

function dispatchEvent(event: AgentEvent) {
  eventListeners.forEach((listener) => {
    try {
      listener(event)
    } catch (e) {
      console.error('Error in event listener:', e)
    }
  })
}

// Set up the global dispatcher once
if (typeof window !== 'undefined') {
  window.__onAgentEvent = dispatchEvent
}

/**
 * Hook to subscribe to agent events
 * Supports multiple listeners
 */
export function useAgentEvent(callback: (event: AgentEvent) => void) {
  useEffect(() => {
    eventListeners.add(callback)

    return () => {
      eventListeners.delete(callback)
    }
  }, [callback])
}

/**
 * Hook to manage chat state with pywebview backend
 */
export function usePywebviewChat() {
  const { api, isReady } = usePywebview()
  const [isStreaming, setIsStreaming] = useState(false)
  const [error, setError] = useState<string | null>(null)

  const sendQuery = useCallback(
    async (query: string): Promise<boolean> => {
      if (!api) {
        setError('API not ready')
        return false
      }

      setError(null)
      setIsStreaming(true)

      try {
        const result = await api.sendQuery(query)
        if (!result.success) {
          setError(result.error || 'Failed to send query')
          setIsStreaming(false)
          return false
        }
        return true
      } catch (e) {
        setError(e instanceof Error ? e.message : 'Unknown error')
        setIsStreaming(false)
        return false
      }
    },
    [api]
  )

  const newSession = useCallback(async (): Promise<string | null> => {
    if (!api) {
      setError('API not ready')
      return null
    }

    setError(null)

    try {
      const result = await api.newSession()
      if (!result.success) {
        setError(result.error || 'Failed to create new session')
        return null
      }
      return result.thread_id || null
    } catch (e) {
      setError(e instanceof Error ? e.message : 'Unknown error')
      return null
    }
  }, [api])

  const getExcelStatus = useCallback(async () => {
    if (!api) {
      return null
    }

    try {
      return await api.getExcelStatus()
    } catch {
      return null
    }
  }, [api])

  const stopStreaming = useCallback(async () => {
    if (!api) return

    try {
      await api.stopStreaming()
      setIsStreaming(false)
    } catch {
      // Ignore errors
    }
  }, [api])

  // Listen for query_end to update streaming state
  useAgentEvent((event) => {
    if (event.type === 'query_end' || event.type === 'error') {
      setIsStreaming(false)
    }
  })

  return {
    api,
    isReady,
    isStreaming,
    error,
    sendQuery,
    newSession,
    getExcelStatus,
    stopStreaming,
    clearError: () => setError(null),
  }
}
