import { useState, useCallback, useRef, useEffect } from 'react'
import type { AgentEvent } from '@/types/events'
import { usePywebviewChat, useAgentEvent } from './usePywebview'

export interface Message {
  id: string
  role: 'user' | 'assistant'
  content: string
  timestamp: Date
  isStreaming?: boolean
  thinking?: string
  toolCalls?: ToolCall[]
}

export interface ToolCall {
  id: string
  name: string
  args: string
  result?: string
  status: 'pending' | 'running' | 'completed' | 'error'
}

/**
 * Generate a unique ID
 */
function generateId(): string {
  return `${Date.now()}-${Math.random().toString(36).substr(2, 9)}`
}

/**
 * Hook to manage chat state
 */
export function useChat() {
  const [messages, setMessages] = useState<Message[]>([])
  const [currentAssistantMessage, setCurrentAssistantMessage] = useState<Message | null>(null)
  const messageEndRef = useRef<HTMLDivElement>(null)

  const { isReady, isStreaming, error, sendQuery, newSession, clearError } = usePywebviewChat()

  // Auto-scroll to bottom
  useEffect(() => {
    messageEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages, currentAssistantMessage])

  // Handle agent events
  useAgentEvent((event: AgentEvent) => {
    handleEvent(event)
  })

  const upsertToolCall = useCallback((toolCalls: ToolCall[], toolCall: ToolCall) => {
    const index = toolCalls.findIndex(tc => tc.id === toolCall.id)
    if (index >= 0) {
      const updated = [...toolCalls]
      updated[index] = { ...updated[index], ...toolCall }
      return updated
    }
    return [...toolCalls, toolCall]
  }, [])

  const updateToolCallById = useCallback(
    (toolCalls: ToolCall[], id: string, updater: (toolCall: ToolCall) => ToolCall) => {
      const index = toolCalls.findIndex(tc => tc.id === id)
      if (index >= 0) {
        const updated = [...toolCalls]
        updated[index] = updater(updated[index])
        return updated
      }
      return toolCalls
    },
    []
  )

  const updateLastRunningByName = useCallback(
    (toolCalls: ToolCall[], name: string, updater: (toolCall: ToolCall) => ToolCall) => {
      for (let i = toolCalls.length - 1; i >= 0; i -= 1) {
        if (toolCalls[i].name === name && toolCalls[i].status === 'running') {
          const updated = [...toolCalls]
          updated[i] = updater(updated[i])
          return updated
        }
      }
      return toolCalls
    },
    []
  )

  const handleEvent = useCallback((event: AgentEvent) => {
    switch (event.type) {
      case 'query_start': {
        // User message is added when sending
        // Start a new assistant message
        setCurrentAssistantMessage({
          id: generateId(),
          role: 'assistant',
          content: '',
          timestamp: new Date(),
          isStreaming: true,
          thinking: '',
          toolCalls: [],
        })
        break
      }

      case 'query_end': {
        // Finalize the current assistant message
        if (currentAssistantMessage) {
          setMessages((prev) => [...prev, { ...currentAssistantMessage, isStreaming: false }])
          setCurrentAssistantMessage(null)
        }
        break
      }

      case 'error': {
        if (currentAssistantMessage) {
          setMessages((prev) => [
            ...prev,
            {
              ...currentAssistantMessage,
              content: event.error_message || 'An error occurred',
              isStreaming: false,
            },
          ])
          setCurrentAssistantMessage(null)
        }
        break
      }

      case 'thinking': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) =>
            prev
              ? { ...prev, thinking: (prev.thinking || '') + (event.content || '') }
              : null
          )
        }
        break
      }

      case 'text': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) =>
            prev
              ? { ...prev, content: prev.content + (event.content || '') }
              : null
          )
        }
        break
      }

      case 'tool_call_start': {
        if (currentAssistantMessage) {
          const id = event.tool_call_id || generateId()
          const toolCall: ToolCall = {
            id,
            name: event.tool_name || 'unknown',
            args: event.tool_args || '',
            status: 'running',
          }
          setCurrentAssistantMessage((prev) =>
            prev
              ? {
                  ...prev,
                  toolCalls: upsertToolCall(prev.toolCalls || [], toolCall),
                }
              : null
          )
        }
        break
      }

      case 'tool_call_args': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) =>
            prev
              ? {
                  ...prev,
                  toolCalls: event.tool_call_id
                    ? updateToolCallById(
                        prev.toolCalls || [],
                        event.tool_call_id,
                        (tc) => ({ ...tc, args: tc.args + (event.content || '') })
                      )
                    : updateLastRunningByName(
                        prev.toolCalls || [],
                        event.tool_name || 'unknown',
                        (tc) => ({ ...tc, args: tc.args + (event.content || '') })
                      ),
                }
              : null
          )
        }
        break
      }

      case 'tool_result': {
        if (currentAssistantMessage) {
          const status = ((event.data as { status?: string } | null)?.status === 'error')
            ? 'error'
            : 'completed'
          setCurrentAssistantMessage((prev) =>
            prev
              ? {
                  ...prev,
                  toolCalls: event.tool_call_id
                    ? updateToolCallById(
                        prev.toolCalls || [],
                        event.tool_call_id,
                        (tc) => ({
                          ...tc,
                          result: event.content || '',
                          status,
                        })
                      )
                    : updateLastRunningByName(
                        prev.toolCalls || [],
                        event.tool_name || 'unknown',
                        (tc) => ({
                          ...tc,
                          result: event.content || '',
                          status,
                        })
                      ),
                }
              : null
          )
        }
        break
      }
    }
  }, [currentAssistantMessage, updateLastRunningByName, updateToolCallById, upsertToolCall])

  const addUserMessage = useCallback((content: string) => {
    const message: Message = {
      id: generateId(),
      role: 'user',
      content,
      timestamp: new Date(),
    }
    setMessages((prev) => [...prev, message])
  }, [])

  const handleSend = useCallback(
    async (query: string) => {
      if (!query.trim() || isStreaming) return

      addUserMessage(query)
      await sendQuery(query)
    },
    [addUserMessage, sendQuery, isStreaming]
  )

  const handleNewSession = useCallback(async () => {
    await newSession()
    setMessages([])
    setCurrentAssistantMessage(null)
  }, [newSession])

  // Get all messages including the current streaming one
  const allMessages = currentAssistantMessage
    ? [...messages, currentAssistantMessage]
    : messages

  return {
    messages: allMessages,
    isReady,
    isStreaming,
    error,
    messageEndRef,
    handleSend,
    handleNewSession,
    clearError,
  }
}
