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
  const [currentToolCall, setCurrentToolCall] = useState<ToolCall | null>(null)
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
          // Append any pending tool call
          let finalMessage = currentAssistantMessage
          if (currentToolCall && currentToolCall.status !== 'pending') {
            finalMessage = {
              ...finalMessage,
              toolCalls: [...(finalMessage.toolCalls || []), currentToolCall],
            }
          }
          setMessages((prev) => [...prev, { ...finalMessage, isStreaming: false }])
          setCurrentAssistantMessage(null)
          setCurrentToolCall(null)
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
          // Finalize any pending tool call before adding text
          if (currentToolCall) {
            setCurrentAssistantMessage((prev) =>
              prev
                ? {
                    ...prev,
                    content: prev.content + (event.content || ''),
                    toolCalls: [...(prev.toolCalls || []), currentToolCall],
                  }
                : null
            )
            setCurrentToolCall(null)
          } else {
            setCurrentAssistantMessage((prev) =>
              prev
                ? { ...prev, content: prev.content + (event.content || '') }
                : null
            )
          }
        }
        break
      }

      case 'tool_call_start': {
        // Finalize previous tool call if any
        if (currentToolCall && currentAssistantMessage) {
          setCurrentAssistantMessage((prev) =>
            prev
              ? {
                  ...prev,
                  toolCalls: [...(prev.toolCalls || []), currentToolCall],
                }
              : null
          )
        }
        // Start a new tool call
        setCurrentToolCall({
          id: generateId(),
          name: event.tool_name || 'unknown',
          args: event.tool_args || '',
          status: 'running',
        })
        break
      }

      case 'tool_call_args': {
        if (currentToolCall) {
          setCurrentToolCall((prev) =>
            prev ? { ...prev, args: prev.args + (event.content || '') } : null
          )
        }
        break
      }

      case 'tool_result': {
        if (currentToolCall) {
          setCurrentToolCall((prev) =>
            prev
              ? {
                  ...prev,
                  result: event.content || '',
                  status: 'completed',
                }
              : null
          )
        }
        break
      }
    }
  }, [currentAssistantMessage, currentToolCall])

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
    setCurrentToolCall(null)
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
