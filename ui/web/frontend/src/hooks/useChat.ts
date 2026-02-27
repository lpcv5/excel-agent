import { useState, useCallback, useRef, useEffect } from 'react'
import type { AgentEvent, TodoItem } from '@/types/events'
import { usePywebviewChat, useAgentEvent } from './usePywebview'

// Message content block types - supports chronological ordering
export type ContentBlockType = 'text' | 'thinking' | 'tool_call'

export interface ContentBlock {
  id: string
  type: ContentBlockType
  timestamp: number
  // For text/thinking
  content?: string
  // For tool_call
  toolCall?: ToolCall
}

export interface Message {
  id: string
  role: 'user' | 'assistant'
  content: string // Keep for user messages
  timestamp: Date
  isStreaming?: boolean
  // For assistant messages - chronological content blocks
  blocks?: ContentBlock[]
  // Track running tool calls for updates
  toolCallMap?: Map<string, ToolCall>
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
  const [todos, setTodos] = useState<TodoItem[]>([])
  const messageEndRef = useRef<HTMLDivElement>(null)
  const scrollViewportRef = useRef<HTMLDivElement | null>(null)
  const [isUserAtBottom, setIsUserAtBottom] = useState(true)
  const scrollHandlerRef = useRef<(() => void) | null>(null)

  const { isReady, isStreaming, error, sendQuery, newSession, clearError } = usePywebviewChat()

  // Set up scroll tracking using callback ref
  const setScrollViewportRef = useCallback((node: HTMLDivElement | null) => {
    // Clean up previous listener
    if (scrollHandlerRef.current && scrollViewportRef.current) {
      scrollViewportRef.current.removeEventListener('scroll', scrollHandlerRef.current)
    }

    scrollViewportRef.current = node

    if (node) {
      const handleScroll = () => {
        const distanceFromBottom = node.scrollHeight - node.scrollTop - node.clientHeight
        setIsUserAtBottom(distanceFromBottom < 80)
      }

      scrollHandlerRef.current = handleScroll
      handleScroll() // Initial check
      node.addEventListener('scroll', handleScroll)
    }
  }, [])

  // Auto-scroll to bottom only if user is already near bottom
  useEffect(() => {
    if (isUserAtBottom) {
      messageEndRef.current?.scrollIntoView({ behavior: 'smooth' })
    }
  }, [messages, currentAssistantMessage, isUserAtBottom])

  // Handle agent events
  useAgentEvent((event: AgentEvent) => {
    handleEvent(event)
  })

  const handleEvent = useCallback((event: AgentEvent) => {
    switch (event.type) {
      case 'query_start': {
        // Start a new assistant message with chronological blocks
        setCurrentAssistantMessage({
          id: generateId(),
          role: 'assistant',
          content: '',
          timestamp: new Date(),
          isStreaming: true,
          blocks: [],
          toolCallMap: new Map(),
        })
        break
      }

      case 'query_end': {
        if (currentAssistantMessage) {
          // Finalize - convert toolCallMap to final form
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
          setCurrentAssistantMessage((prev) => {
            if (!prev) return null
            const blocks = prev.blocks || []
            // Append to last thinking block or create new one
            const lastBlock = blocks[blocks.length - 1]
            if (lastBlock?.type === 'thinking') {
              const updatedBlocks = [...blocks]
              updatedBlocks[updatedBlocks.length - 1] = {
                ...lastBlock,
                content: (lastBlock.content || '') + (event.content || ''),
              }
              return { ...prev, blocks: updatedBlocks }
            }
            return {
              ...prev,
              blocks: [...blocks, {
                id: generateId(),
                type: 'thinking',
                timestamp: Date.now(),
                content: event.content || '',
              }],
            }
          })
        }
        break
      }

      case 'text': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) => {
            if (!prev) return null
            const blocks = prev.blocks || []
            // Append to last text block or create new one
            const lastBlock = blocks[blocks.length - 1]
            if (lastBlock?.type === 'text') {
              const updatedBlocks = [...blocks]
              updatedBlocks[updatedBlocks.length - 1] = {
                ...lastBlock,
                content: (lastBlock.content || '') + (event.content || ''),
              }
              return { ...prev, blocks: updatedBlocks }
            }
            return {
              ...prev,
              blocks: [...blocks, {
                id: generateId(),
                type: 'text',
                timestamp: Date.now(),
                content: event.content || '',
              }],
            }
          })
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
          setCurrentAssistantMessage((prev) => {
            if (!prev) return null
            const toolCallMap = new Map(prev.toolCallMap || [])
            toolCallMap.set(id, toolCall)
            return {
              ...prev,
              toolCallMap,
              blocks: [...(prev.blocks || []), {
                id: `tool-${id}`,
                type: 'tool_call',
                timestamp: Date.now(),
                toolCall,
              }],
            }
          })
        }
        break
      }

      case 'tool_call_args': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) => {
            if (!prev) return null
            const toolCallMap = new Map(prev.toolCallMap || [])
            const toolId = event.tool_call_id

            // Find and update the tool call
            for (const [id, tc] of toolCallMap) {
              if (id === toolId || (tc.name === event.tool_name && tc.status === 'running')) {
                toolCallMap.set(id, {
                  ...tc,
                  args: tc.args + (event.content || ''),
                })
                break
              }
            }

            // Update the block as well
            const blocks = (prev.blocks || []).map(block => {
              if (block.type === 'tool_call' && block.toolCall) {
                const tc = block.toolCall
                if (tc.id === toolId || (tc.name === event.tool_name && tc.status === 'running')) {
                  return {
                    ...block,
                    toolCall: {
                      ...tc,
                      args: tc.args + (event.content || ''),
                    },
                  }
                }
              }
              return block
            })

            return { ...prev, toolCallMap, blocks }
          })
        }
        break
      }

      case 'tool_result': {
        if (currentAssistantMessage) {
          setCurrentAssistantMessage((prev) => {
            if (!prev) return null
            const status: 'error' | 'completed' = ((event.data as { status?: string } | null)?.status === 'error')
              ? 'error'
              : 'completed'

            const toolCallMap = new Map(prev.toolCallMap || [])
            const toolId = event.tool_call_id

            // Update in map
            for (const [id, tc] of toolCallMap) {
              if (id === toolId || (tc.name === event.tool_name && tc.status === 'running')) {
                toolCallMap.set(id, {
                  ...tc,
                  result: event.content || '',
                  status,
                })
                break
              }
            }

            // Update the block
            const blocks = (prev.blocks || []).map(block => {
              if (block.type === 'tool_call' && block.toolCall) {
                const tc = block.toolCall
                if (tc.id === toolId || (tc.name === event.tool_name && tc.status === 'running')) {
                  return {
                    ...block,
                    toolCall: {
                      ...tc,
                      result: event.content || '',
                      status,
                    },
                  }
                }
              }
              return block
            })

            return { ...prev, toolCallMap, blocks }
          })
        }
        break
      }

      case 'todo_update': {
        if (event.todos) {
          setTodos(event.todos)
        }
        break
      }
    }
  }, [currentAssistantMessage])

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
    setTodos([])
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
    scrollViewportRef: setScrollViewportRef,
    handleSend,
    handleNewSession,
    clearError,
    todos,
  }
}
