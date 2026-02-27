import { useState } from 'react'
import { User, Bot, MoreVertical, Copy, Check, Lightbulb } from 'lucide-react'
import { cn } from '@/lib/utils'
import type { Message, ContentBlock } from '@/hooks/useChat'
import { ToolCallGroup } from './ToolCallDisplay'
import { Markdown } from './Markdown'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'

interface MessageItemProps {
  message: Message
}

export function MessageItem({ message }: MessageItemProps) {
  const isUser = message.role === 'user'
  const [showMenu, setShowMenu] = useState(false)
  const [copied, setCopied] = useState(false)

  const handleCopy = async () => {
    const text = getMessageText(message)
    try {
      await navigator.clipboard.writeText(text)
      setCopied(true)
      setTimeout(() => setCopied(false), 2000)
    } catch (e) {
      console.error('Failed to copy:', e)
    }
    setShowMenu(false)
  }

  // Check if message uses new blocks structure
  const hasBlocks = message.blocks && message.blocks.length > 0

  return (
    <div
      className={cn(
        'group relative flex gap-3',
        isUser ? 'flex-row-reverse' : 'flex-row'
      )}
    >
      {/* Avatar */}
      <Avatar className={cn(isUser && 'bg-primary')}>
        <AvatarFallback
          className={cn(
            isUser ? 'bg-primary text-primary-foreground' : 'bg-gradient-to-br from-primary to-primary/70 text-primary-foreground'
          )}
        >
          {isUser ? <User className="h-5 w-5" /> : <Bot className="h-5 w-5" />}
        </AvatarFallback>
      </Avatar>

      {/* Content Card */}
      <div
        className={cn(
          'relative flex min-w-0 max-w-[85%] flex-col overflow-hidden',
          isUser
            ? 'items-end rounded-2xl rounded-tr-sm bg-primary text-primary-foreground'
            : 'items-start rounded-2xl rounded-tl-sm border bg-card'
        )}
      >
        {/* Menu Button - Top Right Corner */}
        {!message.isStreaming && (
          <div className={cn('absolute -top-1 z-10', isUser ? '-left-1' : '-right-1')}>
            <DropdownMenu open={showMenu} onOpenChange={setShowMenu}>
              <DropdownMenuTrigger asChild>
                <button
                  className={cn(
                    'flex h-6 w-6 items-center justify-center rounded-full shadow-sm transition-all',
                    isUser
                      ? 'bg-primary-foreground/20 text-primary-foreground hover:bg-primary-foreground/30'
                      : 'border bg-background hover:bg-accent',
                    showMenu ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'
                  )}
                >
                  <MoreVertical className="h-3.5 w-3.5" />
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align={isUser ? 'start' : 'end'} className="min-w-[140px]">
                <DropdownMenuItem onClick={handleCopy}>
                  {copied ? (
                    <Check className="mr-2 h-4 w-4 text-success" />
                  ) : (
                    <Copy className="mr-2 h-4 w-4" />
                  )}
                  <span>{copied ? 'Copied!' : 'Copy'}</span>
                </DropdownMenuItem>
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        )}

        {/* Card Content */}
        <div className="flex min-w-0 max-w-full flex-col gap-2 p-3">
          {isUser ? (
            // User message - simple text
            <div className="whitespace-pre-wrap break-words text-sm">
              {message.content}
            </div>
          ) : hasBlocks ? (
            // Assistant message with chronological blocks
            <ContentBlocksRenderer
              blocks={message.blocks!}
              isStreaming={message.isStreaming}
            />
          ) : (
            // Fallback for old format or empty
            message.content ? (
              <Markdown content={message.content} />
            ) : message.isStreaming ? (
              <div className="flex items-center gap-2 text-muted-foreground">
                <div className="h-2 w-2 animate-pulse rounded-full bg-current" />
                <span className="text-sm">Thinking...</span>
              </div>
            ) : null
          )}
        </div>

        {/* Timestamp */}
        <div
          className={cn(
            'px-3 pb-1.5 pt-0',
            isUser ? 'text-primary-foreground/60' : 'text-muted-foreground'
          )}
        >
          <span className="text-xs">
            {formatTime(message.timestamp)}
          </span>
        </div>
      </div>
    </div>
  )
}

/**
 * Render content blocks in chronological order
 */
function ContentBlocksRenderer({
  blocks,
  isStreaming,
}: {
  blocks: ContentBlock[]
  isStreaming?: boolean
}) {
  // Group consecutive tool calls together
  const groupedBlocks: Array<ContentBlock | ContentBlock[]> = []
  let currentToolGroup: ContentBlock[] = []

  for (const block of blocks) {
    if (block.type === 'tool_call') {
      currentToolGroup.push(block)
    } else {
      // Flush tool group if any
      if (currentToolGroup.length > 0) {
        groupedBlocks.push([...currentToolGroup])
        currentToolGroup = []
      }
      groupedBlocks.push(block)
    }
  }
  // Don't forget the last group
  if (currentToolGroup.length > 0) {
    groupedBlocks.push([...currentToolGroup])
  }

  return (
    <div className="flex flex-col gap-3">
      {groupedBlocks.map((item, index) => {
        if (Array.isArray(item)) {
          // Tool call group
          const toolCalls = item
            .filter(b => b.type === 'tool_call')
            .map(b => b.toolCall!)
            .filter(Boolean)

          return (
            <ToolCallGroup
              key={`tools-${index}`}
              toolCalls={toolCalls}
              isStreaming={isStreaming}
            />
          )
        }

        // Single content block
        const block = item
        const key = block.id || `block-${index}`

        switch (block.type) {
          case 'thinking':
            return (
              <ThinkingBlock key={key} content={block.content || ''} />
            )

          case 'text':
            return block.content ? (
              <div key={key} className="prose prose-sm dark:prose-invert max-w-none">
                <Markdown content={block.content} />
              </div>
            ) : null

          default:
            return null
        }
      })}

      {/* Streaming indicator at the end */}
      {isStreaming && blocks.length > 0 && (
        <div className="flex items-center gap-1 text-muted-foreground">
          <div className="h-1.5 w-1.5 animate-pulse rounded-full bg-current" />
        </div>
      )}
    </div>
  )
}

/**
 * Thinking block with collapsible content
 */
function ThinkingBlock({ content }: { content: string }) {
  const [expanded, setExpanded] = useState(false)

  if (!content) return null

  return (
    <div className="rounded-lg border border-warning/40 bg-warning/10 text-sm">
      <button
        type="button"
        onClick={() => setExpanded(!expanded)}
        className="flex w-full items-center gap-2 p-2 text-left text-warning hover:bg-warning/20 dark:text-warning-foreground dark:hover:bg-warning/20"
      >
        <Lightbulb className="h-4 w-4 shrink-0" />
        <span className="font-medium">Thinking</span>
        <span className="ml-auto text-xs text-warning/70 dark:text-warning-foreground/70">
          {expanded ? 'Click to hide' : 'Click to expand'}
        </span>
      </button>
      {expanded && (
        <div className="border-t border-warning/40 px-3 py-2 text-warning dark:text-warning-foreground">
          <pre className="whitespace-pre-wrap break-words font-mono text-xs">
            {content}
          </pre>
        </div>
      )}
    </div>
  )
}

/**
 * Get text content from message for copying
 */
function getMessageText(message: Message): string {
  if (message.role === 'user') {
    return message.content || ''
  }

  // For assistant messages with blocks
  if (message.blocks && message.blocks.length > 0) {
    const parts: string[] = []

    for (const block of message.blocks) {
      if (block.type === 'text' && block.content) {
        parts.push(block.content)
      } else if (block.type === 'thinking' && block.content) {
        parts.push(`[Thinking]\n${block.content}`)
      } else if (block.type === 'tool_call' && block.toolCall) {
        parts.push(
          `[Tool: ${block.toolCall.name}]\nArgs: ${block.toolCall.args}\nResult: ${block.toolCall.result || 'N/A'}`
        )
      }
    }

    return parts.join('\n\n')
  }

  return message.content || ''
}

function formatTime(date: Date): string {
  return date.toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
  })
}
