import { useState } from 'react'
import { User, Bot, Brain, MoreVertical, Copy, Check } from 'lucide-react'
import { cn } from '@/lib/utils'
import type { Message } from '@/hooks/useChat'
import { ToolCallList } from './ToolCallDisplay'
import { Markdown } from './Markdown'
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuTrigger,
} from '@/components/ui/dropdown-menu'
import { Avatar, AvatarFallback } from '@/components/ui/avatar'
import { Alert, AlertDescription } from '@/components/ui/alert'

interface MessageItemProps {
  message: Message
}

export function MessageItem({ message }: MessageItemProps) {
  const isUser = message.role === 'user'
  const [showMenu, setShowMenu] = useState(false)
  const [copied, setCopied] = useState(false)

  const handleCopy = async () => {
    const text = message.content || ''
    try {
      await navigator.clipboard.writeText(text)
      setCopied(true)
      setTimeout(() => setCopied(false), 2000)
    } catch (e) {
      console.error('Failed to copy:', e)
    }
    setShowMenu(false)
  }

  // Get all content including thinking and tool calls for copying
  const getFullContent = () => {
    let content = message.content || ''
    if (message.thinking) {
      content = `[Thinking]\n${message.thinking}\n\n${content}`
    }
    if (message.toolCalls && message.toolCalls.length > 0) {
      const toolContent = message.toolCalls.map(tc =>
        `[Tool: ${tc.name}]\nArgs: ${tc.args}\nResult: ${tc.result || 'N/A'}`
      ).join('\n\n')
      content = content ? `${content}\n\n${toolContent}` : toolContent
    }
    return content
  }

  const handleCopyFull = async () => {
    try {
      await navigator.clipboard.writeText(getFullContent())
      setCopied(true)
      setTimeout(() => setCopied(false), 2000)
    } catch (e) {
      console.error('Failed to copy:', e)
    }
    setShowMenu(false)
  }

  return (
    <div
      className={cn(
        'group relative flex gap-3',
        isUser ? 'flex-row-reverse' : 'flex-row'
      )}
    >
      {/* Avatar */}
      <Avatar className={cn(isUser && 'bg-primary')}>
        <AvatarFallback className={cn(
          isUser ? 'bg-primary text-primary-foreground' : 'bg-muted border'
        )}>
          {isUser ? <User className="h-5 w-5" /> : <Bot className="h-5 w-5" />}
        </AvatarFallback>
      </Avatar>

      {/* Content Card */}
      <div
        className={cn(
          'relative flex min-w-0 max-w-[85%] flex-col overflow-hidden rounded-xl border bg-card',
          isUser ? 'items-end' : 'items-start'
        )}
      >
        {/* Menu Button - Top Right Corner */}
        {!message.isStreaming && (
          <div className="absolute -right-1 -top-1 z-10">
            <DropdownMenu open={showMenu} onOpenChange={setShowMenu}>
              <DropdownMenuTrigger asChild>
                <button
                  className={cn(
                    'flex h-7 w-7 items-center justify-center rounded-full border bg-background shadow-sm transition-all',
                    'hover:bg-accent hover:shadow-md',
                    showMenu ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'
                  )}
                >
                  <MoreVertical className="h-4 w-4 text-muted-foreground" />
                </button>
              </DropdownMenuTrigger>
              <DropdownMenuContent align="end" className="min-w-[160px]">
                <DropdownMenuItem onClick={handleCopy}>
                  {copied ? (
                    <Check className="mr-2 h-4 w-4 text-green-500" />
                  ) : (
                    <Copy className="mr-2 h-4 w-4" />
                  )}
                  <span>{copied ? 'Copied!' : 'Copy message'}</span>
                </DropdownMenuItem>
                {!isUser && (message.thinking || (message.toolCalls && message.toolCalls.length > 0)) && (
                  <DropdownMenuItem onClick={handleCopyFull}>
                    {copied ? (
                      <Check className="mr-2 h-4 w-4 text-green-500" />
                    ) : (
                      <Copy className="mr-2 h-4 w-4" />
                    )}
                    <span>{copied ? 'Copied!' : 'Copy with tools'}</span>
                  </DropdownMenuItem>
                )}
              </DropdownMenuContent>
            </DropdownMenu>
          </div>
        )}

        {/* Card Content */}
        <div className="flex min-w-0 max-w-full flex-col gap-2 p-3">
          {/* Thinking */}
          {!isUser && message.thinking && (
            <Alert variant="warning" className="p-3">
              <Brain className="h-4 w-4" />
              <AlertDescription className="whitespace-pre-wrap">
                {message.thinking}
              </AlertDescription>
            </Alert>
          )}

          {/* Tool Calls */}
          {!isUser && message.toolCalls && message.toolCalls.length > 0 && (
            <ToolCallList toolCalls={message.toolCalls} isStreaming={message.isStreaming} />
          )}

          {/* Message Content */}
          <div className="max-w-full px-1">
            {message.content ? (
              isUser ? (
                <div className="whitespace-pre-wrap break-words text-sm">
                  {message.content}
                </div>
              ) : (
                <Markdown content={message.content} />
              )
            ) : (message.isStreaming && !message.toolCalls?.length ? (
              <span className="animate-pulse">...</span>
            ) : null)}
          </div>
        </div>

        {/* Timestamp */}
        <div className="border-t px-3 py-1.5">
          <span className="text-xs text-muted-foreground">
            {formatTime(message.timestamp)}
          </span>
        </div>
      </div>
    </div>
  )
}

function formatTime(date: Date): string {
  return date.toLocaleTimeString(undefined, {
    hour: '2-digit',
    minute: '2-digit',
  })
}
