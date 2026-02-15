import { useState } from 'react'
import { Wrench, CheckCircle, Loader2, ChevronDown, ChevronUp } from 'lucide-react'
import { Badge } from '@/components/ui/badge'
import { Button } from '@/components/ui/button'
import type { ToolCall } from '@/hooks/useChat'

interface ToolCallDisplayProps {
  toolCall: ToolCall
  isStreaming?: boolean
}

export function ToolCallDisplay({ toolCall, isStreaming }: ToolCallDisplayProps) {
  const isRunning = toolCall.status === 'running' || isStreaming
  const isCompleted = toolCall.status === 'completed'
  const isError = toolCall.status === 'error'
  const [expanded, setExpanded] = useState(false)

  return (
    <div className="flex min-w-0 flex-col gap-1 rounded-lg border bg-muted/50 p-3 text-sm">
      <div className="flex items-center gap-2">
        {isRunning ? (
          <Loader2 className="h-4 w-4 animate-spin text-blue-500" />
        ) : isError ? (
          <Wrench className="h-4 w-4 text-red-500" />
        ) : isCompleted ? (
          <CheckCircle className="h-4 w-4 text-green-500" />
        ) : (
          <Wrench className="h-4 w-4 text-muted-foreground" />
        )}
        <span className="font-medium">{toolCall.name}</span>
        <Badge variant={isError ? 'destructive' : isCompleted ? 'success' : 'secondary'} className="ml-auto">
          {isRunning ? 'Running' : isError ? 'Error' : isCompleted ? 'Completed' : 'Pending'}
        </Badge>
        <Button
          variant="ghost"
          size="sm"
          className="ml-1 h-7 px-2"
          onClick={() => setExpanded((prev) => !prev)}
        >
          {expanded ? (
            <>
              <ChevronUp className="mr-1 h-3 w-3" />
              Hide
            </>
          ) : (
            <>
              <ChevronDown className="mr-1 h-3 w-3" />
              Details
            </>
          )}
        </Button>
      </div>

      {expanded && (
        <>
          {toolCall.args && toolCall.args !== '{}' && (
            <pre className="mt-2 max-w-full overflow-x-auto whitespace-pre-wrap break-words rounded bg-background/50 p-2 text-xs text-muted-foreground">
              {formatArgs(toolCall.args)}
            </pre>
          )}

          {toolCall.result && (
            <div className="mt-2 min-w-0 max-w-full">
              <span className="text-xs font-medium text-muted-foreground">Result:</span>
              <pre className="mt-1 max-h-32 max-w-full overflow-auto whitespace-pre-wrap break-words rounded bg-background/50 p-2 text-xs text-muted-foreground">
                {truncateResult(toolCall.result, 500)}
              </pre>
            </div>
          )}
        </>
      )}
    </div>
  )
}

function formatArgs(args: string): string {
  try {
    const parsed = JSON.parse(args)
    return JSON.stringify(parsed, null, 2)
  } catch {
    return args
  }
}

function truncateResult(result: string, maxLength: number): string {
  if (result.length <= maxLength) {
    try {
      const parsed = JSON.parse(result)
      return JSON.stringify(parsed, null, 2)
    } catch {
      return result
    }
  }
  return result.slice(0, maxLength) + '...'
}

interface ToolCallListProps {
  toolCalls: ToolCall[]
  isStreaming?: boolean
}

export function ToolCallList({ toolCalls, isStreaming }: ToolCallListProps) {
  if (toolCalls.length === 0) return null

  return (
    <div className="mt-2 flex min-w-0 flex-col gap-2">
      {toolCalls.map((toolCall, index) => (
        <ToolCallDisplay
          key={toolCall.id || index}
          toolCall={toolCall}
          isStreaming={isStreaming && index === toolCalls.length - 1}
        />
      ))}
    </div>
  )
}
