import { useState } from 'react'
import { Wrench, CheckCircle, Loader2, ChevronDown, ChevronRight } from 'lucide-react'
import { cn } from '@/lib/utils'
import type { ToolCall } from '@/hooks/useChat'
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
} from '@/components/ui/dialog'

interface ToolCallDisplayProps {
  toolCall: ToolCall
  isStreaming?: boolean
  compact?: boolean
  onBadgeClick?: (toolCall: ToolCall) => void
}

export function ToolCallDisplay({ toolCall, isStreaming, compact = false, onBadgeClick }: ToolCallDisplayProps) {
  const isRunning = toolCall.status === 'running' || isStreaming
  const isCompleted = toolCall.status === 'completed'
  const isError = toolCall.status === 'error'
  const [expanded, setExpanded] = useState(false)

  if (compact) {
    return (
      <button
        type="button"
        onClick={() => onBadgeClick?.(toolCall)}
        className={cn(
          'inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-xs font-medium transition-all',
          'hover:scale-105 hover:shadow-sm cursor-pointer',
          isRunning && 'bg-info text-white shadow-sm shadow-info/25',
          isCompleted && 'bg-success text-white shadow-sm shadow-success/25',
          isError && 'bg-destructive text-white shadow-sm shadow-destructive/25',
          !isRunning && !isCompleted && !isError && 'bg-muted-foreground text-white shadow-sm'
        )}
      >
        {isRunning ? (
          <Loader2 className="h-3.5 w-3.5 animate-spin" />
        ) : isCompleted ? (
          <CheckCircle className="h-3.5 w-3.5" />
        ) : (
          <Wrench className="h-3.5 w-3.5" />
        )}
        <span>{toolCall.name}</span>
      </button>
    )
  }

  return (
    <div
      className={cn(
        'rounded-lg border text-sm transition-colors',
        isRunning && 'border-info/50 bg-info/10',
        isCompleted && 'border-success/50 bg-success/10',
        isError && 'border-destructive/50 bg-destructive/10',
        !isRunning && !isCompleted && !isError && 'border-border bg-muted/30'
      )}
    >
      {/* Header - always visible */}
      <button
        type="button"
        onClick={() => setExpanded(!expanded)}
        className="flex w-full items-center gap-2 p-2 text-left hover:bg-muted/30"
      >
        {expanded ? (
          <ChevronDown className="h-4 w-4 shrink-0 text-muted-foreground" />
        ) : (
          <ChevronRight className="h-4 w-4 shrink-0 text-muted-foreground" />
        )}

        {isRunning ? (
          <Loader2 className="h-4 w-4 shrink-0 animate-spin text-info" />
        ) : isCompleted ? (
          <CheckCircle className="h-4 w-4 shrink-0 text-success" />
        ) : isError ? (
          <Wrench className="h-4 w-4 shrink-0 text-destructive" />
        ) : (
          <Wrench className="h-4 w-4 shrink-0 text-muted-foreground" />
        )}

        <span className="font-medium">{toolCall.name}</span>

        <span
          className={cn(
            'ml-auto rounded-full px-2 py-0.5 text-xs font-medium',
            isRunning && 'bg-info/15 text-info dark:text-info-foreground',
            isCompleted && 'bg-success/15 text-success dark:text-success-foreground',
            isError && 'bg-destructive/15 text-destructive dark:text-destructive-foreground'
          )}
        >
          {isRunning ? 'Running' : isCompleted ? 'Done' : isError ? 'Error' : 'Pending'}
        </span>
      </button>

      {/* Details - expandable */}
      {expanded && (
        <div className="border-t px-3 py-2">
          {toolCall.args && toolCall.args !== '{}' && (
            <div className="mb-2">
              <span className="text-xs font-medium text-muted-foreground">Arguments:</span>
              <pre className="mt-1 max-h-32 max-w-full overflow-auto whitespace-pre-wrap break-words rounded bg-background/50 p-2 text-xs">
                {formatArgs(toolCall.args)}
              </pre>
            </div>
          )}

          {toolCall.result && (
            <div>
              <span className="text-xs font-medium text-muted-foreground">Result:</span>
              <pre className="mt-1 max-h-48 max-w-full overflow-auto whitespace-pre-wrap break-words rounded bg-background/50 p-2 text-xs">
                {truncateResult(toolCall.result, 1000)}
              </pre>
            </div>
          )}
        </div>
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
  return result.slice(0, maxLength) + '\n... (truncated)'
}

interface ToolCallGroupProps {
  toolCalls: ToolCall[]
  isStreaming?: boolean
}

/**
 * Group multiple tool calls into a collapsible section
 */
export function ToolCallGroup({ toolCalls, isStreaming }: ToolCallGroupProps) {
  const [expanded, setExpanded] = useState(false)
  const [selectedTool, setSelectedTool] = useState<ToolCall | null>(null)

  if (toolCalls.length === 0) return null

  const completedCount = toolCalls.filter(tc => tc.status === 'completed').length
  const runningCount = toolCalls.filter(tc => tc.status === 'running').length
  const errorCount = toolCalls.filter(tc => tc.status === 'error').length

  return (
    <>
      <div className="rounded-lg border bg-muted/20 text-sm">
        {/* Summary header */}
        <button
          type="button"
          onClick={() => setExpanded(!expanded)}
          className="flex w-full items-center gap-2 p-2 text-left hover:bg-muted/30"
        >
          {expanded ? (
            <ChevronDown className="h-4 w-4 text-muted-foreground" />
          ) : (
            <ChevronRight className="h-4 w-4 text-muted-foreground" />
          )}

          <Wrench className="h-4 w-4 text-muted-foreground" />

          <span className="font-medium">
            {toolCalls.length === 1 ? '1 tool call' : `${toolCalls.length} tool calls`}
          </span>

          {/* Status summary */}
          <div className="ml-auto flex items-center gap-2 text-xs">
            {runningCount > 0 && (
              <span className="flex items-center gap-1 rounded-full bg-info/15 px-2 py-0.5 font-medium text-info dark:text-info-foreground">
                <Loader2 className="h-3 w-3 animate-spin" />
                {runningCount}
              </span>
            )}
            {completedCount > 0 && (
              <span className="rounded-full bg-success/15 px-2 py-0.5 font-medium text-success dark:text-success-foreground">
                {completedCount}
              </span>
            )}
            {errorCount > 0 && (
              <span className="rounded-full bg-destructive/15 px-2 py-0.5 font-medium text-destructive dark:text-destructive-foreground">
                {errorCount}
              </span>
            )}
          </div>
        </button>

        {/* Tool calls list - only show if expanded */}
        {expanded && (
          <div className="border-t p-2">
            <div className="flex flex-wrap gap-1.5">
              {toolCalls.map((toolCall, index) => (
                <ToolCallDisplay
                  key={toolCall.id || index}
                  toolCall={toolCall}
                  isStreaming={isStreaming && index === toolCalls.length - 1 && toolCall.status === 'running'}
                  compact
                  onBadgeClick={setSelectedTool}
                />
              ))}
            </div>
          </div>
        )}
      </div>

      {/* Tool Detail Dialog */}
      <Dialog open={!!selectedTool} onOpenChange={(open) => !open && setSelectedTool(null)}>
        <DialogContent className="max-w-2xl max-h-[80vh] overflow-hidden flex flex-col">
          <DialogHeader>
            <div className="flex items-center gap-2">
              {selectedTool?.status === 'running' ? (
                <Loader2 className="h-5 w-5 animate-spin text-info" />
              ) : selectedTool?.status === 'completed' ? (
                <CheckCircle className="h-5 w-5 text-success" />
              ) : (
                <Wrench className="h-5 w-5 text-destructive" />
              )}
              <DialogTitle className="font-mono">
                {selectedTool?.name || 'Tool Call'}
              </DialogTitle>
            </div>
          </DialogHeader>

          {selectedTool && (
            <div className="flex-1 overflow-auto space-y-4">
              {/* Arguments */}
              {selectedTool.args && selectedTool.args !== '{}' && (
                <div>
                  <h4 className="mb-2 text-sm font-medium text-muted-foreground">Arguments</h4>
                  <pre className="overflow-auto whitespace-pre-wrap break-words rounded-lg bg-muted p-3 text-xs font-mono">
                    {formatArgs(selectedTool.args)}
                  </pre>
                </div>
              )}

              {/* Result */}
              {selectedTool.result && (
                <div>
                  <h4 className="mb-2 text-sm font-medium text-muted-foreground">Result</h4>
                  <pre className="max-h-64 overflow-auto whitespace-pre-wrap break-words rounded-lg bg-muted p-3 text-xs font-mono">
                    {selectedTool.result}
                  </pre>
                </div>
              )}

              {/* Status */}
              <div className="flex items-center gap-2">
                <span className="text-sm text-muted-foreground">Status:</span>
                <span
                  className={cn(
                    'rounded-full px-2.5 py-0.5 text-xs font-medium',
                    selectedTool.status === 'running' && 'bg-info/15 text-info dark:text-info-foreground',
                    selectedTool.status === 'completed' && 'bg-success/15 text-success dark:text-success-foreground',
                    selectedTool.status === 'error' && 'bg-destructive/15 text-destructive dark:text-destructive-foreground',
                    selectedTool.status === 'pending' && 'bg-muted text-muted-foreground'
                  )}
                >
                  {selectedTool.status}
                </span>
              </div>
            </div>
          )}
        </DialogContent>
      </Dialog>
    </>
  )
}

interface ToolCallListProps {
  toolCalls: ToolCall[]
  isStreaming?: boolean
}

export function ToolCallList({ toolCalls, isStreaming }: ToolCallListProps) {
  if (toolCalls.length === 0) return null

  return (
    <div className="flex flex-col gap-1.5">
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
