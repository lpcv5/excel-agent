import { useState } from 'react'
import { Circle, Loader2, CheckCircle2, ChevronUp, ChevronDown, X } from 'lucide-react'
import { cn } from '@/lib/utils'
import type { TodoItem } from '@/types/events'

interface TodoPanelProps {
  todos: TodoItem[]
}

export function TodoPanel({ todos }: TodoPanelProps) {
  const [isCollapsed, setIsCollapsed] = useState(false)
  const [isHidden, setIsHidden] = useState(false)

  if (todos.length === 0 || isHidden) {
    return null
  }

  const completedCount = todos.filter((t) => t.status === 'completed').length
  const progressPercent = Math.round((completedCount / todos.length) * 100)
  const hasInProgress = todos.some((t) => t.status === 'in_progress')

  return (
    <div className="border-t bg-muted/30">
      {/* Progress bar at top */}
      <div className="h-1 w-full bg-muted">
        <div
          className={cn(
            'h-full transition-all duration-300',
            progressPercent === 100 ? 'bg-success' : 'bg-primary'
          )}
          style={{ width: `${progressPercent}%` }}
        />
      </div>

      <div className="px-4 py-2">
        {/* Header */}
        <div className="flex items-center justify-between">
          <button
            type="button"
            onClick={() => setIsCollapsed(!isCollapsed)}
            className="flex items-center gap-2 text-sm font-medium hover:text-foreground/80"
          >
            {isCollapsed ? (
              <ChevronUp className="h-4 w-4" />
            ) : (
              <ChevronDown className="h-4 w-4" />
            )}
            <span>Tasks</span>
            <span className="text-muted-foreground">
              {completedCount}/{todos.length} completed
            </span>
            {hasInProgress && (
              <Loader2 className="h-3 w-3 animate-spin text-primary" />
            )}
          </button>

          <button
            type="button"
            onClick={() => setIsHidden(true)}
            className="rounded p-1 text-muted-foreground hover:bg-muted hover:text-foreground"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {/* Task list - vertical layout, one task per row */}
        {!isCollapsed && (
          <div className="mt-2 flex flex-col gap-1">
            {todos.map((todo, index) => (
              <div
                key={index}
                className={cn(
                  'flex items-center gap-2 rounded-md px-2 py-1.5 text-sm transition-colors',
                  todo.status === 'in_progress' && 'bg-primary/10 text-primary',
                  todo.status === 'completed' && 'text-muted-foreground line-through'
                )}
              >
                {todo.status === 'pending' && (
                  <Circle className="h-4 w-4 shrink-0 text-muted-foreground" />
                )}
                {todo.status === 'in_progress' && (
                  <Loader2 className="h-4 w-4 shrink-0 animate-spin" />
                )}
                {todo.status === 'completed' && (
                  <CheckCircle2 className="h-4 w-4 shrink-0 text-success" />
                )}
                <span className="truncate">{todo.content}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  )
}
