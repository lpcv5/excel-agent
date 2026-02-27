import { useState, useCallback, useEffect, type KeyboardEvent } from 'react'
import { Send, Loader2 } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'

interface InputBarProps {
  onSend: (query: string) => void
  isStreaming: boolean
  disabled?: boolean
  quickQuery?: string
}

export function InputBar({ onSend, isStreaming, disabled, quickQuery }: InputBarProps) {
  const [input, setInput] = useState('')

  // Handle quick query from example clicks
  useEffect(() => {
    if (quickQuery) {
      setInput(quickQuery)
    }
  }, [quickQuery])

  const handleSend = useCallback(() => {
    const trimmed = input.trim()
    if (trimmed && !isStreaming && !disabled) {
      onSend(trimmed)
      setInput('')
    }
  }, [input, isStreaming, disabled, onSend])

  const handleKeyDown = useCallback(
    (e: KeyboardEvent<HTMLInputElement>) => {
      if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault()
        handleSend()
      }
    },
    [handleSend]
  )

  return (
    <div className="flex gap-2 border-t bg-background p-4">
      <Input
        value={input}
        onChange={(e) => setInput(e.target.value)}
        onKeyDown={handleKeyDown}
        placeholder="Ask about your Excel files..."
        disabled={disabled || isStreaming}
        className="flex-1"
      />
      <Button
        onClick={handleSend}
        disabled={disabled || isStreaming || !input.trim()}
        size="icon"
      >
        {isStreaming ? (
          <Loader2 className="h-4 w-4 animate-spin" />
        ) : (
          <Send className="h-4 w-4" />
        )}
      </Button>
    </div>
  )
}
