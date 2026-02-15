import { ScrollArea } from '@/components/ui/scroll-area'
import { Button } from '@/components/ui/button'
import { MessageItem } from './MessageItem'
import type { Message } from '@/hooks/useChat'

interface MessageListProps {
  messages: Message[]
  messageEndRef: React.RefObject<HTMLDivElement | null>
  viewportRef?: React.RefCallback<HTMLDivElement> | React.Ref<HTMLDivElement>
  onExampleClick?: (text: string) => void
}

export function MessageList({ messages, messageEndRef, viewportRef, onExampleClick }: MessageListProps) {
  if (messages.length === 0) {
    return (
      <div className="flex h-full flex-col items-center justify-center gap-4 text-muted-foreground">
        <div className="text-center">
          <h2 className="text-xl font-semibold">Excel Agent</h2>
          <p className="mt-2 text-sm">Ask me to help with your Excel files</p>
        </div>
        <div className="grid max-w-md gap-2 text-sm">
          <ExampleQuery text="Read sales.xlsx and show the first 10 rows" onClick={onExampleClick} />
          <ExampleQuery text="Create a pivot table from the data in Sheet1" onClick={onExampleClick} />
          <ExampleQuery text="Apply conditional formatting to column B" onClick={onExampleClick} />
        </div>
      </div>
    )
  }

  return (
    <ScrollArea className="h-full" viewportRef={viewportRef}>
      <div className="flex min-w-0 flex-col gap-4 p-4">
        {messages.map((message) => (
          <MessageItem key={message.id} message={message} />
        ))}
        <div ref={messageEndRef} />
      </div>
    </ScrollArea>
  )
}

interface ExampleQueryProps {
  text: string
  onClick?: (text: string) => void
}

function ExampleQuery({ text, onClick }: ExampleQueryProps) {
  return (
    <Button
      variant="ghost"
      className="h-auto justify-start rounded-lg border bg-muted/50 p-3 text-left text-muted-foreground hover:bg-muted"
      onClick={() => onClick?.(text)}
    >
      {text}
    </Button>
  )
}
