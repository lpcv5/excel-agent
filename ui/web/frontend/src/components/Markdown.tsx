import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'

interface MarkdownProps {
  content: string
}

export function Markdown({ content }: MarkdownProps) {
  return (
    <ReactMarkdown
      remarkPlugins={[remarkGfm]}
      components={{
        p: ({ children }) => (
          <p className="whitespace-pre-wrap break-words text-sm leading-relaxed">
            {children}
          </p>
        ),
        a: ({ href, children }) => (
          <a
            href={href}
            className="text-primary underline underline-offset-4 hover:opacity-80"
            target="_blank"
            rel="noreferrer"
          >
            {children}
          </a>
        ),
        code: ({ children }) => (
          <code className="rounded bg-muted px-1.5 py-0.5 text-xs">
            {children}
          </code>
        ),
        pre: ({ children }) => (
          <pre className="max-w-full overflow-x-auto rounded bg-muted p-2 text-xs">
            {children}
          </pre>
        ),
        ul: ({ children }) => (
          <ul className="list-disc space-y-1 pl-5 text-sm">{children}</ul>
        ),
        ol: ({ children }) => (
          <ol className="list-decimal space-y-1 pl-5 text-sm">{children}</ol>
        ),
        blockquote: ({ children }) => (
          <blockquote className="border-l-2 border-muted-foreground/30 pl-3 text-sm text-muted-foreground">
            {children}
          </blockquote>
        ),
        table: ({ children }) => (
          <div className="max-w-full overflow-x-auto">
            <table className="w-full border-collapse text-sm">{children}</table>
          </div>
        ),
        th: ({ children }) => (
          <th className="border px-2 py-1 text-left font-medium">{children}</th>
        ),
        td: ({ children }) => (
          <td className="border px-2 py-1 align-top">{children}</td>
        ),
        h1: ({ children }) => (
          <h1 className="text-lg font-semibold">{children}</h1>
        ),
        h2: ({ children }) => (
          <h2 className="text-base font-semibold">{children}</h2>
        ),
        h3: ({ children }) => (
          <h3 className="text-sm font-semibold">{children}</h3>
        ),
        h4: ({ children }) => (
          <h4 className="text-sm font-medium">{children}</h4>
        ),
        h5: ({ children }) => (
          <h5 className="text-sm font-medium">{children}</h5>
        ),
        h6: ({ children }) => (
          <h6 className="text-sm font-medium">{children}</h6>
        ),
      }}
    >
      {content}
    </ReactMarkdown>
  )
}
