import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import rehypeHighlight from "rehype-highlight";

export function MarkdownRenderer({ content }: { content: string }) {
  return (
    <div className="markdown-body">
      <ReactMarkdown
        remarkPlugins={[remarkGfm]}
        rehypePlugins={[rehypeHighlight]}
        children={content}
        components={{
          pre({ children }) {
            return <pre className="rounded-lg border bg-surface p-3 overflow-x-auto text-xs-plus">{children}</pre>;
          },
          code({ className, children, ...props }) {
            if (!className) {
              return <code className="rounded bg-surface px-1 py-0.5 text-xs-plus font-mono" {...props}>{children}</code>;
            }
            return <code className={className} {...props}>{children}</code>;
          },
          table({ children }) {
            return <div className="my-2 overflow-x-auto rounded-lg border"><table className="min-w-full text-xs-plus">{children}</table></div>;
          },
          th({ children }) {
            return <th className="border-b bg-surface px-3 py-2 text-left text-xs font-semibold text-muted-foreground">{children}</th>;
          },
          td({ children }) {
            return <td className="border-b px-3 py-2 text-foreground">{children}</td>;
          },
          blockquote({ children }) {
            return <blockquote className="border-l-2 border-primary/40 pl-3 italic text-muted-foreground">{children}</blockquote>;
          },
          a({ href, children }) {
            return <a href={href} target="_blank" rel="noopener noreferrer" className="text-primary underline underline-offset-2 hover:text-primary/80">{children}</a>;
          },
        }}
      />
    </div>
  );
}
