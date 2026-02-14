import type { Message, MessageBlock, ToolCall } from "@/types";

function formatToolCall(call: ToolCall): string {
  const lines = [`[Tool: ${call.name}] (${call.status})`];
  if (call.args && Object.keys(call.args).length > 0) {
    lines.push(`Params: ${JSON.stringify(call.args, null, 2)}`);
  }
  if (call.result) {
    lines.push(`Result: ${call.result}`);
  }
  return lines.join("\n");
}

function formatBlock(block: MessageBlock): string {
  if (block.type === "text") return block.content;
  if (block.type === "thinking") return "";
  return formatToolCall(block.call);
}

/** Serialize a message including tool calls into copyable plain text. */
export function formatMessageForCopy(message: Message): string {
  const blocks = message.blocks ?? [];
  if (blocks.length === 0) return message.content;
  return blocks.map(formatBlock).filter(Boolean).join("\n\n");
}
