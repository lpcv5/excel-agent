export type Theme = "light" | "dark" | "system";

export type ToolCallStatus = "running" | "success" | "error";

export interface ToolCall {
  id: string;
  name: string;
  status: ToolCallStatus;
  args?: Record<string, unknown>;
  result?: string;
  duration_ms?: number;
}

export interface TaskItem {
  id: string;
  label: string;
  status: "pending" | "in_progress" | "completed";
}

export interface Conversation {
  id: string;
  title: string;
  createdAt: string;
  updatedAt: string;
  messageCount: number;
}

export type MessageBlock =
  | { type: "thinking"; content: string }
  | { type: "text"; content: string }
  | { type: "tool_call"; call: ToolCall };

export interface Message {
  id: string;
  conversationId: string;
  role: "user" | "assistant" | "system";
  content: string;
  createdAt: string;
  blocks?: MessageBlock[];
  tasks?: TaskItem[];
  error?: string;
}
