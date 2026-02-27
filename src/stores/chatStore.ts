import { create } from "zustand";
import type { Conversation, Message, ToolCall, TaskItem } from "@/types";

const uid = () => crypto.randomUUID();
const now = () => new Date().toISOString();

interface ChatState {
  conversations: Conversation[];
  activeConversationId: string | null;
  messages: Record<string, Message[]>;
  streaming: boolean;
  streamingMessageId: string | null;

  createConversation: (title?: string) => string;
  deleteConversation: (id: string) => void;
  setActiveConversation: (id: string | null) => void;
  addUserMessage: (content: string) => Message | null;
  addAssistantMessage: () => Message | null;
  appendToThinking: (messageId: string, token: string) => void;
  appendToContent: (messageId: string, token: string) => void;
  addToolCall: (messageId: string, call: ToolCall) => void;
  updateToolCall: (messageId: string, callId: string, patch: Partial<ToolCall>) => void;
  setTasks: (messageId: string, tasks: TaskItem[]) => void;
  setStreamingDone: (error?: string) => void;
}

function updateMsg(
  messages: Record<string, Message[]>,
  convId: string,
  messageId: string,
  updater: (m: Message) => Message,
) {
  return {
    messages: {
      ...messages,
      [convId]: (messages[convId] ?? []).map((m) =>
        m.id === messageId ? updater(m) : m,
      ),
    },
  };
}

export const useChatStore = create<ChatState>((set, get) => ({
  conversations: [],
  activeConversationId: null,
  messages: {},
  streaming: false,
  streamingMessageId: null,

  createConversation: (title?) => {
    const id = uid();
    set((s) => ({
      conversations: [
        { id, title: title || "新对话", createdAt: now(), updatedAt: now(), messageCount: 0 },
        ...s.conversations,
      ],
      messages: { ...s.messages, [id]: [] },
      activeConversationId: id,
    }));
    return id;
  },

  deleteConversation: (id) => {
    set((s) => {
      const convs = s.conversations.filter((c) => c.id !== id);
      const msgs = { ...s.messages };
      delete msgs[id];
      return {
        conversations: convs,
        messages: msgs,
        activeConversationId:
          s.activeConversationId === id ? (convs[0]?.id ?? null) : s.activeConversationId,
      };
    });
  },

  setActiveConversation: (id) => set({ activeConversationId: id }),

  addUserMessage: (content) => {
    let convId = get().activeConversationId;
    if (!convId) convId = get().createConversation(content.slice(0, 30));
    const msg: Message = { id: uid(), conversationId: convId, role: "user", content, createdAt: now() };
    set((s) => {
      const prev = s.messages[convId!] ?? [];
      return {
        messages: { ...s.messages, [convId!]: [...prev, msg] },
        conversations: s.conversations.map((c) =>
          c.id === convId
            ? { ...c, updatedAt: now(), messageCount: c.messageCount + 1, title: prev.length === 0 ? content.slice(0, 30) : c.title }
            : c,
        ),
      };
    });
    return msg;
  },

  addAssistantMessage: () => {
    const convId = get().activeConversationId;
    if (!convId) return null;
    const msg: Message = {
      id: uid(), conversationId: convId, role: "assistant",
      content: "", blocks: [], tasks: [], createdAt: now(),
    };
    set((s) => ({
      messages: { ...s.messages, [convId]: [...(s.messages[convId] ?? []), msg] },
      conversations: s.conversations.map((c) =>
        c.id === convId ? { ...c, updatedAt: now(), messageCount: c.messageCount + 1 } : c,
      ),
      streaming: true,
      streamingMessageId: msg.id,
    }));
    return msg;
  },

  appendToThinking: (messageId, token) => {
    set((s) => {
      const convId = s.activeConversationId;
      if (!convId) return s;
      return updateMsg(s.messages, convId, messageId, (m) => {
        const blocks = [...(m.blocks ?? [])];
        const last = blocks[blocks.length - 1];
        if (last?.type === "thinking") {
          blocks[blocks.length - 1] = { type: "thinking", content: last.content + token };
        } else {
          blocks.push({ type: "thinking", content: token });
        }
        return { ...m, blocks };
      });
    });
  },

  appendToContent: (messageId, token) => {
    set((s) => {
      const convId = s.activeConversationId;
      if (!convId) return s;
      return updateMsg(s.messages, convId, messageId, (m) => {
        const blocks = [...(m.blocks ?? [])];
        const last = blocks[blocks.length - 1];
        if (last?.type === "text") {
          blocks[blocks.length - 1] = { type: "text", content: last.content + token };
        } else {
          blocks.push({ type: "text", content: token });
        }
        return { ...m, content: m.content + token, blocks };
      });
    });
  },

  addToolCall: (messageId, call) => {
    set((s) => {
      const convId = s.activeConversationId;
      if (!convId) return s;
      return updateMsg(s.messages, convId, messageId, (m) => ({
        ...m,
        blocks: [...(m.blocks ?? []), { type: "tool_call", call }],
      }));
    });
  },

  updateToolCall: (messageId, callId, patch) => {
    set((s) => {
      const convId = s.activeConversationId;
      if (!convId) return s;
      return updateMsg(s.messages, convId, messageId, (m) => ({
        ...m,
        blocks: (m.blocks ?? []).map((b) =>
          b.type === "tool_call" && b.call.id === callId
            ? { type: "tool_call", call: { ...b.call, ...patch } }
            : b,
        ),
      }));
    });
  },

  setTasks: (messageId, tasks) => {
    set((s) => {
      const convId = s.activeConversationId;
      if (!convId) return s;
      return updateMsg(s.messages, convId, messageId, (m) => ({ ...m, tasks }));
    });
  },

  setStreamingDone: (error?: string) => {
    const msgId = get().streamingMessageId;
    const convId = get().activeConversationId;
    if (error && msgId && convId) {
      set((s) => ({
        ...updateMsg(s.messages, convId, msgId, (m) => ({ ...m, error })),
        streaming: false,
        streamingMessageId: null,
      }));
    } else {
      set({ streaming: false, streamingMessageId: null });
    }
  },
}));
