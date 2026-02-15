/**
 * Event types for streaming agent output.
 * Corresponds to excel_agent/events.py
 */

export type EventType =
  | 'query_start'
  | 'query_end'
  | 'error'
  | 'thinking'
  | 'text'
  | 'refusal'
  | 'tool_call_start'
  | 'tool_call_args'
  | 'tool_call_end'
  | 'tool_result'

export interface AgentEvent {
  type: EventType
  timestamp: string
  content: string | null
  tool_name: string | null
  tool_args: string | null
  tool_call_id: string | null
  error_message: string | null
  data: Record<string, unknown>
}

// Type guards for event types
export function isQueryStartEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'query_start'
}

export function isQueryEndEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'query_end'
}

export function isErrorEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'error'
}

export function isThinkingEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'thinking'
}

export function isTextEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'text'
}

export function isRefusalEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'refusal'
}

export function isToolCallStartEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'tool_call_start'
}

export function isToolCallArgsEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'tool_call_args'
}

export function isToolResultEvent(event: AgentEvent): event is AgentEvent {
  return event.type === 'tool_result'
}
