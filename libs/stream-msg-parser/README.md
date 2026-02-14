# stream-msg-parser

Custom message parser for LangGraph streaming with `stream_mode="messages"`.

## Overview

This library provides parsing for LangGraph AIMessage streaming output with `stream_mode="messages"`, supporting:

- Token-level content streaming
- Reasoning/thinking content detection
- Tool call lifecycle tracking (start/end events)

## Installation

```bash
pip install stream-msg-parser
```

## Usage

```python
from stream_msg_parser import MessageParser

parser = MessageParser(track_tool_lifecycle=True)

async for event in parser.aparse(stream):
    if isinstance(event, ContentEvent):
        print(f"Content: {event.content}")
    elif isinstance(event, ToolCallStartEvent):
        print(f"Tool call: {event.name}")
    elif isinstance(event, ToolCallEndEvent):
        print(f"Tool result: {event.result}")
```

## Event Types

- `ContentEvent` - Token-level content (text or reasoning)
- `ToolCallStartEvent` - Tool call initiated
- `ToolCallEndEvent` - Tool call completed
- `ErrorEvent` - Error occurred
- `CompleteEvent` - Stream completed
