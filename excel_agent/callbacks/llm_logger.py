"""LLM call logging callback handler.

This module provides a LangChain callback handler that logs
LLM API calls for debugging purposes.
"""

import json
import logging
import time
from datetime import datetime
from typing import Any, Optional
from uuid import UUID

from langchain_core.callbacks import BaseCallbackHandler
from langchain_core.outputs import LLMResult


class LLMLoggingCallbackHandler(BaseCallbackHandler):
    """LangChain callback handler for logging LLM call details.

    Records:
    - Request timestamp, prompt, model parameters
    - Response timestamp, content, token usage
    - Call timing (elapsed time)
    - Error information if call fails
    """

    def __init__(
        self,
        logger: Optional[logging.Logger] = None,
        log_full_prompt: bool = True,
        log_token_usage: bool = True,
        log_timing: bool = True,
    ):
        """Initialize the callback handler.

        Args:
            logger: Optional logger instance. Creates new one if not provided.
            log_full_prompt: Whether to log full prompt content
            log_token_usage: Whether to log token usage statistics
            log_timing: Whether to log call timing information
        """
        self.logger = logger or logging.getLogger("excel_agent.llm")
        self.log_full_prompt = log_full_prompt
        self.log_token_usage = log_token_usage
        self.log_timing = log_timing

        # Temporary state for tracking calls
        self._call_start_times: dict[str, float] = {}
        self._call_prompts: dict[str, list] = {}

    def on_chat_model_start(
        self,
        serialized: dict[str, Any],
        messages: list,
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call starts."""
        run_id_str = str(run_id)
        self._call_start_times[run_id_str] = time.perf_counter()
        self._call_prompts[run_id_str] = messages

        log_data: dict[str, Any] = {
            "event": "llm_start",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_full_prompt:
            # Extract message content
            prompt_content = []
            for msg_list in messages:
                for msg in msg_list:
                    prompt_content.append({
                        "type": msg.__class__.__name__,
                        "content": getattr(msg, "content", str(msg)),
                    })
            log_data["prompt"] = prompt_content

        # Log model parameters (simplified - don't include tools to reduce log size)
        if "invocation_params" in kwargs:
            params = kwargs["invocation_params"]
            log_data["model_params"] = {
                k: v for k, v in params.items()
                if k in ["model", "model_name", "temperature", "stream"]
            }

        self.logger.debug(
            f"LLM Request:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

    def on_chat_model_end(
        self,
        response: LLMResult,
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call ends."""
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_end",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        # Calculate elapsed time
        if self.log_timing and start_time:
            elapsed_ms = (time.perf_counter() - start_time) * 1000
            log_data["elapsed_ms"] = round(elapsed_ms, 2)

        # Extract response content
        if response.generations:
            content = []
            for gen_list in response.generations:
                for gen in gen_list:
                    content.append({"text": gen.text})
            log_data["response"] = content

        # Log token usage
        if self.log_token_usage and response.llm_output:
            token_usage = response.llm_output.get("token_usage", {})
            if token_usage:
                log_data["token_usage"] = {
                    "prompt_tokens": token_usage.get("prompt_tokens"),
                    "completion_tokens": token_usage.get("completion_tokens"),
                    "total_tokens": token_usage.get("total_tokens"),
                }

        self.logger.debug(
            f"LLM Response:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

        # Clean up prompt cache
        self._call_prompts.pop(run_id_str, None)

    def on_chat_model_error(
        self,
        error: BaseException,
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call fails."""
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_error",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
            "error_type": type(error).__name__,
            "error_message": str(error),
        }

        if self.log_timing and start_time:
            elapsed_ms = (time.perf_counter() - start_time) * 1000
            log_data["elapsed_ms"] = round(elapsed_ms, 2)

        self.logger.error(
            f"LLM Error:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

        # Clean up
        self._call_prompts.pop(run_id_str, None)

    # Also implement on_llm_* methods for compatibility with non-chat models
    # and OpenAI Responses API

    def on_llm_start(
        self,
        serialized: dict[str, Any],
        prompts: list[str],
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call starts (non-chat mode)."""
        run_id_str = str(run_id)
        self._call_start_times[run_id_str] = time.perf_counter()
        self._call_prompts[run_id_str] = prompts

        log_data: dict[str, Any] = {
            "event": "llm_start",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_full_prompt:
            log_data["prompt"] = [
                {"type": "prompt", "content": p} for p in prompts
            ]

        # Log model parameters (simplified - don't include tools)
        if "invocation_params" in kwargs:
            params = kwargs["invocation_params"]
            log_data["model_params"] = {
                k: v for k, v in params.items()
                if k in ["model", "model_name", "temperature", "stream"]
            }

        self.logger.debug(
            f"LLM Request:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

    def on_llm_end(
        self,
        response: LLMResult,
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call ends (non-chat mode)."""
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_end",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        # Calculate elapsed time
        if self.log_timing and start_time:
            elapsed_ms = (time.perf_counter() - start_time) * 1000
            log_data["elapsed_ms"] = round(elapsed_ms, 2)

        # Extract response content
        if response.generations:
            content = []
            for gen_list in response.generations:
                for gen in gen_list:
                    content.append({"text": gen.text})
            log_data["response"] = content

        # Log token usage
        if self.log_token_usage and response.llm_output:
            token_usage = response.llm_output.get("token_usage", {})
            if token_usage:
                log_data["token_usage"] = {
                    "prompt_tokens": token_usage.get("prompt_tokens"),
                    "completion_tokens": token_usage.get("completion_tokens"),
                    "total_tokens": token_usage.get("total_tokens"),
                }

        self.logger.debug(
            f"LLM Response:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

        # Clean up
        self._call_prompts.pop(run_id_str, None)

    def on_llm_error(
        self,
        error: BaseException,
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        """Triggered when LLM call fails (non-chat mode)."""
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_error",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
            "error_type": type(error).__name__,
            "error_message": str(error),
        }

        if self.log_timing and start_time:
            elapsed_ms = (time.perf_counter() - start_time) * 1000
            log_data["elapsed_ms"] = round(elapsed_ms, 2)

        self.logger.error(
            f"LLM Error:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

        # Clean up
        self._call_prompts.pop(run_id_str, None)
