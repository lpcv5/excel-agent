"""LLM call logging callback handler."""

import json
import logging
import time
from datetime import datetime
from typing import Any, Optional
from uuid import UUID

from langchain_core.callbacks import BaseCallbackHandler
from langchain_core.outputs import LLMResult


class LLMLoggingCallbackHandler(BaseCallbackHandler):
    def __init__(
        self,
        logger: Optional[logging.Logger] = None,
        log_full_prompt: bool = True,
        log_token_usage: bool = True,
        log_timing: bool = True,
    ):
        self.logger = logger or logging.getLogger("agent.llm")
        self.log_full_prompt = log_full_prompt
        self.log_token_usage = log_token_usage
        self.log_timing = log_timing
        self._call_start_times: dict[str, float] = {}
        self._call_prompts: dict[str, list] = {}

    def on_chat_model_start(
        self, serialized: dict[str, Any], messages: list, *, run_id: UUID, **kwargs: Any
    ) -> None:
        run_id_str = str(run_id)
        self._call_start_times[run_id_str] = time.perf_counter()
        self._call_prompts[run_id_str] = messages

        log_data: dict[str, Any] = {
            "event": "llm_start",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_full_prompt:
            prompt_content = []
            for msg_list in messages:
                for msg in msg_list:
                    prompt_content.append(
                        {
                            "type": msg.__class__.__name__,
                            "content": getattr(msg, "content", str(msg)),
                        }
                    )
            log_data["prompt"] = prompt_content

        if "invocation_params" in kwargs:
            params = kwargs["invocation_params"]
            log_data["model_params"] = {
                k: v
                for k, v in params.items()
                if k in ["model", "model_name", "temperature", "stream"]
            }

        self.logger.debug(
            f"LLM Request:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

    def on_chat_model_end(
        self, response: LLMResult, *, run_id: UUID, **kwargs: Any
    ) -> None:
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_end",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_timing and start_time:
            log_data["elapsed_ms"] = round((time.perf_counter() - start_time) * 1000, 2)

        if response.generations:
            log_data["response"] = [
                {"text": gen.text}
                for gen_list in response.generations
                for gen in gen_list
            ]

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
        self._call_prompts.pop(run_id_str, None)

    def on_chat_model_error(
        self, error: BaseException, *, run_id: UUID, **kwargs: Any
    ) -> None:
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
            log_data["elapsed_ms"] = round((time.perf_counter() - start_time) * 1000, 2)

        self.logger.error(
            f"LLM Error:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )
        self._call_prompts.pop(run_id_str, None)

    def on_llm_start(
        self,
        serialized: dict[str, Any],
        prompts: list[str],
        *,
        run_id: UUID,
        **kwargs: Any,
    ) -> None:
        run_id_str = str(run_id)
        self._call_start_times[run_id_str] = time.perf_counter()
        self._call_prompts[run_id_str] = prompts

        log_data: dict[str, Any] = {
            "event": "llm_start",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_full_prompt:
            log_data["prompt"] = [{"type": "prompt", "content": p} for p in prompts]

        if "invocation_params" in kwargs:
            params = kwargs["invocation_params"]
            log_data["model_params"] = {
                k: v
                for k, v in params.items()
                if k in ["model", "model_name", "temperature", "stream"]
            }

        self.logger.debug(
            f"LLM Request:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )

    def on_llm_end(self, response: LLMResult, *, run_id: UUID, **kwargs: Any) -> None:
        run_id_str = str(run_id)
        start_time = self._call_start_times.pop(run_id_str, None)

        log_data: dict[str, Any] = {
            "event": "llm_end",
            "timestamp": datetime.now().isoformat(),
            "run_id": run_id_str,
        }

        if self.log_timing and start_time:
            log_data["elapsed_ms"] = round((time.perf_counter() - start_time) * 1000, 2)

        if response.generations:
            log_data["response"] = [
                {"text": gen.text}
                for gen_list in response.generations
                for gen in gen_list
            ]

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
        self._call_prompts.pop(run_id_str, None)

    def on_llm_error(
        self, error: BaseException, *, run_id: UUID, **kwargs: Any
    ) -> None:
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
            log_data["elapsed_ms"] = round((time.perf_counter() - start_time) * 1000, 2)

        self.logger.error(
            f"LLM Error:\n{json.dumps(log_data, indent=2, ensure_ascii=False, default=str)}"
        )
        self._call_prompts.pop(run_id_str, None)
