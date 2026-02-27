"""Model provider configuration for Excel Agent."""

import os
from dataclasses import dataclass, field
from typing import Any, Optional

from langchain_core.language_models import BaseChatModel
from langchain_openai import ChatOpenAI


@dataclass
class ProviderConfig:
    name: str
    api_key_env: str
    api_base: Optional[str] = None
    default_model: str = ""
    extra_params: dict[str, Any] = field(default_factory=dict)


PREDEFINED_PROVIDERS: dict[str, ProviderConfig] = {
    "zhipu": ProviderConfig(
        name="zhipu",
        api_key_env="ZAI_API_KEY",
        api_base="https://open.bigmodel.cn/api/coding/paas/v4",
        default_model="glm-4.7",
        extra_params={"temperature": 0.6},
    ),
    "openai": ProviderConfig(
        name="openai",
        api_key_env="OPENAI_API_KEY",
        api_base=None,
        default_model="gpt-4o-mini",
        extra_params={"temperature": 0.7},
    ),
    "deepseek": ProviderConfig(
        name="deepseek",
        api_key_env="DEEPSEEK_API_KEY",
        api_base="https://api.deepseek.com/v1",
        default_model="deepseek-chat",
        extra_params={"temperature": 0.7},
    ),
    "moonshot": ProviderConfig(
        name="moonshot",
        api_key_env="MOONSHOT_API_KEY",
        api_base="https://api.moonshot.cn/v1",
        default_model="moonshot-v1-8k",
        extra_params={"temperature": 0.7},
    ),
}


@dataclass
class ModelProvider:
    provider: str
    model_name: str
    config: Optional[ProviderConfig] = None
    explicit_api_key: Optional[str] = None
    explicit_base_url: Optional[str] = None

    @classmethod
    def from_string(cls, model_string: str) -> "ModelProvider":
        if ":" in model_string:
            provider_name, model_name = model_string.split(":", 1)
        else:
            provider_name = "openai"
            model_name = model_string

        config = PREDEFINED_PROVIDERS.get(provider_name)
        if config is None:
            config = cls._create_custom_provider_config(provider_name)

        return cls(provider=provider_name, model_name=model_name, config=config)

    @classmethod
    def from_model_entry(cls, entry: Any) -> "ModelProvider":
        """Create a ModelProvider from a ModelEntry (settings_service.ModelEntry)."""
        config = PREDEFINED_PROVIDERS.get(entry.provider)
        if config is None:
            upper = entry.provider.upper()
            config = ProviderConfig(
                name=entry.provider,
                api_key_env=f"{upper}_API_KEY",
                api_base=entry.base_url,
                default_model="",
                extra_params={"temperature": 0.7},
            )
        return cls(
            provider=entry.provider,
            model_name=entry.model_name,
            config=config,
            explicit_api_key=entry.api_key or None,
            explicit_base_url=entry.base_url or None,
        )

    @staticmethod
    def _create_custom_provider_config(provider_name: str) -> ProviderConfig:
        upper_name = provider_name.upper()
        api_key_env = f"{upper_name}_API_KEY"
        api_base_env = f"{upper_name}_API_BASE"
        api_base = os.getenv(api_base_env)
        return ProviderConfig(
            name=provider_name,
            api_key_env=api_key_env,
            api_base=api_base,
            default_model="",
            extra_params={"temperature": 0.7},
        )

    def create_model(self) -> BaseChatModel:
        if self.config is None:
            raise ValueError(f"Unknown provider: {self.provider}")

        api_key = self.explicit_api_key or os.getenv(self.config.api_key_env)
        if not api_key:
            raise ValueError(
                f"API key not found. Please set {self.config.api_key_env} environment variable."
            )

        base_url = self.explicit_base_url or self.config.api_base

        params: dict[str, Any] = {
            "model": self.model_name,
            "openai_api_key": api_key,
            **self.config.extra_params,
        }

        if base_url:
            params["openai_api_base"] = base_url

        return ChatOpenAI(**params)

    def __str__(self) -> str:
        return f"{self.provider}:{self.model_name}"


def create_model(model_spec: str | BaseChatModel) -> BaseChatModel:
    if isinstance(model_spec, BaseChatModel):
        return model_spec
    provider = ModelProvider.from_string(model_spec)
    return provider.create_model()
