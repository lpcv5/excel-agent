"""Model provider configuration for Excel Agent.

This module provides a flexible model configuration system that supports:
- Predefined providers (zhipu, openai)
- Custom providers via environment variables
- Direct model class instantiation

Usage:
    # Using predefined provider
    provider = ModelProvider.from_string("zhipu:glm-5")
    model = provider.create_model()

    # Using custom provider via environment
    provider = ModelProvider.from_string("custom:my-model")
    # Requires CUSTOM_API_KEY and CUSTOM_API_BASE env vars

    # Direct model class
    from langchain_openai import ChatOpenAI
    model = ChatOpenAI(model="gpt-4")
"""

import os
from dataclasses import dataclass, field
from typing import Any, Optional

from langchain_core.language_models import BaseChatModel
from langchain_openai import ChatOpenAI


@dataclass
class ProviderConfig:
    """Configuration for a model provider.

    Attributes:
        name: Provider name (e.g., "zhipu", "openai")
        api_key_env: Environment variable name for API key
        api_base: Base URL for API (optional)
        default_model: Default model name
        extra_params: Additional parameters for ChatOpenAI
    """

    name: str
    api_key_env: str
    api_base: Optional[str] = None
    default_model: str = ""
    extra_params: dict[str, Any] = field(default_factory=dict)


# Predefined provider configurations
PREDEFINED_PROVIDERS: dict[str, ProviderConfig] = {
    "zhipu": ProviderConfig(
        name="zhipu",
        api_key_env="ZAI_API_KEY",
        api_base="https://open.bigmodel.cn/api/coding/paas/v4",
        default_model="glm-5",
        extra_params={"temperature": 0.6},
    ),
    "openai": ProviderConfig(
        name="openai",
        api_key_env="OPENAI_API_KEY",
        api_base=None,  # Use default OpenAI base
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
    """Model provider that creates LLM instances.

    Attributes:
        provider: Provider name (e.g., "zhipu", "openai", "custom")
        model_name: Model name (e.g., "glm-5", "gpt-4o")
        config: Provider configuration
    """

    provider: str
    model_name: str
    config: Optional[ProviderConfig] = None

    @classmethod
    def from_string(cls, model_string: str) -> "ModelProvider":
        """Parse model string and create provider.

        Args:
            model_string: Model string in format "provider:model" or just "model"

        Returns:
            ModelProvider instance

        Examples:
            >>> provider = ModelProvider.from_string("zhipu:glm-5")
            >>> provider = ModelProvider.from_string("openai:gpt-4o")
            >>> provider = ModelProvider.from_string("gpt-4")  # defaults to openai
        """
        if ":" in model_string:
            provider_name, model_name = model_string.split(":", 1)
        else:
            provider_name = "openai"
            model_name = model_string

        config = PREDEFINED_PROVIDERS.get(provider_name)

        # For unknown providers, try to create config from environment
        if config is None:
            config = cls._create_custom_provider_config(provider_name)

        return cls(provider=provider_name, model_name=model_name, config=config)

    @staticmethod
    def _create_custom_provider_config(provider_name: str) -> ProviderConfig:
        """Create config for custom provider from environment variables.

        Environment variables:
            {PROVIDER}_API_KEY: API key for the provider
            {PROVIDER}_API_BASE: Base URL for the API (optional)

        Args:
            provider_name: Provider name (uppercased for env var lookup)

        Returns:
            ProviderConfig instance
        """
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
        """Create and return the LLM instance.

        Returns:
            BaseChatModel instance

        Raises:
            ValueError: If API key is not configured
        """
        if self.config is None:
            raise ValueError(f"Unknown provider: {self.provider}")

        api_key = os.getenv(self.config.api_key_env)
        if not api_key:
            raise ValueError(
                f"API key not found. Please set {self.config.api_key_env} environment variable."
            )

        params: dict[str, Any] = {
            "model": self.model_name,
            "openai_api_key": api_key,
            **self.config.extra_params,
        }

        if self.config.api_base:
            params["openai_api_base"] = self.config.api_base

        return ChatOpenAI(**params)

    def __str__(self) -> str:
        return f"{self.provider}:{self.model_name}"


def create_model(model_spec: str | BaseChatModel) -> BaseChatModel:
    """Create a model from specification or return existing model.

    Args:
        model_spec: Model specification string (e.g., "zhipu:glm-5")
                   or a BaseChatModel instance (returned as-is)

    Returns:
        BaseChatModel instance

    Examples:
        >>> model = create_model("zhipu:glm-5")
        >>> model = create_model(ChatOpenAI(model="gpt-4"))
    """
    if isinstance(model_spec, BaseChatModel):
        return model_spec

    provider = ModelProvider.from_string(model_spec)
    return provider.create_model()
