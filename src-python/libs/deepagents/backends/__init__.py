"""Memory backends for pluggable file storage."""

from libs.deepagents.backends.composite import CompositeBackend
from libs.deepagents.backends.filesystem import FilesystemBackend
from libs.deepagents.backends.local_shell import LocalShellBackend
from libs.deepagents.backends.protocol import BackendProtocol
from libs.deepagents.backends.state import StateBackend
from libs.deepagents.backends.store import (
    BackendContext,
    NamespaceFactory,
    StoreBackend,
)

__all__ = [
    "BackendContext",
    "BackendProtocol",
    "CompositeBackend",
    "FilesystemBackend",
    "LocalShellBackend",
    "NamespaceFactory",
    "StateBackend",
    "StoreBackend",
]
