from importlib.metadata import PackageNotFoundError, version  # type: ignore

from com_interfaces.iunknown import (CArgObject, Guid, IUnknown, interface,
                                     method, structure)


try:
    __version__ = version(__name__)
except PackageNotFoundError:  # pragma: no cover
    # tried to get version of the package but it's not installed in environment
    __version__ = "unknown"
