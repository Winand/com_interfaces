from ctypes import c_wchar_p
from ctypes.wintypes import DWORD, BOOL
from typing import Union as U

from com_interfaces import IUnknown, interface, method, Guid


class DT:
    "Additional data types for type hints"
    DWORD = U[DWORD, int]
    LPCOLESTR = U[c_wchar_p, str]  # https://stackoverflow.com/a/1607840
    BOOL = U[BOOL, bool]


@interface
class IPersistFile(IUnknown):
    """
    Enables an object to be loaded from or saved to a disk file.

    IPersistFile https://learn.microsoft.com/en-us/windows/win32/api/objidl/nn-objidl-ipersistfile
    IPersist (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ObjIdl.h#L9140
    IPersistFile (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ObjIdl.h#L10336
    """
    clsid = CLSID_ShellLink = Guid("{00021401-0000-0000-C000-000000000046}")
    iid = IID_IPersistFile = Guid("{0000010b-0000-0000-C000-000000000046}")

    @method(index=5)
    def Load(self, pszFileName: DT.LPCOLESTR, dwMode: DT.DWORD):
        """
        Opens the specified file and initializes an object from the file contents
        """

    def load(self, pszFileName: str, dwMode: int=0):
        "Helper method for Load"
        buf = c_wchar_p(pszFileName)
        if self.Load(buf, dwMode):
            raise WindowsError("Load failed.")

    @method(index=6)
    def Save(self, pszFileName: DT.LPCOLESTR, fRemember: DT.BOOL):
        """
        Saves a copy of the object to the specified file
        """
