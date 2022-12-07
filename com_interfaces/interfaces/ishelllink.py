from ctypes import POINTER, Structure, byref, create_unicode_buffer, pointer
from ctypes.wintypes import BYTE, DWORD, INT, MAX_PATH, USHORT, WCHAR, WORD
from typing import Optional
from typing import Union as U

from com_interfaces import CArgObject, IUnknown, interface, method
from com_interfaces.interfaces import IPersistFile


################################ ENUMERATIONS #################################
class SLGP:  # SLGP_FLAGS https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L11230
    SHORTPATH = 1
    UNCPRIORITY = 2
    RAWPATH = 4
    RELATIVEPRIORITY = 8

################################# STRUCTURES ##################################
class FILETIME(Structure):
    _fields_ = [
        ("dwLowDateTime", DWORD),
        ("dwHighDateTime", DWORD),
    ]

class WIN32_FIND_DATA(Structure):
    _fields_ = [
        ("dwFileAttributes", DWORD),
        ("ftCreationTime", FILETIME),
        ("ftLastAccessTime", FILETIME),
        ("ftLastWriteTime", FILETIME),
        ("nFileSizeHigh", DWORD),
        ("nFileSizeLow", DWORD),
        ("dwReserved0", DWORD),
        ("dwReserved1", DWORD),
        ("cFileName", WCHAR * MAX_PATH),
        ("cAlternateFileName", WCHAR * 14),
        ("dwFileType", DWORD),  # Obsolete. Do not use.
        ("dwCreatorType", DWORD),  # Obsolete. Do not use
        ("wFinderFlags", WORD),  # Obsolete. Do not use
    ]

class SHITEMID(Structure):
    _fields_ = [
        ("cb", USHORT),
        ("abID", BYTE * 1),
    ]
class ITEMIDLIST(Structure):
    _fields_ = [
        ("mkid", SHITEMID)]
###############################################################################


class DT:
    "Additional data types for type hints"
    DWORD = U[DWORD, int]
    INT = U[INT, int]
    WIN32_FIND_DATA_p = U[POINTER(WIN32_FIND_DATA), CArgObject]
    PIDLIST_ABSOLUTE = POINTER(POINTER(ITEMIDLIST))  # https://microsoft.public.win32.programmer.ui.narkive.com/p5Xl5twk/where-is-pidlist-absolute-defined


@interface
class IShellLink(IUnknown):
    """
    Exposes methods that create, modify, and resolve Shell links.

    Shell Links https://learn.microsoft.com/en-us/windows/win32/shell/links
    IShellLinkW https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllinkw
    IShellLinkW (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L11527
    """
    clsid = CLSID_ShellLink = "{00021401-0000-0000-C000-000000000046}"
    iid = IID_IShellLink  = "{000214F9-0000-0000-C000-000000000046}"
    LPTSTR = WCHAR * MAX_PATH  # https://habr.com/ru/post/164193

    @method(index=3)
    def GetPath(self, pszFile: LPTSTR, cch: DT.INT, pfd: DT.WIN32_FIND_DATA_p, fFlags: DT.DWORD):
        """
        Gets the path and file name of the target of a Shell link object.
        """
    
    def get_path(self, path: Optional[str]) -> Optional[str]:
        "Helper method for GetPath"
        if path:
            pf = self.query_interface(IPersistFile)
            pf.load(path)
        # create_unicode_buffer is used instead of c_wchar_p for mutable strings
        buf = create_unicode_buffer(MAX_PATH)
        fd = WIN32_FIND_DATA()
        if not self.GetPath(buf, len(buf), byref(fd), SLGP.UNCPRIORITY):
            return buf.value
    
    @method(index=4)
    def GetIDList(self, ppidl: DT.PIDLIST_ABSOLUTE):
        """
        Gets the list of item identifiers for the target of a Shell link object.
        """

    def get_id_list(self):
        "Helper method for GetIDList"
        idlist = POINTER(ITEMIDLIST)()
        self.GetIDList(pointer(idlist))
        return idlist
