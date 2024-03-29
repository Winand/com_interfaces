from ctypes import POINTER, byref, c_wchar_p, create_unicode_buffer, pointer
from ctypes.wintypes import BYTE, DWORD, INT, MAX_PATH, USHORT, WCHAR, WORD
from os import PathLike
from typing import Optional, Tuple
from typing import Union as U

from com_interfaces import CArgObject, IUnknown, interface, method, structure
from com_interfaces.interfaces import IPersistFile


################################ ENUMERATIONS #################################
class SLGP:  # SLGP_FLAGS https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L11230
    SHORTPATH = 1
    UNCPRIORITY = 2
    RAWPATH = 4
    RELATIVEPRIORITY = 8

################################# STRUCTURES ##################################
@structure
class FILETIME:
    dwLowDateTime: DWORD
    dwHighDateTime: DWORD

@structure
class WIN32_FIND_DATA:
    WCHAR__MAX_PATH = WCHAR * MAX_PATH
    WCHAR__14 = WCHAR * 14

    dwFileAttributes: DWORD
    ftCreationTime: FILETIME
    ftLastAccessTime: FILETIME
    ftLastWriteTime: FILETIME
    nFileSizeHigh: DWORD
    nFileSizeLow: DWORD
    dwReserved0: DWORD
    dwReserved1: DWORD
    cFileName: WCHAR__MAX_PATH
    cAlternateFileName: WCHAR__14
    dwFileType: DWORD  # Obsolete. Do not use.
    dwCreatorType: DWORD  # Obsolete. Do not use
    wFinderFlags: WORD  # Obsolete. Do not use

@structure
class SHITEMID:
    BYTE__1 = BYTE * 1

    cb: USHORT
    abID: BYTE__1

@structure
class ITEMIDLIST:
    mkid: SHITEMID
###############################################################################


class DT:
    "Additional data types for type hints"
    DWORD = U[DWORD, int]
    INT = U[INT, int]
    WIN32_FIND_DATA_p = U[POINTER(WIN32_FIND_DATA), CArgObject]
    # https://microsoft.public.win32.programmer.ui.narkive.com/p5Xl5twk/where-is-pidlist-absolute-defined noqa
    PIDLIST_ABSOLUTE = POINTER(POINTER(ITEMIDLIST))
    LPWSTR = WCHAR * MAX_PATH  # https://habr.com/ru/post/164193
    LPCWSTR = U[c_wchar_p, str]
    StrPath = U[PathLike, str]


@interface("{000214F9-0000-0000-C000-000000000046}", clsid="{00021401-0000-0000-C000-000000000046}")
class IShellLink(IUnknown):
    """
    Exposes methods that create, modify, and resolve Shell links.

    Shell Links https://learn.microsoft.com/en-us/windows/win32/shell/links
    IShellLinkW https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllinkw
    IShellLinkW (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L11527
    """

    @method(index=3)
    def GetPath(self, pszFile: DT.LPWSTR, cch: DT.INT, pfd: DT.WIN32_FIND_DATA_p, fFlags: DT.DWORD):
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

    @method(index=7)
    def SetDescription(self, pszName: DT.LPCWSTR):
        """
        Sets the description for a Shell link object
        """

    @method(index=9)
    def SetWorkingDirectory(self, pszDir: DT.LPCWSTR):
        """
        Sets the name of the working directory for a Shell link object
        """

    @method(index=17)
    def SetIconLocation(self, pszIconPath: DT.LPCWSTR, iIcon: DT.INT):
        """
        Sets the location (path and index) of the icon for a Shell link object
        """

    @method(index=20)
    def SetPath(self, pszFile: DT.LPCWSTR):
        """
        Sets the path and file name for the target of a Shell link object
        """

    def create_link(self, link_path: DT.StrPath, src_path: DT.StrPath,
                    description: str="", icon_path: Optional[Tuple[DT.StrPath, int]]=None,
                    working_dir: Optional[DT.StrPath]=None):
        """
        Helper method to create a link.
        Based on https://stackoverflow.com/q/22647661 C++ example
        """
        self.SetPath(str(src_path))
        if working_dir:
            self.SetWorkingDirectory(str(working_dir))
        if description:
            self.SetDescription(description)
        if icon_path:
            self.SetIconLocation(str(icon_path[0]), icon_path[1])
        ppf = self.query_interface(IPersistFile)
        ppf.Save(str(link_path), False)
