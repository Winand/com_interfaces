from ctypes.wintypes import HWND, INT, ULARGE_INTEGER
from typing import Union as U

from com_interfaces import IUnknown, interface, method, Guid


class DT:
    "Additional data types"
    TBPFLAG = U[INT, int]
    HWND = U[HWND, int]
    ULONGLONG = U[ULARGE_INTEGER, int]


class TBPF:  # TBPFLAG
    NOPROGRESS = 0
    INDETERMINATE = 1
    NORMAL = 2  # green
    ERROR = 4  # red
    PAUSED = 8  # yellow


@interface
class ITaskBarList3(IUnknown):
    """
    Exposes methods that control the taskbar

    https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-itaskbarlist3
    ITaskBarList https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14087
    ITaskBarList2 https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14205
    ITaskBarList3 https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.16299.0/um/ShObjIdl_core.h#L14382
    """
    clsid = CLSID_TaskbarList = Guid("{56FDF344-FD6D-11d0-958A-006097C9A090}")
    iid = IID_ITaskbarList3 = Guid("{EA1AFB91-9E28-4B86-90E9-9E9F8A5EEFAF}")

    @method(index=9)
    def SetProgressValue(self, hwnd: DT.HWND, ullCompleted: DT.ULONGLONG,
                         ullTotal: DT.ULONGLONG):
        """
        Displays or updates a progress bar hosted in a taskbar button to show
        the specific percentage completed of the full operation.
        """

    @method(index=10)
    def SetProgressState(self, hwnd: DT.HWND, tbpFlags: DT.TBPFLAG):
        """
        Sets the type and state of the progress indicator displayed on
        a taskbar button.
        """
