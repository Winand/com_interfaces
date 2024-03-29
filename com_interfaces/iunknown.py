"""
Module provides functionality for instantiating and accessing Windows COM
objects through interfaces.

A COM interface is described using class with methods. All COM interfaces are
inherited from IUnknown interface.
An interface has @interface decorator and each COM method declaration has
@method(index=...) decorator where `index` is an index of that method in COM
interface.
E.g. IUnknown::QueryInterface has index 0, AddRef - index 1, Release - index 2.

@method saves index and arguments in a __com_func__ variable of a method. Then
@interface collects all methods containing that variable into __func_table__
dict of a class. At runtime IUnknown.__init__ creates an instance of a object
(if object pointer is not passed in arguments) and replaces all declarations
with generated methods which call methods of that object. So if __init__ is
overriden then super().__init__ should be called.

Example of a COM method declaration:
    @interface
    class IExample(IUnknown):
        @method(index=5)
        def Load(self, pszFileName: DT.LPCOLESTR, dwMode: DT.DWORD):
            ...
Argument type hint may be a Union. Only the first type is used for COM method
initialization. So you can use Union[c_wchar_p, str] to pass string w/o linting
error. Optionally you may raise NotImplementedError in COM method declaration
so you never call those stub methods in runtime accidentally, e.g. in case of
not initializing a method (missing @method decorator, IUnknown.__init__ not
called etc.).

Alternatively COM methods can be described in __methods__ dict of a class:
    @interface
    class IExample(IUnknown):
        __methods__ = {
            "Method1": {'index': 1, 'args': {"hwnd": DT.HWND}},
            "Method2": {'index': 5, 'args': (HWND, INT)},
            "Method3": {'index': 6},
        }
But this is not recommended because those methods are not recognized by linter.
"""

import logging
from ctypes import (HRESULT, POINTER, WINFUNCTYPE, Structure, byref, c_short,
                    c_ubyte, c_uint, c_void_p, cast, oledll)
from dataclasses import dataclass, fields
from types import FunctionType, new_class
from typing import Optional, Type, TypeVar
from typing import Union as U
from typing import get_args, get_type_hints

T = TypeVar('T')

ole32 = oledll.ole32
CLSCTX_INPROC_SERVER = 0x1

# Access COM methods from Python https://stackoverflow.com/q/48986244
ole32.CoInitialize(None)  # instead of `import pythoncom`


# FIXME: if @structure is used linter doesn't understand byref(self)
class Guid(Structure):
    """
    A GUID identifies an object such as a COM interfaces, or a COM class object
    https://learn.microsoft.com/en-us/windows/win32/api/guiddef/ns-guiddef-guid
    """
    _fields_ = [("Data1", c_uint),
                ("Data2", c_short),
                ("Data3", c_short),
                ("Data4", c_ubyte*8)]

    def __init__(self, name: str):
        ole32.CLSIDFromString(name, byref(self))


def structure(cls: Type[T]) -> Type[T]:
    "dataclass-like class to ctypes Structure conversion"
    if len(cls.__bases__) != 1:
        raise TypeError('Multiple inheritance is not supported')
    cls_struct = dataclass(new_class(
        cls.__name__, (Structure,),
        exec_body=lambda ns: ns.update(cls.__dict__)
    ))
    setattr(cls_struct, '_fields_', [
        (i.name, (get_args(i.type) or [i.type])[0]) for i in fields(cls_struct)
    ])
    # (!) dataclass is used for printing contents, but dataclass initialization
    # is not needed, replace it with the original __init__.
    # FIXME: why init=False fails
    setattr(cls_struct, '__init__', cls.__init__)
    return cls_struct


def interface(iid: Optional[U[Type[T], str]] = None,
              clsid: Optional[str] = None) -> Type[T]:
    """
    @interface class decorator collects all methods containing __com_func__
    variable into __func_table__ dict of a class.
    """
    def interface_(cls: Type[T]):
        if iid is not cls:  # if @interface is used w/ brackets
            if str and isinstance(iid, str):
                setattr(cls, 'iid', Guid(iid))
            if clsid:
                setattr(cls, 'clsid', Guid(clsid))
        # if "clsid" not in Cls.__dict__ or "iid" not in Cls.__dict__:
        #     raise ValueError(f"{Cls.__name__}: clsid / iid class variables not found")
        if len(cls.__bases__) != 1:
            # https://stackoverflow.com/questions/70222391/com-multiple-interface-inheritance
            raise TypeError('Multiple inheritance is not supported')
        # if Cls.__bases__[0] is object:
        #     if Cls.__name__ != 'IUnknown':
        #         logging.warning(f"COM interfaces should be derived from IUnknown, not {Cls.__name__}")
        __func_table__ = getattr(cls.__bases__[0], '__func_table__', {}).copy()
        for member_name, member in cls.__dict__.items():
            if not isinstance(member, FunctionType):
                continue
            __com_func__ = getattr(member, '__com_func__', None)
            if not __com_func__:
                continue
            __func_table__[member_name] = __com_func__
        __methods__ = cls.__dict__.get('__methods__')
        if isinstance(__methods__, dict):
            # Collect COM methods from __methods__ dict:
            # __methods__ = {
            #     "Method1": {'index': 1, 'args': {"hwnd": DT.HWND}},
            #     "Method2": {'index': 5, 'args': (HWND, INT)},
            #     "Method3": {'index': 6},
            # }
            for member_name, info in __methods__.items():
                if member_name in __func_table__:
                    logging.warning("Overriding existing method %s.%s",
                                    cls.__name__, member_name)
                args = info.get('args', ())
                if isinstance(args, dict):
                    args = tuple(args.values())
                __func_table__[member_name] = {
                    'index': info['index'],
                    'args': WINFUNCTYPE(HRESULT, c_void_p,
                        *((get_args(i) or [i])[0] for i in args)
                    )
                }
        setattr(cls, '__func_table__', __func_table__)
        return cls
    if isinstance(iid, (str, type(None))):
        # If @interface(..) is used w/ brackets linter generates type error
        return interface_  # type: ignore
    # If @interface is used w/o brackets a class is passed to `iid` arg
    return interface_(iid)


def method(index):
    "Saves index and arguments in a __com_func__ variable of a method"
    # https://stackoverflow.com/a/2367605
    def func_decorator(func):
        type_hints = get_type_hints(func)
        # Type of return value is not used.
        # Return type is HRESULT https://stackoverflow.com/a/20733034
        type_hints.pop('return', None)
        func.__com_func__ = {
            'index': index,
            'args': WINFUNCTYPE(HRESULT, c_void_p,
                *((get_args(i) or [i])[0] for i in type_hints.values())
            )
        }
        return func
    return func_decorator


def create_instance(clsid: Guid, iid: Guid):
    """
    Helper method for CoCreateInstance. Creates and default-initializes
    a single object of the class associated with a specified CLSID.
    """
    ptr = c_void_p()
    ole32.CoCreateInstance(byref(clsid), 0, CLSCTX_INPROC_SERVER,
                           byref(iid), byref(ptr))
    return ptr


CArgObject = type(byref(c_void_p()))


class DT:
    "Additional data types for type hints"
    REFIID = U[POINTER(Guid), CArgObject]
    void_pp = U[POINTER(c_void_p), CArgObject]


# pylint: disable=invalid-name
@interface
class IUnknown:
    """
    The IUnknown interface enables clients to retrieve pointers to other
    interfaces on a given object through the QueryInterface method, and to
    manage the existence of the object through the AddRef and Release methods.
    All other COM interfaces are inherited, directly or indirectly, from
    IUnknown.

    IUnknown https://learn.microsoft.com/en-us/windows/win32/api/unknwn/nn-unknwn-iunknown  # noqa
    IUnknown (DCOM) https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-dcom/2b4db106-fb79-4a67-b45f-63654f19c54c
    IUnknown (source) https://github.com/tpn/winsdk-10/blob/9b69fd26ac0c7d0b83d378dba01080e93349c2ed/Include/10.0.14393.0/um/Unknwn.h#L108
    """
    clsid: Guid
    iid: Guid = Guid("{00000000-0000-0000-C000-000000000046}")
    __func_table__ = {}
    T = TypeVar('T', bound="IUnknown")

    def __init__(self, ptr: Optional[c_void_p] = None):
        "Creates an instance and generates methods"
        self.ptr = ptr or create_instance(self.clsid, self.iid)
        # Access COM methods from Python https://stackoverflow.com/a/49053176
        # ctypes + COM access https://stackoverflow.com/a/12638860
        vtable = cast(self.ptr, POINTER(c_void_p))
        wk = c_void_p(vtable[0])
        functions = cast(wk, POINTER(c_void_p))  # method list
        for func_name, __com_opts__ in self.__func_table__.items():
            # Variable in a loop https://www.coursera.org/learn/golang-webservices-1/discussions/threads/0i1G0HswEemBSQpvxxG8fA/replies/m_pdt1kPQqS6XbdZD6Kkiw  # noqa
            win_func = __com_opts__['args'](functions[__com_opts__['index']])
            setattr(self, func_name,
                lambda *args, f=win_func: f(self.ptr, *args)
            )

    def query_interface(self, IID: Type[T]) -> T:
        "Helper method for QueryInterface"
        ptr = c_void_p()
        self.QueryInterface(byref(IID.iid), byref(ptr))
        return IID(ptr)

    @method(index=0)
    def QueryInterface(self, riid: DT.REFIID, ppvObject: DT.void_pp) -> HRESULT:
        "Retrieves pointers to the supported interfaces on an object."
        raise NotImplementedError

    @method(index=1)
    def AddRef(self) -> HRESULT:
        """
        Increments the reference count for an interface pointer to a COM object
        """
        raise NotImplementedError

    @method(index=2)
    def Release(self) -> HRESULT:
        "Decrements the reference count for an interface on a COM object"
        raise NotImplementedError

    def __del__(self):
        if self.ptr:
            self.Release()

    def is_accessible(self):
        "Check pointer to an object"
        return bool(self.ptr)
