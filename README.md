**com-interfaces** package provides functionality for declaring Windows COM interfaces,
instantiating COM objects and accessing them through interfaces.

A COM interface is described using class with methods. All COM interfaces are
inherited from `IUnknown` interface.
An interface has `@interface` decorator and each COM method declaration has
`@method(index=...)` decorator where `index` is an index of that method in COM interface.
E.g. IUnknown::QueryInterface has index 0, AddRef - index 1, Release - index 2.

`@method` saves index and arguments in a `__com_func__` variable of a method. Then
`@interface` collects all methods containing that variable into `__func_table__` dict
of a class. At runtime `IUnknown.__init__` creates an instance of a object
(if object pointer is not passed in arguments) and replaces all declarations
with generated methods which call methods of that object. So if `__init__` is
overriden then `super().__init__` should be called.

Example of a COM method declaration:
```
@interface
class IExample(IUnknown):
    @method(index=5)
    def Load(self, pszFileName: DT.LPCOLESTR, dwMode: DT.DWORD):
        ...
```
Argument type hint may be a `Union`. Only the first type is used for COM method
initialization. So you can use `Union[c_wchar_p, str]` to pass string w/o linting error.
Optionally you may raise `NotImplementedError` in COM method declaration so you never call
those stub methods in runtime accidentally, e.g. in case of not initializing a method
(missing `@method` decorator, `IUnknown.__init__` not called etc.).

Alternatively COM methods can be described in `__methods__` dict of a class:
```
@interface
class IExample(IUnknown):
    __methods__ = {
        "Method1": {'index': 1, 'args': {"hwnd": DT.HWND}},
        "Method2": {'index': 5, 'args': (HWND, INT)},
        "Method3": {'index': 6},
    }
```
But this is not recommended because those methods are not recognized by linter.

# IID and CLSID in Component Object Model ([COM](https://learn.microsoft.com/en-us/windows/win32/com/component-object-model--com--portal))
IID is an interface identificator and [CLSID](https://learn.microsoft.com/en-us/windows/win32/com/clsid-key-hklm)
identifies COM class objects. IID should be specified for any interface.
CLSID is needed to instantiate an object which implements that interface.

IID and optionally CLSID are specified as str arguments of `@interface` decorator:
```
@interface("{000214F9-0000-0000-C000-000000000046}", clsid="{00021401-0000-0000-C000-000000000046}")
class IShellLink(IUnknown):
    ...
```
Alternatively they can be specifid as `iid` and `clsid` variables of type `Guid` in class body:
```
@interface
class IPersistFile(IUnknown):
    clsid = Guid("{00021401-0000-0000-C000-000000000046}")
    iid = Guid("{0000010b-0000-0000-C000-000000000046}")
```

# IUnknown
In addition to `QueryInterface`, `AddRef` and `Release` methods `IUnknown` class
has `query_interface(IID: type[T]) -> T` helper method, which returns an instance
of a queried interface class.

# Examples
`com_interfaces.interfaces` contains several examples of COM interfaces inherited
from `IUnknown`: `IPersistFile`, `IShellLink`, `ITaskBarList3`.
