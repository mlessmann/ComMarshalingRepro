# Repro of a COM marshalling error between Python and .NET 8

This repo contains a repro of a COM marshalling error that exists when a Python client communicates with a .NET 8 COM server. Specifically, marshalling of the value `null` in object arrays (SAFEARRAY(VARIANT)) used in method parameters as well as string (BSTR) properties causes an exception on the client side:

`(-2147024809, 'The parameter is incorrect.', (0, None, 'Specified OLE variant is invalid.', None, 0, -2147024809), None)`

or

`(-2147467262, 'No such interface supported', (0, 'System.Private.CoreLib', 'OleAut reported a type mismatch.', None, 0, -2147467262), None)`

If the COM server is executed on .NET Framework 4.8 no exceptions occurr.

This repo is based on the [OutOfProcCOM sample](https://github.com/dotnet/samples/tree/main/core/extensions/OutOfProcCOM) from the dotnet/samples repo and was modified to target both .NET Framework 4.8 and .NET 8.

# Preconditions

- Visual Studio 2022
- .NET 8 SDK
- Python 3.11 (probably also works with other Python versions)
- pywin32 installed: `pip install pywin32`

# How to reproduce

1. Compile the solution OutOfProcCOM/OutOfProcCOM.sln
2. Run elevated `OutOfProcCOM\ExeServer\bin\Debug\net8.0-windows\ExeServer.exe /regserver`
3. Run `OutOfProcCOM\ExeServer\bin\Debug\net8.0-windows\ExeServer.exe`
4. Run `Client.py` -> Exceptions

The projects also target .NET Framework 4.8. If you execute that ExeServer.exe, the Python exceptions don't occurr.
