using System;
using System.Runtime.InteropServices;

[ComVisible(true)]
[Guid(Contract.Constants.ServerInterface)]
public interface IServer
{
    string Name { get; set; }

    string MethodWithArrayArgs(object[] arg1, object[] arg2);
}
