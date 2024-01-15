using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace OutOfProcCOM
{
    [ComVisible(true)]
    [Guid(Contract.Constants.ServerClass)]
    [ComDefaultInterface(typeof(IServer))]
    public sealed class ExeServer : IServer
    {
        public string Name { get; set; }

        public string MethodWithArrayArgs(object[] arg1, object[] arg2)
        {
            string FormatArray(object[] array)
            {
                if (array == null)
                {
                    return "<null>";
                }

                var itemStrings = array.Select(i => i.ToString());
                var result = string.Join(", ", itemStrings);
                return $"[{result}]";
            }

            return $"arg1: {FormatArray(arg1)}, arg2: {FormatArray(arg2)}";
        }
    }
}
