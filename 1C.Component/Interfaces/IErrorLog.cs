using System.Runtime.InteropServices;

namespace _1C.Component.Interfaces
{
    [Guid("3127CA40-446E-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IErrorLog
    {
        void AddError(string pszPropName, ref System.Runtime.InteropServices.ComTypes.EXCEPINFO pExepInfo);
    }
}
