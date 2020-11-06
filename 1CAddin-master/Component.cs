using System;
using System.Runtime.InteropServices;

namespace System1C.AddIn
{
    /// <summary>Класс, реализующий пользовательские методы компоненты</summary>
    [ProgId("AddIn.1C_Component")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Component : AddIn, IComponent
    {
        public int Procedure(int parameter)
        {
            return parameter + 100;
        }
    }
}