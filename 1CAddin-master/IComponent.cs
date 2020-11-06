using System.Runtime.InteropServices;
using Impinj.OctaneSdk;

namespace System1C.AddIn
{
    /// <summary>»нтерфейс с объ€влени€ми пользовательских методов и свойств компоненты</summary>
    /// 
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IComponent
    {
        [RussianName("ћетодЌа–усскомязыке")]
        int Procedure(int parameter);

        string Test();
        void TestEvent();

        void Connect(string ip);

        void Disconnect();

        bool IsConnected();

        void StartRead();

        void StopRead();
    }
}