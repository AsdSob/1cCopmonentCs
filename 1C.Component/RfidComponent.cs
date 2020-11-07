using System;
using System.Runtime.InteropServices;
using System.Text;
using _1C.Component.Interfaces;
using Impinj.OctaneSdk;


namespace _1C.Component
{
    [ComVisible(true), Guid("29C03F10-7459-4368-BD9D-CA26B386C142"), ProgId("AddIn.RfidComponent")]
    public class RfidComponent : IInitDone, ILanguageExtender
    {
        public RfidComponent() { }


        #region Properties

        private const string addInName = "RfidComponent";

        private StringBuilder _lastError;
        private ImpinjReader _reader;

        public StringBuilder LastError
        {
            get { return _lastError; }

            set { _lastError = value; }
        }
        public ImpinjReader Reader
        {
            get { return _reader; }

            set { _reader = value; }
        }

        #endregion

        #region IInitDone implementation
        public void Init([MarshalAs(UnmanagedType.IDispatch)] object pConnection)
        {
            AddIn.load(pConnection);
        }

        public void Done() { }

        public void setMemManager(ref object mem) { }


        public void GetInfo(ref object pInfo)
        {
            ((System.Array)pInfo).SetValue("2000", 0);
        }
        #endregion

        #region ILanguageExtender implementation

        public void RegisterExtensionAs(ref string bstrExtensionName)
        {
            bstrExtensionName = addInName;
        }

        #region Methods

        public void CallAsFunc(int lMethodNum, ref object pvarRetValue, [MarshalAs(UnmanagedType.SafeArray)] ref Array paParam)
        {
            //Здесь внешняя компонента выполняет код Функций
        }

        public void CallAsProc(int lMethodNum, [MarshalAs(UnmanagedType.SafeArray)] ref Array paParams)
        {
            //Здесь внешняя компонента выполняет код процедур
        }

        public void GetNMethods(ref int plMethods)
        {
            //Здесь 1С получает количество доступных из ВК методов
        }

        public void FindMethod(string bstrMethodName, ref int plMethodNum)
        {
            //Здесь 1С получает числовой идентификатор метода (процедуры или функции) по имени (названию) процедуры или функции
        }

        public void GetMethodName(int lMethodNum, int lMethodAlias, ref string pbstrMethodName)
        {
            //Здесь 1С (теоретически) получает имя метода по его идентификатору. lMethodAlias - номер синонима.
        }

        public void GetNParams(int lMethodNum, ref int plParams)
        {
            //Здесь 1С получает количество параметров у метода (процедуры или функции)
        }

        public void GetParamDefValue(int lMethodNum, int lParamNum, ref object pvarParamDefValue)
        {
            //Здесь 1С получает значения параметров процедуры или функции по умолчанию
        }

        public void HasRetVal(int lMethodNum, ref bool pboolRetValue)
        {
            //Здесь 1С узнает, возвращает ли метод значение (т.е. является процедурой или функцией)
        }
        #endregion

        #region Properties

        public void FindProp(string bstrPropName, ref int plPropNum)
        {
            //Здесь 1С ищет числовой идентификатор свойства по его текстовому имени
        }

        public void GetNProps(ref int plProps)
        {
            //Здесь 1С получает количество доступных из ВК свойств
        }

        public void GetPropName(int lPropNum, int lPropAlias, ref string pbstrPropName)
        {
            //Здесь 1С узнает имя свойства по его идентификатору. lPropAlias - номер псевдонима
        }

        public void GetPropVal(int lPropNum, ref object pvarPropVal)
        {
            //Здесь 1С узнает значения свойств 
        }

        public void SetPropVal(int lPropNum, ref object varPropVal)
        {
            //Здесь 1С изменяет значения свойств
        }

        public void IsPropReadable(int lPropNum, ref bool pboolPropRead)
        {
            //Здесь 1С узнает, какие свойства доступны для чтения
            pboolPropRead = true; // Все свойства доступны для чтения
        }

        public void IsPropWritable(int lPropNum, ref bool pboolPropWrite)
        {
            //Здесь 1С узнает, какие свойства доступны для записи
            pboolPropWrite = true; // Все свойства доступны для записи
        }

        #endregion
        #endregion

        #region Methods and events helpers

        public string Test()
        {
            var x = 4;
            AddIn.AsyncEvent.SetEventBufferDepth(1000);

            AddIn.AsyncEvent.GetEventBufferDepth(ref x);

            Reader = new ImpinjReader("192.168.1.69", "1");
            return "buffer = " + x;
        }

        public void TestEvent()
        {
            for (int i = 0; i < 500; i++)
            {
                AddIn.AsyncEvent.ExternalEvent("Asterisk.AddIn", "AAA", i.ToString());
            }
        }

        #endregion

        #region Reader methods

        public void ConnectReader()
        {
            //readers.Add(new ImpinjReader(textBox1.Text, "Reader #1"));

            if (Reader.IsConnected)
            {
                Reader.Disconnect();
            }

            try
            {
                Reader.Connect();

                Settings settings = Reader.QueryDefaultSettings();
                settings.Report.IncludeAntennaPortNumber = true;
                settings.Report.IncludePhaseAngle = true;
                settings.Report.IncludeChannel = true;
                settings.Report.IncludeDopplerFrequency = true;
                settings.Report.IncludeFastId = true;
                settings.Report.IncludeFirstSeenTime = true;
                settings.Report.IncludeLastSeenTime = true;
                settings.Report.IncludePeakRssi = true;
                settings.Report.IncludeSeenCount = true;
                settings.Report.IncludePcBits = true;

                settings.ReaderMode = ReaderMode.MaxThroughput; //.AutoSetDenseReader;
                settings.SearchMode = SearchMode.DualTarget; //.DualTarget;
                settings.Session = 1;
                settings.TagPopulationEstimate = Convert.ToUInt16(200);

                SetAntennaSettings(settings);
                settings.Report.Mode = ReportMode.Individual;

                Reader.ApplySettings(settings);
            }
            catch (OctaneSdkException ee)
            {
                Console.WriteLine("Octane SDK exception: Reader #1" + ee.Message, "error");
            }
            catch (Exception ee)
            {
                // Handle other .NET errors.
                Console.WriteLine("Exception : Reader #1" + ee.Message, "error");
            }
        }

        private void SetSettings(Settings settings)
        {
            settings.Report.IncludeAntennaPortNumber = true;
            settings.Report.IncludePhaseAngle = true;
            settings.Report.IncludeChannel = true;
            settings.Report.IncludeDopplerFrequency = true;
            settings.Report.IncludeFastId = true;
            settings.Report.IncludeFirstSeenTime = true;
            settings.Report.IncludeLastSeenTime = true;
            settings.Report.IncludePeakRssi = true;
            settings.Report.IncludeSeenCount = true;
            settings.Report.IncludePcBits = true;
            settings.Report.IncludeSeenCount = true;

            //ReaderMode.AutoSetDenseReaderDeepScan | Rx = -70 | Tx = 15/20
            //ReaderMode.MaxThrouput | Rx = -80 | Tx = 15

            settings.ReaderMode = ReaderMode.AutoSetDenseReaderDeepScan;//.AutoSetDenseReader;
            settings.SearchMode = SearchMode.DualTarget;//.DualTarget;
            settings.Session = 1;
            settings.TagPopulationEstimate = 200;
            settings.Report.Mode = ReportMode.Individual;
        }

        private void SetAntennaSettings(Settings settings)
        {
            settings.Antennas.DisableAll();

            for (int i = 1; i <= settings.Antennas.Length; i++)
            {
                var antenna = settings.Antennas.GetAntenna((ushort)i);

                antenna.IsEnabled = true;
                antenna.TxPowerInDbm = Convert.ToDouble(15);
                antenna.RxSensitivityInDbm = Convert.ToDouble(-80);
            }
        }

        public void Connect(string ip)
        {
            if (Reader != null && Reader.IsConnected)
            {
                Reader.Stop();
                Reader.Disconnect();
            }

            Reader = new ImpinjReader(ip, "1");
        }

        public void Disconnect()
        {
            if (Reader == null || !Reader.IsConnected) return;

            Reader.Stop();
            Reader.Disconnect();
        }

        public bool IsConnected()
        {
            return Reader.IsConnected;
        }

        public void StartRead()
        {
            if (Reader == null) return;

            ConnectReader();

            if (!Reader.IsConnected) return;
            Reader.Stop();

            Reader.TagsReported += DisplayTag;
            Reader.Start();
        }

        public void StopRead()
        {
            if (Reader == null) return;

            if(!Reader.IsConnected) return;
            
            Reader.Stop();
            Reader.TagsReported -= DisplayTag;
            Disconnect();
        }

        private void DisplayTag(ImpinjReader reader, TagReport report)
        {
            foreach (Tag tag in report)
            {
                //TODO: добавить метод AsyncEvent.ExternalEvent
            }
        }

        #endregion
    }
}
