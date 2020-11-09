using System;
using System.Collections.Concurrent;
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
        private ConcurrentDictionary<string, int> _tags;
        private Settings _setting;

        public Settings Setting
        {
            get { return _setting;}
            set { _setting = value; }
        }
        public ConcurrentDictionary<string, int> Tags
        {
            get { return _tags; }

            set { _tags = value; }
        }
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
            return "Компонента подключена успешно";
        }

        public void TestEvent()
        {
            for (int i = 0; i < 50; i++)
            {
                AddIn.AsyncEvent.ExternalEvent("Test", "AAA", i.ToString());
            }
        }

        #endregion

        #region Reader methods

        private void ConnectReader()
        {
            Disconnect();

            try
            {
                Reader.Connect();
                Setting = Reader.QueryDefaultSettings();
                SetReaderSettings();
                SetAntennaSettings();
                Reader.ApplySettings(Setting);
            }
            catch (OctaneSdkException ee)
            {
                Console.WriteLine("Octane SDK exception: Reader " + ee.Message, "error");
            }
            catch (Exception ee)
            {
                // Handle other .NET errors.
                Console.WriteLine("Exception : Reader " + ee.Message, "error");
            }
        }

        private void SetReaderSettings()
        {
            var settings = Setting;
            try
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
                settings.ReaderMode = ReaderMode.MaxThroughput; //.AutoSetDenseReader;
                settings.SearchMode = SearchMode.DualTarget; //.DualTarget;
                settings.Session = 1;
                settings.TagPopulationEstimate = Convert.ToUInt16(200);
                settings.Report.Mode = ReportMode.Individual;
            }
            catch (OctaneSdkException ee)
            {
                Console.WriteLine("Octane SDK exception: Reader " + ee.Message, "error");
            }
            catch (Exception ee)
            {
                // Handle other .NET errors.
                Console.WriteLine("Exception : Reader " + ee.Message, "error");
            }
        }

        private void SetAntennaSettings()
        {
            var settings = Setting;
            settings.Antennas.DisableAll();

            for (int i = 1; i <= settings.Antennas.Length; i++)
            {
                var antenna = settings.Antennas.GetAntenna((ushort)i);

                antenna.IsEnabled = true;
                //antenna.TxPowerInDbm = Convert.ToDouble(15);
                //antenna.RxSensitivityInDbm = Convert.ToDouble(-80);
                antenna.MaxRxSensitivity = true;
                antenna.MaxTxPower = true;
            }
        }

        public void SetReader(string ip)
        {
            if (string.IsNullOrWhiteSpace(ip)) return;
            Reader = new ImpinjReader(ip, "1");

            var x = 0;
            AddIn.AsyncEvent.CleanBuffer();
            AddIn.AsyncEvent.GetEventBufferDepth(ref x);

            if (x < 100000)
            {
                AddIn.AsyncEvent.SetEventBufferDepth(100000);
            }
        }

        public void Connect()
        {
           if(Reader == null) return;

           if (IsConnected())
           {
               Reader.Stop();
           }
           else
           {
               ConnectReader();
           }
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
            Connect();

            Tags = new ConcurrentDictionary<string, int>();
            Reader.TagsReported += DisplayTag;
            Reader.Start();
        }

        public void StopRead()
        {
            if (Reader == null || !IsConnected()) return;

            Reader.TagsReported -= DisplayTag;
            Disconnect();
        }

        public int GetTagsCount()
        {
            return Tags.Count;
        }

        private void DisplayTag(ImpinjReader reader, TagReport report)
        {
            foreach (Tag tag in report)
            {
                if (!Tags.TryGetValue(tag.Epc.ToString(), out int val))
                {
                    if (Tags.TryAdd(tag.Epc.ToString(), tag.AntennaPortNumber))
                    {
                        AddIn.AsyncEvent.ExternalEvent(addInName, "TagRead", $"{tag.Epc.ToString()} : {tag.AntennaPortNumber}");
                    }
                }
                else
                {
                    Tags.TryUpdate(tag.Epc.ToString(), tag.AntennaPortNumber, val);
                }

            }
        }

        #endregion
    }
}
