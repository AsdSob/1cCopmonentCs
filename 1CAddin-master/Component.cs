using System;
using System.Runtime.InteropServices;
using Impinj.OctaneSdk;

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

        public string Test()
        {
            return "вороде работает";
        }

        public void TestEvent()
        {
            for (int i = 0; i < 5; i++)
            {
                AsyncEvent.ExternalEvent("111", "2222", "3333");
            }
        }



#region Reader methods

        public ImpinjReader Reader { get; set; }

        public void Connect(string ip)
        {
            //readers.Add(new ImpinjReader(textBox1.Text, "Reader #1"));
            Reader = new ImpinjReader(ip, "1");

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

        public void Disconnect()
        {
            if (Reader == null || !Reader.IsConnected) return;

            StopRead();
            Reader.Disconnect();
        }

        public bool IsConnected()
        {
            return Reader.IsConnected;
        }

        public void StartRead()
        {
            if (Reader == null || !Reader.IsConnected) return;

            Reader.TagsReported += DisplayTag;
            Reader.Start();
        }

        public void StopRead()
        {
            if (Reader == null || !Reader.IsConnected) return;

            Reader.Stop();
            Reader.TagsReported -= DisplayTag;
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