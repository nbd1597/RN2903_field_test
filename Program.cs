using ClosedXML.Excel;
using System;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Threading;

namespace COMPortReader
{
    public class SerialPortReader
    {
        private static bool _continue;
        private static SerialPort _serialPort;
        private static string ExcelName;
        private static string SaveDirectory;

        public static void Main()
        {
            string message;

            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;
            
            

            // Create a new SerialPort object with default settings.
            _serialPort = new SerialPort();

            // Allow the user to set the appropriate properties.
            _serialPort.PortName = SetPortName(_serialPort.PortName);
            _serialPort.BaudRate = 57600;

            // Set the read/write timeouts
            _serialPort.ReadTimeout = 500;
            _serialPort.WriteTimeout = 500;

            // Set excel name and save location
            SaveDirectory = SetDirectory(Directory.GetDirectoryRoot(Directory.GetCurrentDirectory()));
            ExcelName = SetExcelName("test.xlsx");

            _serialPort.Open();
            
            
            Thread readThread;
            Console.Write("mode TX or RX: ");
            string mode = Console.ReadLine();
            
            if (mode == "RX" || mode =="rx")  
            {
                readThread = new Thread(RX);
                Console.Write("RX mode\r\n");
            }
            else if(mode == "TX" || mode =="tx") 
            {
                readThread = new Thread(TX);
                Console.Write("TX mode\r\n");
            }
            else   
            {
                Console.Write("invalid\r\n");
                return;
            } 
            Lora_Radio Lora = new Lora_Radio();
            
            
            _serialPort.Write("sys reset\r\n");
            Thread.Sleep(2000);
            setup(Lora);
            _serialPort.Write("mac pause\r\n");
            Thread.Sleep(3500);
            // Console.WriteLine(_serialPort.ReadLine());
            Console.WriteLine("Type QUIT to exit");
            readThread.Start();
            
            _continue = true;
            
            
            while (_continue)
            {
                message = Console.ReadLine();

                if (stringComparer.Equals("quit", message))
                {
                    _continue = false;
                }
            }

            readThread.Join();
            _serialPort.Close();
            Console.WriteLine(Directory.GetCurrentDirectory());
        }
        public static void setup(Lora_Radio Lora)
        {
            Lora.ChangeDefault();
            _serialPort.Write(String.Format("radio set freq {0}\r\n", Lora.freq));
            Console.WriteLine("fre {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set sf {0}\r\n", Lora.SF));
            Console.WriteLine("sf {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set pwr {0}\r\n", Lora.pwr));
            Console.WriteLine("pwr {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set crc {0}\r\n", Lora.crc));
            Console.WriteLine("crc {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set iqi {0}\r\n", Lora.iqi));
            Console.WriteLine("iqi {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set cr {0}\r\n", Lora.cr));
            Console.WriteLine("cr {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set wdt {0}\r\n", Lora.wdt));
            Console.WriteLine("wdt {0}", _serialPort.ReadLine());
            _serialPort.Write(String.Format("radio set bw {0}\r\n", Lora.bw));
            Console.WriteLine("bw {0}", _serialPort.ReadLine());  
            _serialPort.Write(String.Format("radio set sync {0}\r\n", Lora.sync));
            Console.WriteLine("sync {0}", _serialPort.ReadLine());        


        }
        public class Lora_Radio
        {
            public int freq;
            public string SF;
            public int pwr;
            public string crc;
            public string iqi;
            public string cr;
            public int wdt;
            public int bw;
            public int sync;
            public Lora_Radio()
            {
                freq = 923300000;
                SF = "sf12";
                pwr = 2;
                bw = 125;
                cr = "4/5";
                iqi = "off";
                crc = "on";
                wdt = 15000;
                sync = 34;
            }
            public void ChangeDefault()
            {
                string temp = "";
                Console.WriteLine("freq (902000000 to 928000000): ");
                temp = Console.ReadLine();
                if(temp != "")  freq = int.Parse(temp);
                Console.WriteLine("bw: ");
                temp = Console.ReadLine();
                if(temp != "")  bw = int.Parse(temp);
                Console.WriteLine("power: ");
                temp = Console.ReadLine();
                if(temp != "")  pwr = int.Parse(temp);
                Console.WriteLine("sf: ");
                temp = Console.ReadLine();
                if(temp != "")  SF = temp;
                Console.WriteLine("coderate: ");
                temp = Console.ReadLine();
                if(temp != "")  cr = temp;


            }


        }

        public class DataPackage
        {
            public static int ReceiveCount = 0;
            public static DataTable dataTable = new DataTable();

            public static void ReceiptCount(string source)
            {
                if (!dataTable.Columns.Contains("Receipt count"))
                {
                    dataTable.Columns.Add("Receipt count", typeof(int));
                }
                if (source.Contains("Receive Finished"))
                {
                    ReceiveCount++;
                }

            }

            public static string getBetween(string strSource, string strStart, string strEnd)
            {
                const int kNotFound = -1;

                var startIdx = strSource.IndexOf(strStart);
                if (startIdx != kNotFound)
                {
                    startIdx += strStart.Length;
                    var endIdx = strSource.IndexOf(strEnd, startIdx);
                    if (endIdx > startIdx)
                    {
                        return strSource.Substring(startIdx, endIdx - startIdx);
                    }
                }
                return String.Empty;
            }
            public static string GetRX(string source, string col)
            {
                if (!dataTable.Columns.Contains(col))
                {
                    dataTable.Columns.Add(col, typeof(string));
                }
                return source.Substring(11);
            }

            public static string GetData(string source, string keyword1, string keyword2, string col)
            {
                if (!dataTable.Columns.Contains(col))
                {
                    dataTable.Columns.Add(col, typeof(string));
                }
                if (source.Contains(keyword1))
                {
                    string dt = getBetween(source, keyword1, keyword2);
                    //Int32.TryParse(dt, out int dtVal);
                    //DataRow row = dataTable.NewRow();
                    //row[col] = dtVal;
                    //dataTable.Rows.InsertAt(row, dataTable.Columns.IndexOf(col));
                    return dt;
                }else
                {
                    return null;
                }
            }

            public static void ExtractData(string source)
            {
                string[] data = new string[2];
                // data[0] = GetData(source, "RSSI:", "dBm", "RSSI (dBm)");
                // data[1]= GetData(source, "SNR:", "dB", "SNR (dB)");
                // data[2]= GetData(source, "Payload_size:", "bytes", "Payload Size (bytes)");
                // data[3] = GetData(source, "Payload_data:", ":End_payload_data", "Payload Data");
                // data[0] = GetData(source, "radio_rx ", "\r\n", "Payload_data");
                data[0] = GetRX(source, "Payload_Data");
                _serialPort.Write("radio get snr\r\n");
                // Thread.Sleep(2);
                string message = String.Format("snr {0}\r\n" ,_serialPort.ReadLine());
                Console.Write(message);
                data[1] = GetData(message, "snr", "\r\n", "SNR");
                DataRow row = dataTable.NewRow();
                for(int i = 0; i < 2; i++)
                {
                    row[i] = data[i];
                }
                dataTable.Rows.Add(row);
            }

                public static void ExportExcel()
            {
                DataRow countReceipt = dataTable.NewRow();
                countReceipt["Receipt count"] = ReceiveCount;
                dataTable.Rows.Add(countReceipt);

                IXLWorkbook wb = new XLWorkbook();
                wb.Worksheets.Add(dataTable, "result");
                wb.SaveAs(SaveDirectory + ExcelName);
            }
        }

        public static void RX()
        {
           
            _serialPort.Write("radio rx 0\r\n");
            
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);
                    if(String.Compare(message, 0, "radio_rx", 0, 8, true) == 0)
                    {
                        
                        // Thread.Sleep(1);
                        DataPackage.ExtractData(message);
                        DataPackage.ReceiptCount(message);
                        _serialPort.Write("radio rx 0\r\n");
                        
                    }
                    if(String.Compare(message, 0, "radio_err", 0, 9, true) == 0)
                    {
                        _serialPort.Write("radio rx 0\r\n");
                    }
                    
                    
                    /*Dp.GetData(message, "Payload number", ";", "Payload number");
                    Dp.GetData(message, "RSSI", "dBm", "RSSI (dBm)");
                    Dp.GetData(message, "SNR", "dB", "SNR (dB)");
                    Dp.GetData(message, "Payload size", "bytes", "Payload Size (bytes)");
                    Dp.GetData(message, "Payload data", ";", "Payload Data");*/
                    
                    // _serialPort.Write("radio rx 0\r\n");
                }
                catch (TimeoutException) { }
            }
            DataPackage.ExportExcel();
        }
        public static void TX()
        {
            int tx_count = 0;
            string packet;
            try
            {
                while (_continue)
                {
                    if (tx_count > 9)   packet = String.Format("radio tx abc{0}\r\n", tx_count);
                    else    packet = String.Format("radio tx abc0{0}\r\n", tx_count);
                    _serialPort.Write(packet);
                    Console.WriteLine("TX packet: {0}", packet);
                    string message = _serialPort.ReadLine();
                    Console.WriteLine("count {0} {1}",tx_count,message);

                    
                    // message = _serialPort.ReadLine();
                    // Console.WriteLine(message);
                    tx_count++;
                    if(tx_count >= 100) break;
                    // message = _serialPort.ReadLine();
                    // Console.WriteLine(message);
                    // while(String.Compare(message, 0, "radio_tx_ok", 0, 11, true) != 0)
                    // {
                    //     message = _serialPort.ReadLine();
                    //     Console.WriteLine("???:{0}",message);
                    //     continue;
                    // }
                    Thread.Sleep(1500);
                    
                    
                }
            }
            catch (TimeoutException)
            {
                
            }
        }

        public static string SetPortName(string defaultPortName)
        {
            string portName;

            Console.WriteLine("Available Ports:");
            foreach (string s in SerialPort.GetPortNames())
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("COM port({0}): ", defaultPortName);
            portName = Console.ReadLine();

            if (portName == "")
            {
                portName = defaultPortName;
            }
            return portName;
        }

        public static string SetExcelName(string defaultExcelName)
        {
            string ExcelName;

            Console.Write("Excelname(DefaultName: {0}): ", defaultExcelName);
            ExcelName = Console.ReadLine() + ".xlsx";

            if (ExcelName == "")
            {
                ExcelName = defaultExcelName;
            }
            return ExcelName;
        }

        public static string SetDirectory(string defaultDirectory)
        {
            string Dirlink;

            Console.Write("Set directory(Default: {0}): ", defaultDirectory);
            Dirlink = Console.ReadLine() + "/";

            if (Dirlink == "")
            {
                Dirlink = defaultDirectory;
            }
            return Dirlink;
        }
    }
}