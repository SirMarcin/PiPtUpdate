using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Configuration;
using OSIsoft.AF.Asset;
using OSIsoft.AF.Search;
using OSIsoft.AF;
using OSIsoft.AF.Time;
using OSIsoft.AF.Data;
using OSIsoft.AF.PI;
using System.Net.Mail;
using System.Net;
using System.IO;



namespace UpdatePiPoints2
{


    class Program
    {
        static void Main(string[] args)
        {
            #region Logowanie
            string today = DateTime.Now.ToString("dd_MM_yyyy");
            string fileName = "C:\\Users\\cbuczynm\\Maintanance AF i ARC\\LogiZAktualizacjiPIPoints\\updatePiLog_" + today + ".txt";
            FileStream filestream = new FileStream(fileName, FileMode.Create);
            var streamwriter = new StreamWriter(filestream);
            streamwriter.AutoFlush = true;
            Console.SetOut(streamwriter);
            Console.SetError(streamwriter);
            #endregion

            #region PiConnection
            PIServers piServers = new PIServers();
            PIServer piServer = piServers["cp-pisrv1"];
            piServer.Connect();
            Console.WriteLine("Połączony z {0}", piServer);
            #endregion

            #region Ładowanie najnowszego i poprzedniego pliku interfejsu GIS
            var files = new DirectoryInfo("X:\\GIS_LOD_1\\archiwum\\vector").GetFiles("cieplo_wodo_rds_*.txt");
            //var files = new DirectoryInfo("C:\\Users\\cbuczynm\\Desktop").GetFiles("cieplo_wodo_rds_*.txt");
            Console.WriteLine("tusięwykładam");
            string latestFile = "";
            DateTime lastUpdated = DateTime.MinValue;
            foreach (var file in files)
            {
                if (file.LastWriteTime > lastUpdated)
                {
                    lastUpdated = file.LastWriteTime;
                    latestFile = file.FullName;
                }
            }
            Console.WriteLine("Ostatni plik interfejsu : {0}", latestFile);

            var logLatestFile = File.ReadAllLines(latestFile);
            var logLatestList = new List<string>(logLatestFile);
            logLatestList.Remove(logLatestList.First());
            for (int i = 0; i < logLatestList.Capacity - 1; i++)
            {
                logLatestList[i] = logLatestList[i].Replace("!Brak wezla Maximo", "");
                logLatestList[i] = logLatestList[i].Replace("!Brak odpowiedniego obiegu", "");
            }
            Console.WriteLine("Plik zawiera {0} rekordów", logLatestList.Capacity);


            string previousFile = "";
            lastUpdated = DateTime.MinValue;
            foreach (var file in files)
            {
                if (file.LastWriteTime > lastUpdated && file.FullName != latestFile)
                {
                    lastUpdated = file.LastWriteTime;
                    previousFile = file.FullName;
                }
            }
            Console.WriteLine("Poprzedni plik interfejsu : {0}", previousFile);
            var logPreviousFile = File.ReadAllLines(previousFile);
            var logPreviousList = new List<string>(logPreviousFile);
            logPreviousList.Remove(logPreviousList.First());
            for (int i = 0; i < logPreviousList.Capacity - 1; i++)
            {
                logPreviousList[i] = logPreviousList[i].Replace("!Brak wezla Maximo", "");
                logPreviousList[i] = logPreviousList[i].Replace("!Brak odpowiedniego obiegu", "");
            }
            Console.WriteLine("Plik zawiera {0} rekordów", logPreviousList.Capacity);

            #endregion

            #region Ładowanie listy kodów RDS PP istniejących w CM

            //Ciepłomierze
            var filesVectorMeter = new DirectoryInfo("X:\\RDS_PP\\VECTOR").GetFiles("rdspp_cieplomierze*.csv");

            string vectorMeterFile = "";
            lastUpdated = DateTime.MinValue;
            foreach (var file in filesVectorMeter)
            {
                if (file.LastWriteTime > lastUpdated && file.FullName != latestFile)
                {
                    lastUpdated = file.LastWriteTime;
                    vectorMeterFile = file.FullName;
                }
            }
            Console.WriteLine("Plik Vector Cieplomierzy : {0}", vectorMeterFile);
            var logvectorMeterFile = File.ReadAllLines(vectorMeterFile);
            var logvectorMeterList = new List<string>(logvectorMeterFile);
            logvectorMeterList.Remove(logvectorMeterList.First());
            
            Console.WriteLine("Plik zawiera {0} rekordów z kodami RDSPP Ciepłomierzy", logvectorMeterList.Capacity);

            //Wodomierze
            var filesVectorWodo = new DirectoryInfo("X:\\RDS_PP\\VECTOR").GetFiles("rdspp_wodomierze*.csv");

            string vectorWodoFile = "";
            lastUpdated = DateTime.MinValue;
            foreach (var file in filesVectorWodo)
            {
                if (file.LastWriteTime > lastUpdated && file.FullName != latestFile)
                {
                    lastUpdated = file.LastWriteTime;
                    vectorWodoFile = file.FullName;
                }
            }
            Console.WriteLine("Plik Vector Wodomierzy : {0}", vectorWodoFile);
            var logvectorWodoFile = File.ReadAllLines(vectorWodoFile);
            var logvectorWodoList = new List<string>(logvectorWodoFile);
            logvectorWodoList.Remove(logvectorMeterList.First());

            Console.WriteLine("Plik zawiera {0} rekordów z kodami RDSPP Wodomierzy", logvectorWodoList.Capacity);


            #endregion

            #region Parsowanie plików interfejsu GIS

            var listOfPreviousPoints = new List<Point>();

            foreach (var item in logPreviousList)
            {
                string[] parameters = item.Split('|');
                Point point = new Point();
                point.UPW = Convert.ToInt32(parameters[0]);
                point.UPW_Kombit = Convert.ToInt32(parameters[1]);
                point.Typ_przelicznika = parameters[2];
                point.WL_id = Convert.ToInt32(parameters[3]);
                point.Numer_przelicznika = parameters[4];
                point.Cieplomierz_RDS_PP = parameters[5];
                point.WWU1_numer = parameters[6];
                point.WWU1_RDS_PP = parameters[7];
                point.WWU2_numer = parameters[8];
                point.WWU2_RDS_PP = parameters[9];
                point.WCW_numer = parameters[10];
                point.WCW_RDS_PP = parameters[11];
                point.WCO_Numer = parameters[12];
                point.WCO_RDS_PP = parameters[13];
                point.adres_wf = parameters[14];
                point.wskaznik_kontrolny = parameters[15];

                listOfPreviousPoints.Add(point);
            }

            var listOfLatestPoints = new List<Point>();

            foreach (var item in logLatestList)
            {
                string[] parameters = item.Split('|');
                Point point = new Point();
                point.UPW = Convert.ToInt32(parameters[0]);
                point.UPW_Kombit = Convert.ToInt32(parameters[1]);
                point.Typ_przelicznika = parameters[2];
                point.WL_id = Convert.ToInt32(parameters[3]);
                point.Numer_przelicznika = parameters[4];
                point.Cieplomierz_RDS_PP = parameters[5];
                point.WWU1_numer = parameters[6];
                point.WWU1_RDS_PP = parameters[7];
                point.WWU2_numer = parameters[8];
                point.WWU2_RDS_PP = parameters[9];
                point.WCW_numer = parameters[10];
                point.WCW_RDS_PP = parameters[11];
                point.WCO_Numer = parameters[12];
                point.WCO_RDS_PP = parameters[13];
                point.adres_wf = parameters[14];
                point.wskaznik_kontrolny = parameters[15];

                int newCounter = 0;

                foreach (var prevPoint in listOfPreviousPoints)
                {
                    if (prevPoint.UPW_Kombit == point.UPW_Kombit)
                    {
                        newCounter++;
                        if (point.Cieplomierz_RDS_PP != prevPoint.Cieplomierz_RDS_PP ||
                            point.WWU1_RDS_PP != prevPoint.WWU1_RDS_PP ||
                            point.WWU2_RDS_PP != prevPoint.WWU2_RDS_PP ||
                            point.WCW_RDS_PP != prevPoint.WCW_RDS_PP ||
                            point.WCO_RDS_PP != prevPoint.WCO_RDS_PP
                            )
                        {
                            point.isChanged = true;
                        }
                        else
                        {

                        }
                    }
                }

                if (newCounter == 0)
                {
                    point.isNew = true;
                }
                else
                {

                }


                listOfLatestPoints.Add(point);
            }

            int numberOfChangedPoints = 0;
            foreach (var point in listOfLatestPoints)
            {
                if (point.isChanged)
                {
                    numberOfChangedPoints++;
                }
            }
            Console.WriteLine("Liczba rekodrów w których RDSPP uleglo zmianie : {0}", numberOfChangedPoints);
            Console.WriteLine("Sa to nastepujace punkty sieci : ");
            foreach (var point in listOfLatestPoints)
            {
                if (point.isChanged)
                {
                    Console.WriteLine(point.UPW_Kombit + " - " + point.adres_wf);
                }
            }

            int numberOfNewPoints = 0;
            foreach (var point in listOfLatestPoints)
            {
                if (point.isNew)
                {
                    numberOfNewPoints++;
                }
            }
            Console.WriteLine("Liczba nowych rekodrów : {0}", numberOfNewPoints);
            Console.WriteLine("Sa to nastepujace punkty sieci : ");
            foreach (var point in listOfLatestPoints)
            {
                if (point.isNew)
                {
                    Console.WriteLine(point.UPW_Kombit + " - " + point.adres_wf);
                }
            }

            #endregion

            #region Parsowanie plików z listą RDSPP Vectora

            //Ciepłomierze
            var listOfVectorMeters = new List<string>();

            foreach (var item in logvectorMeterList)
            {
                string[] parameters = item.Split('|');

                listOfVectorMeters.Add(parameters[7]);
            }

            //Wodomierze
            var listOfVectorWodo = new List<string>();

            foreach (var item in logvectorWodoList)
            {
                string[] parameters = item.Split('|');

                listOfVectorWodo.Add(parameters[8]);
            }


            #endregion

            #region Kodowanie nowych RDS PP Ciepłomierzy

            foreach (var point in listOfLatestPoints)
            {
                if (!(String.IsNullOrWhiteSpace(point.Cieplomierz_RDS_PP)) && point.isChanged == false && point.isNew == true)
                {
                    Meter meter = new Meter(point.Cieplomierz_RDS_PP);

                    if (listOfVectorMeters.Contains(meter.LicznikEnergii))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikEnergii);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.LicznikEnergii);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikEnergii);
                            piServer.CreatePIPoint(meter.LicznikEnergii);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.LicznikEnergii);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "GJ");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikEnergii);
                    }

                    if (listOfVectorMeters.Contains(meter.LicznikObjetosciPrzeplywu))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikObjetosciPrzeplywu);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.LicznikObjetosciPrzeplywu);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikObjetosciPrzeplywu);
                            piServer.CreatePIPoint(meter.LicznikObjetosciPrzeplywu);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.LicznikObjetosciPrzeplywu);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.LicznikObjetosciPrzeplywu);
                    }

                    if (listOfVectorMeters.Contains(meter.MocChwilowa))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.MocChwilowa);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.MocChwilowa);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.MocChwilowa);
                            piServer.CreatePIPoint(meter.MocChwilowa);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.MocChwilowa);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "kW");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.MocChwilowa);
                    }

                    if (listOfVectorMeters.Contains(meter.PrzeplywChwilowy))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.PrzeplywChwilowy);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.PrzeplywChwilowy);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.PrzeplywChwilowy);
                            piServer.CreatePIPoint(meter.PrzeplywChwilowy);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.PrzeplywChwilowy);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3/h");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.PrzeplywChwilowy);
                    }

                    if (listOfVectorMeters.Contains(meter.TemperaturaZasilania))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaZasilania);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.TemperaturaZasilania);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaZasilania);
                            piServer.CreatePIPoint(meter.TemperaturaZasilania);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.TemperaturaZasilania);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "°C");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaZasilania);
                    }

                    if (listOfVectorMeters.Contains(meter.TemperaturaPowrotu))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaPowrotu);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.TemperaturaPowrotu);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaPowrotu);
                            piServer.CreatePIPoint(meter.TemperaturaPowrotu);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.TemperaturaPowrotu);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "°C");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.TemperaturaPowrotu);
                    }

                    if (listOfVectorMeters.Contains(meter.NumerFabryczny))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.NumerFabryczny);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.NumerFabryczny);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.NumerFabryczny);
                            piServer.CreatePIPoint(meter.NumerFabryczny);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.NumerFabryczny);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "String");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.NumerFabryczny);
                    }

                    if (listOfVectorMeters.Contains(meter.KodBledu))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.KodBledu);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.KodBledu);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.KodBledu);
                            piServer.CreatePIPoint(meter.KodBledu);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.KodBledu);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Int32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.KodBledu);
                    }

                    if (listOfVectorMeters.Contains(meter.GodzinyPracy))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.GodzinyPracy);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.GodzinyPracy);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.GodzinyPracy);
                            piServer.CreatePIPoint(meter.GodzinyPracy);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, meter.GodzinyPracy);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Int32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "h");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + meter.GodzinyPracy);
                    }

                }
                else
                {
                    continue;
                }



            }

            #endregion

            #region Kodowanie nowych RDS PP Wodomierzy

            //WWU_1
            foreach (var point in listOfLatestPoints)
            {
                if (!(String.IsNullOrWhiteSpace(point.WWU1_RDS_PP)) && point.isChanged == false && point.isNew == true)
                {
                    Wodo WWU_1 = new Wodo(point.WWU1_RDS_PP);

                    if (listOfVectorWodo.Contains(WWU_1.StanWodomierza))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_1.StanWodomierza);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WWU_1.StanWodomierza);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_1.StanWodomierza);
                            piServer.CreatePIPoint(WWU_1.StanWodomierza);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WWU_1.StanWodomierza);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_1.StanWodomierza);
                    }
                }
                else
                {
                    continue;
                }

            }

            //WWU_2
            foreach (var point in listOfLatestPoints)
            {
                if (!(String.IsNullOrWhiteSpace(point.WWU2_RDS_PP)) && point.isChanged == false && point.isNew == true)
                {
                    Wodo WWU_2 = new Wodo(point.WWU2_RDS_PP);

                    if (listOfVectorWodo.Contains(WWU_2.StanWodomierza))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_2.StanWodomierza);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WWU_2.StanWodomierza);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_2.StanWodomierza);
                            piServer.CreatePIPoint(WWU_2.StanWodomierza);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WWU_2.StanWodomierza);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WWU_2.StanWodomierza);
                    }
                }
                else
                {
                    continue;
                }

            }

            //WCW
            foreach (var point in listOfLatestPoints)
            {
                if (!(String.IsNullOrWhiteSpace(point.WCW_RDS_PP)) && point.isChanged == false && point.isNew == true)
                {
                    Wodo WCW = new Wodo(point.WCW_RDS_PP);

                    if (listOfVectorWodo.Contains(WCW.StanWodomierza))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCW.StanWodomierza);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WCW.StanWodomierza);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCW.StanWodomierza);
                            piServer.CreatePIPoint(WCW.StanWodomierza);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WCW.StanWodomierza);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCW.StanWodomierza);
                    }
                }
                else
                {
                    continue;
                }

            }

            //WCO
            foreach (var point in listOfLatestPoints)
            {
                if (!(String.IsNullOrWhiteSpace(point.WCO_RDS_PP)) && point.isChanged == false && point.isNew == true)
                {
                    Wodo WCO = new Wodo(point.WCO_RDS_PP);

                    if (listOfVectorWodo.Contains(WCO.StanWodomierza))
                    {
                        Console.WriteLine("JEST W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCO.StanWodomierza);
                        try
                        {
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WCO.StanWodomierza);
                            Console.WriteLine("JEST JUŻ W SYSTEM PI");
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("NIE znalazlem W SYSTEM PI więc tworzę : " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCO.StanWodomierza);
                            piServer.CreatePIPoint(WCO.StanWodomierza);
                            PIPoint myPIPoint = PIPoint.FindPIPoint(piServer, WCO.StanWodomierza);
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointType, "Float32");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Descriptor, "");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DisplayDigits, "3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.EngineeringUnits, "m3");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSource, "UFLOD");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Archiving, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.Compressing, "1");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionDeviation, "0");
                            myPIPoint.SetAttribute(PICommonPointAttributes.CompressionMaximum, "28800");
                            myPIPoint.SetAttribute(PICommonPointAttributes.DataSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(w) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(w)");
                            myPIPoint.SetAttribute(PICommonPointAttributes.PointSecurity, "piadmin: A(r,w) | piadmins: A(r,w) | PIInterfaces: A(r,w) | PICoresightSRV: A(r) | PIAnalysisServices: A(r,w) | PILO_write: A(r) | PILO_read: A(r) | PILO_super: A(r,w) | PI_Buffers: A(r)");
                            myPIPoint.SaveAttributes();
                            Console.WriteLine("utworzono {0}", myPIPoint.Name);
                        }
                    }
                    else
                    {
                        Console.WriteLine("NIE MA W VECTORZE! - " + point.UPW_Kombit + " - " + point.adres_wf + " - kod : " + WCO.StanWodomierza);
                    }
                }
                else
                {
                    continue;
                }

            }

            #endregion

            Console.WriteLine("Zakończono aktualizowanie tagów PI na podstawie referencji z GIS");

        }


    }
}