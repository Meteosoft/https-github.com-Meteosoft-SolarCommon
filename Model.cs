#define TESTING
#define UseAPIForPowerFlow
//#define UseAPIForPower

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Meteosoft.Common;
// ReSharper disable ClassNeverInstantiated.Global
// ReSharper disable UnassignedField.Global
// ReSharper disable UnusedAutoPropertyAccessor.Global
// ReSharper disable CollectionNeverUpdated.Global
// ReSharper disable UnusedMember.Global
// ReSharper disable MemberCanBePrivate.Global

namespace Meteosoft.SolarCommon
{
    /// <summary>The PV Monthly Shapes class</summary>
    public class Shapes
    {
        public class PVMonthlyShapes
        {
            public int Month { get; set; }
            public List<PVHourlyShape> HourlyShapes { get; set; } = new List<PVHourlyShape>();
        }

        public class PVHourlyShape
        {
            public int Hour { get; set; }
            public double Correction { get; set; }
        }

        public static Dictionary<int, Dictionary<int, double>> ShapeMaps { get; set; } = new Dictionary<int, Dictionary<int, double>>();
        public static List<PVMonthlyShapes> MonthlyShapes { get; set; } = new List<PVMonthlyShapes>();

        public static Dictionary<int, Dictionary<int, double>> CreateAndSaveDefaultMonthlyShapeMaps(string filePathAndName)
        {
            InitialiseShapes();
            SaveShapes(filePathAndName);
            CreateShapeMap();
            return ShapeMaps;
        }

        public static void CreateShapeMap()
        {
            //ShapeMaps = new Dictionary<int, Dictionary<int, double>>();
            ShapeMaps.Clear();
            foreach (PVMonthlyShapes monthlyShape in MonthlyShapes)
            {
                Dictionary<int, double> monthlyMap = new Dictionary<int, double>();
                foreach (PVHourlyShape hourlyShape in monthlyShape.HourlyShapes)
                {
                    monthlyMap.Add(hourlyShape.Hour, hourlyShape.Correction);
                }

                ShapeMaps.Add(monthlyShape.Month, monthlyMap);
            }
        }

        public static void SaveShapes(string filePathAndName)
        {
            using StreamWriter file = new StreamWriter(filePathAndName, false, Encoding.ASCII);
            string json = JSONUtilities.SerialiseClassToJSON(MonthlyShapes);
            file.Write(json);
        }

        public static void ReadShapes(string filePathAndName)
        {
            using StreamReader file = new StreamReader(filePathAndName);
            string json = file.ReadToEnd();
            MonthlyShapes = JSONUtilities.DeserialiseJSONToClass<List<PVMonthlyShapes>>(json);
            if (MonthlyShapes == null)
            {
                InitialiseShapes();
                SaveShapes(filePathAndName);
            }
            else
                CreateShapeMap();
        }

        public static void InitialiseShapes()
        {
            for (int month = 1; month <= 12; month++)
            {
                PVMonthlyShapes pvMonthlyShapes = new PVMonthlyShapes {Month = month};
                for (int hour = 1; hour <= 24; hour++)
                {
                    PVHourlyShape pvHourlyShape = new PVHourlyShape {Hour = hour, Correction = 1.0};
                    pvMonthlyShapes.HourlyShapes.Add(pvHourlyShape);
                }

                MonthlyShapes.Add(pvMonthlyShapes);
            }
        }

        public static Dictionary<int, Dictionary<int, double>> ReadAndCreateMonthlyShapeMaps()
        {






            return null;
        }
    }

    #region WxUnderground main classes
    /// <summary>The WxUnderground class</summary>
    public class WxUnderground
    {
        /// <summary>The WxUndergroundForecast class</summary>
        public class Forecast
        {
            public List<int> CloudCover { get; set; }
            public List<string> DayOfWeek { get; set; }
            public List<string> DayOrNight { get; set; }
            public List<int> ExpirationTimeUtc { get; set; }
            public List<int> IconCode { get; set; }
            public List<int> IconCodeExtend { get; set; }
            public List<int> PrecipChance { get; set; }
            public List<string> PrecipType { get; set; }
            public List<double> PressureMeanSeaLevel { get; set; }
            public List<double> Qpf { get; set; }
            public List<double> QpfSnow { get; set; }
            public List<int> RelativeHumidity { get; set; }
            public List<int> Temperature { get; set; }
            public List<int> TemperatureDewPoint { get; set; }
            public List<int> TemperatureFeelsLike { get; set; }
            public List<int> TemperatureHeatIndex { get; set; }
            public List<int> TemperatureWindChill { get; set; }
            public List<string> UvDescription { get; set; }
            public List<int> UvIndex { get; set; }
            public List<DateTime> ValidTimeLocal { get; set; }
            public List<int> ValidTimeUtc { get; set; }
            public List<double> Visibility { get; set; }
            public List<int> WindDirection { get; set; }
            public List<string> WindDirectionCardinal { get; set; }
            public List<int?> WindGust { get; set; }
            public List<int> WindSpeed { get; set; }
            public List<string> WxPhraseLong { get; set; }
            public List<string> WxPhraseShort { get; set; }
            public List<int> WxSeverity { get; set; }
        }

        public static string ErrorMessage { get; set; }
        public static Forecast GetWUndergroundForecast(string apiKey)
        {
            string json = Data.GetDataUsingAPI($"https://api.weather.com/v3/wx/forecast/hourly/15day?apiKey={apiKey}&geocode=-32.058%2C115.958&units=m&language=en-US&format=json");
            if (!string.IsNullOrEmpty(Data.LastErrorMessage))
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            Forecast wxUndergroundForecast = JSONUtilities.DeserialiseJSONToClass<Forecast>(json);
            if (string.IsNullOrEmpty(JSONUtilities.LastErrorMessage))
                return wxUndergroundForecast;
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return null;
        }

        public static int GetCloudCover(Forecast forecast, DateTime dateTime)
        {
            if (forecast.ValidTimeLocal.Contains(dateTime))
            {
                int arrayIndex = forecast.ValidTimeLocal.IndexOf(dateTime);
                return forecast.CloudCover[arrayIndex];
            }
            return -1;
        }
    }
    #endregion WxUnderground main classes

    /// <summary>The SolarEdge class</summary>
    public class SolarEdge
    {
        #region Inverter
        public class SolarEdgeInverterInternalData
        {
            public string Time { get; set; }
            public float ACCurrent { get; set; }
            public float ACVoltage { get; set; }
            public float ACPower { get; set; }
            public float ACFrequency { get; set; }
            public float ACPowerFactor { get; set; }
            public float ACEnergy { get; set; }
            public float DCCurrent { get; set; }
            public float DCVoltage { get; set; }
            public float DCPower { get; set; }
            public float HSTemp { get; set; }
            public ushort Status { get; set; }
        }

        public class InternalInverterDetailsRoot
        {
            public List<SolarEdgeInverterInternalData> solarEdgeInverterInternalData { get; set; }
        }
        #endregion Inverter

        #region Common
        public class SolarEdgeInverterCommonData
        {
            public string Manufacturer { get; set; }
            public string Model { get; set; }
            public string Version { get; set; }
            public string SerialNumber { get; set; }
        }

        public class CommonInverterDetailsRoot
        {
            public SolarEdgeInverterCommonData solarEdgeInverterCommonData { get; set; }
        }
        #endregion Common

        #region SolarEdge sub-classes
        /// <summary>Tariff times</summary>
        public class TariffTimes
        {
            public DateTime peakStart;
            public DateTime peakEnd;
            public DateTime shoulderStart;
            public DateTime shoulderEnd;
            public DateTime offPeakStart1;
            public DateTime offPeakEnd1;
            public DateTime offPeakStart2;
            public DateTime offPeakEnd2;
            public bool isWeekend;
            //public bool isHoliday;
        }

        /// <summary>SolarEdge tech details</summary>
        public class SolarEdgeDetails
        {
            public string SiteID { get; set; } = "394525";
            public string SolarEdgeAPIKey { get; set; } = "NDX1KPEAJ853R2CYJVQXXUBYLIEKG1C2";
            public string InverterSerialNum { get; set; } = "7F150074-08";
            public double EnergyPurchaseRate { get; set; } = 25.7520;
            public double EnergyFeedInRate { get; set; } = 7.135;
            public double DailySupplyCharge { get; set; } = 92.3175;
            public double GST { get; set; } = 10.0;
            public double SmartPeakRate { get; set; } = 53.8714;
            public double SmartStandardRate { get; set; } = 28.2139;
            public double SmartOffPeakRate { get; set; } = 14.8405;
            public string SmartPeakRateStart { get; set; } = "15:00";
            public string SmartStandardRateStart { get; set; } = "07:00";
            public string SmartOffPeakRateStart { get; set; } = "21:00";
            public string SmartPeakRateEnd { get; set; } = "21:00";
            public string SmartStandardRateEnd { get; set; } = "15:00";
            public string SmartOffPeakRateEnd { get; set; } = "07:00";
            public bool PeakOnWeekends { get; set; } = false;
            public bool UseFlatTariff { get; set; } = true;
            public string UPSServer = "http://10.1.1.13:3052/agent/ppbe.js/init_status.js?s=1530239977045";
            public string DataPath { get; set; } = @"C:\Solar Inverter";
            public string NephiSysIP { get; set; } = "10.1.1.13";
            public string CloudDataPath { get; set; } = @"N:\NephiSys\Archive\";
            public string SolarEdgeCookie { get; set; } = "JSESSIONID=06708BA4D68920D3B063EB78B54ABCE2.sysr1; visid_incap_840553=pRVT08qyQu+6ybiCRsgzFFdqXVsAAAAAQUIPAAAAAACF3W0q+Cjh5Zia7TKjMnwu; _ga=GA1.2.1665910087.1508824897; SolarEdge_Locale=en_AU; __utmz=43987385.1535812068.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); SPRING_SECURITY_REMEMBER_ME_COOKIE=cm9iQG1ldGVvc29mdC5uZXQ6MzY4ODQ1MTM4MjA5MTowNDUyNmU0YWFhOTQyMzk5ZjQ4OWUyMDU5YjI3YTY3ZA; __utma=43987385.1665910087.1508824897.1536037078.1536208476.3; SolarEdge_Client-1.6=83b7e022b997536faae6c8b780cd40cf423b50622ae62b4a6820d6f1bbb05b5b7aab002304e6ca4826301eb741a54f6ec117af9dd0188278dd8a8c2e7c3b6986be5ee72c692f314d81e3fea658e57dbd33ae4e92264d9fa3568323560e4b9e0d5d0e2577c1d5c0acac56b79bed994774615d223bdaeef353e7e0b748cfca43e37d985b8359fd84c27a01548f66085b0353fae0393fdf9f0ea1100c13dedb31cedb1d97e45d97881dba776d43e1b10fcc9ed825e47543307da1afe89deefdb0257651d3aa21cacfca26de9772ae8a46e3def8e9eb7594611a8088968244b18c433d827b3b7c3c0771ec1a03d09aeff7e6b31e8ce2eade057ee04a08c311805ab41cb91fa66e3d666ad6753e3f9a73539f3c6870a6e9ba71041c294c655df5980d55b7feb565bec2359f4f512957e42d416121cf52d943104227a932158c49927b61f0b7e3dea0b1f7cd31a07291c0377ddc602d67990eeba998511da11585868f347b76575878acfe6e6b1e5bf1282bf9546e40fbdfb0674077f074d2b4ec09881448b2c6f2b8f000a1c0203ec0cafc4cdff67d03517b273b837a324bf8edf220b19c32be477ac71587eaa89eb1a5a864d590d5b8fd3fe10d76e9ab8647fef651aa0cb95ec1398b2b51193a120f22cd0d8fc84f658e38acd4f82c83d324c73702c9fcc240d5c533fb5517af2541420c7d17936c6183aa101ef2346ef544357f549eaf25100329be69e2b0e4a0ab40f19eb9b8f078a3a64c9170f896ede5fc5f548c10c6dbd021cf8397c550b73ceaf1da09f3d01a98e130a0f89697efdcc437252be8df75643ee5a13b06e7e56bef857f88951dd71e414eb30833589aae7b7de2988dfe8aeae44d93a84c88edbbe812eb06df3ce698de75c7828501d9b04f82f5037557220d808d6c8eea215f1da394adc41739e59145cdc129085a5ad5c9f3fde0b621910d56b7f0360fec22df0f21573db14c28e3f98c02dce2a2db0f43bfa84b6a7b8ba1a399d8eeba96c7e8387d6b94e39d55871d05282d8201bf878310f318a176e15930d4ec9594ec8fc4654fbbc0b7570fdebc79a75b66ef94e19ebda8f29c141981823b2dccfd0fa69306deca4390675aa1a2131c7a30847884ad9319107ddddabccad9e0bc931f56ac54d7459eb66cf81ff5acdf803178c3bdb41ac2f1b767e4d434853983ad06ab6f68aff846f2c181756797e21b4bb4c03e0dbadb839756dae87e5a1d098d20f803227fef988fe24f40a4374b2fb97233abaede5f23b3ef2d9efc3df2877fbc764db481a79cf18357564e4d60bc6eb34f3bec971b74a348c23d7e67fa0f5b51b522f7af0463312993daf8636426059a64e111dad773cc4acd6d23a52f21698c151000796f0a7d9191af92e004ffff7551c2893e166ddda867315b11228b7a9b4651d1348e074107ca6fdb2e490d06fc64741bce7a15b47a4f784e7f50c51414de0ab2b46e9ed519ab06680e5042768c310205834b30fe75104264447bd927f6bbcc9eeec7e7ceb2b3782d1c752cffe1f02bdc1cc8493fc2dd27a844dabe8bb4ee29f0bfce856610b250b4ed330c31b0542d88dba0eb679c99702b8920c285a654675d3ce5ba337229f156c081ea04a84a2765540697488885ee93058d7c0a07501304cd09d4f18eb9bdf92b8c7e050c998088b17fca70311be8bbc8ab1de0484183d3ae978feac9a231ad42fec0459fd19580449d39e10e66c4ed4bdc8a6c0aeb7a0dd2e9; SolarEdge_SSO-1.4=937763bb8cfba191a45700cb2e3b2bd4252e6e141f873d3433b272040483dc91; SolarEdge_Field_ID=394525; SolarEdge_Locale=en_AU";
                                                        //"JSESSIONID=F7B190D0CD3E80A40D7D1E0267A1E0C7.sysr1; visid_incap_840553=pRVT08qyQu+6ybiCRsgzFFdqXVsAAAAAQUIPAAAAAACF3W0q+Cjh5Zia7TKjMnwu; _ga=GA1.2.1665910087.1508824897; SolarEdge_Locale=en_AU; __utmz=43987385.1535812068.1.1.utmcsr=(direct)|utmccn=(direct)|utmcmd=(none); SPRING_SECURITY_REMEMBER_ME_COOKIE=cm9iQG1ldGVvc29mdC5uZXQ6MzY4ODg2NjQwMTU0MjpmNTAyYzUzOGZlNWE4MjlmNGUxMTAxYTljMTk0MjgzNg; __utma=43987385.1665910087.1508824897.1536037078.1536208476.3; CSRF-TOKEN=EF556FD1B5F6DC1506DDA2D1B72F36341BEEFD66A606137046621FE1EC0BFA701FF4D17B0CE7FE97DE98C77F20AE820BB55E; SolarEdge_Client-1.6=83b7e022b997536faae6c8b780cd40cf423b50622ae62b4a6820d6f1bbb05b5b7aab002304e6ca4826301eb741a54f6ec117af9dd0188278dd8a8c2e7c3b6986be5ee72c692f314d81e3fea658e57dbd33ae4e92264d9fa3568323560e4b9e0d5d0e2577c1d5c0acac56b79bed994774615d223bdaeef353e7e0b748cfca43e37d985b8359fd84c27a01548f66085b0353fae0393fdf9f0ea1100c13dedb31cedb1d97e45d97881dba776d43e1b10fcc9ed825e47543307da1afe89deefdb0257651d3aa21cacfca26de9772ae8a46e3def8e9eb7594611a8088968244b18c433d827b3b7c3c0771ec1a03d09aeff7e6b31e8ce2eade057ee04a08c311805ab41cb91fa66e3d666ad6753e3f9a73539f3c6870a6e9ba71041c294c655df5980d55b7feb565bec2359f4f512957e42d416121cf52d943104227a932158c49927b61f0b7e3dea0b1f7cd31a07291c0377ddc602d67990eeba998511da11585868f347b76575878acfe6e6b1e5bf1282bf9546e40fbdfb0674077f074d2b4ec09881448b2c6f2b8f000a1c0203ec0cafc4cdff67d03517b273b837a324bf8edf220b19c32be477ac71587eaa89eb1a5a864d590d5b8fd3fe10d76e9ab8647fef651aa0cb95ec1398b2b51193a120f22cd0d8fc84f658e38acd4f82c83d324c73702c9fcc240d5c533fb5517af2541420c7d17936c6183aa101ef2346ef544357f549eaf25100329be69e2b0e4a0ab40f19eb9b8f078a3a64c9170f896ede5fc5f548c10c6dbd021cf8397c550b73ceaf1da09f3d01a98e130a0f89697efdcc437252be8df75643ee5a13b06e7e56bef857f88951dd71e414eb30833589aae7b7de2988dfe8aeae44d93a84c88edbbe812eb06df3ce698de75c7828501d9b04f82f5037557220d808d6c8eea215f1da394adc41739e59145cdc129085a5ad5c9f3fde0b621910d56b7f0360fec22df0f21573db14c28e3f98c02dce2a2db0f43bfa84b6a7b8ba1a399d8eeba96c7e8387d6b94e39d55871d05282d8201bf878310f318a176e15930d4ec9594ec8fc4654fbbc0b7570fdebc79a75b66ef94e19ebda8f29c141981823b2dccfd0fa69306deca4390675aa1a2131c7a30847884ad9319107ddddabccad9e0bc931f56ac54d7459eb66cf81ff5acdf803178c3bdb41ac2f1b767e4d434853983ad06ab6f68aff846f2c181756797e21b4bb4c03e0dbadb839756dae87e5a1d098d20f803227fef988fe24f40a4374b2fb97233abaede5f23b3ef2d9efc3df2877fbc764db481a79cf18357564e4d60bc6eb34f3bec971b74a348c23d7e67fa0f5b51b522f7af0463312993daf8636426059a64e111dad773cc4acd6d23a52f21698c151000796f0a7d9191af92e004ffff7551c2893e166ddda867315b11228b7a9b4651d1348e074107ca6fdb2e490d06fc64741bce7a15b47a4f784e7f50c51414de0ab2b46e9ed519ab06680e5042768c310205834b30fe75104264447bd927f6bbcc9eeec7e7ceb2b3782d1c752cffe1f02bdc1cc8f075b88a5a114e80f181992f44a6250ee9850a01a0d943bdd8c32109f91ff23beb679c99702b8920c285a654675d3ce5ba337229f156c081ea04a84a2765540697488885ee93058d7c0a07501304cd09d4f18eb9bdf92b8c7e050c998088b17fca70311be8bbc8ab1de0484183d3ae978feac9a231ad42fec0459fd19580449d39e10e66c4ed4bdc8a6c0aeb7a0dd2e9; SolarEdge_SSO-1.4=937763bb8cfba191a45700cb2e3b2bd4252e6e141f873d3433b272040483dc91; SolarEdge_Locale=en_AU; SolarEdge_Field_ID=394525";
            public bool SimulateBatteryIfNotPresent { get; set; } = true;
            public int BatteryMaxPercent { get; set; } = 99;
            public int BatteryMinPercent { get; set; } = 25;
            public int BatteryMaxEnergy { get; set; } = 10;
        }

        public class DataMeterReadings
        {
            public class Power
            {
                public double Production;
                public double Consumption;
                public double SelfConsumption;
                public double FeedIn;
                public double Purchased;
                public double PhotoVoltaic;
                public double MyForecast;
                public double SolForecast;
                public double MyEstActual;
                public double BatteryCharge;
                public double BatteryPercent;
                public bool ReadingAreValid()
                {
                    return Production > 0.0 || Consumption > 0.0 || SelfConsumption > 0.0 ||
                           FeedIn > 0.0 || Purchased > 0.0;
                }
            }

            public class Energy
            {
                public double Production;
                public double Consumption;
                public double SelfConsumption;
                public double FeedIn;
                public double Purchased;
                public double PhotoVoltaic;
                public double MyForecast;
                public double SolForecast;
                public double MyEstActual;
                public double BatteryCharge;
                public bool ReadingAreValid()
                {
                    return Production > 0.0 || Consumption > 0.0 || SelfConsumption > 0.0 ||
                           FeedIn > 0.0 || Purchased > 0.0;
                }
            }
            public Power PowerReadings { get; set; } = new Power();
            public Energy EnergyReadings { get; set; } = new Energy();
        }

        /// <summary>PV site location details</summary>
        public class Location
        {
            /// <summary>The site location's country</summary>
            public string country { get; set; }

            /// <summary>The site location's state</summary>
            public string state { get; set; }

            /// <summary>The site location's city</summary>
            public string city { get; set; }

            /// <summary>The site location's address (1)</summary>
            public string address { get; set; }

            /// <summary>The site location's address (2)</summary>
            public string address2 { get; set; }

            /// <summary>The site location's post code</summary>
            public string zip { get; set; }

            /// <summary>The site location's time zone</summary>
            public string timeZone { get; set; }

            /// <summary>The site location's country code</summary>
            public string countryCode { get; set; }

            /// <summary>The site location's state code</summary>
            public string stateCode { get; set; }
        }

        /// <summary>PV panel details</summary>
        public class PrimaryModule
        {
            /// <summary>PV panel's manufacturer's name</summary>
            public string manufacturerName { get; set; }

            /// <summary>PV panel's model's name</summary>
            public string modelName { get; set; }

            /// <summary>PV panel's maximum power</summary>
            public double maximumPower { get; set; }

            /// <summary>PV panel's temperature coefficient</summary>
            public double temperatureCoef { get; set; }
        }

        /// <summary>URIs to be used to retrieve data</summary>
        public class Uris
        {
            /// <summary>URI to get DataPeriod</summary>
            public string DATA_PERIOD { get; set; }

            /// <summary>URI to get Overview</summary>
            public string OVERVIEW { get; set; }

            /// <summary>URI to get Details</summary>
            public string DETAILS { get; set; }
        }

        /// <summary>Class to hold public settings</summary>
        public class PublicSettings
        {
            /// <summary>Is the system available to the public?</summary>
            public bool isPublic { get; set; }
        }

        public class SiteDetails
        {
            public Details details { get; set; }
        }

        /// <summary>PV site details</summary>
        public class Details
        {
            /// <summary>PV ID</summary>
            public int id { get; set; }

            /// <summary>PV Name</summary>
            public string name { get; set; }

            /// <summary>PV AccountId</summary>
            public int accountId { get; set; }

            /// <summary>PV Status</summary>
            public string status { get; set; }

            /// <summary>PV PeakPowerKWh</summary>
            public double peakPower { get; set; }

            /// <summary>PV LastUpdateTime</summary>
            public string lastUpdateTime { get; set; }

            /// <summary>PV Currency</summary>
            public string currency { get; set; }

            /// <summary>PV InstallationDate</summary>
            public string installationDate { get; set; }

            /// <summary>PV Notes</summary>
            public string notes { get; set; }

            /// <summary>PV Type</summary>
            public string type { get; set; }

            /// <summary>PV SiteLocation</summary>
            public Location location { get; set; }

            /// <summary>PV PanelInfo</summary>
            public PrimaryModule primaryModule { get; set; }

            /// <summary>PV ImportantUris</summary>
            public Uris uris { get; set; }

            /// <summary>PV public settings</summary>
            public PublicSettings publicSettings { get; set; }
        }

        /// <summary>A class to hold the value of a component at a particular date/time</summary>
        public class Value
        {
            /// <summary>The date/time of the value</summary>
            public string date { get; set; }

            /// <summary>The value</summary>
            public double value { get; set; }

            public Value()
            {
            }

            public Value(string valueDate, double val)
            {
                date = valueDate;
                value = val;
            }
        }

        /// <summary>A class to hold the values of a battery at a particular date/time</summary>
        public class Telemetries
        {
            /// <summary>The date/time of the value</summary>
            public string timeStamp { get; set; }

            /// <summary>The battery power charging/discharging</summary>
            public double power { get; set; }

            /// <summary>The battery state (6=idle, 3=charging, 4=discharging)</summary>
            public int batteryState { get; set; }

            /// <summary>The battery charge percentage</summary>
            public double batteryPercentageState { get; set; }
        }

        public class PowerValues
        {
            public Power powerValues { get; set; }
        }

        /// <summary>A class to hold the power values</summary>
        public class Power
        {
            /// <summary>The time unit (interval) between values</summary>
            public string timeUnit { get; set; }

            /// <summary>The value's units</summary>
            public string unit { get; set; }

            /// <summary>The source of the power values</summary>
            public string measuredBy { get; set; }

            /// <summary>The power values</summary>
            public List<Value> values { get; set; }
        }

        public class Meter
        {
            /// <summary>Type of available reading (meter)</summary>
            public string type { get; set; }

            /// <summary>The meter readings</summary>
            public List<Value> values { get; set; }

            public Meter()
            {
            }

            public Meter(string meterType)
            {
                type = meterType;
                values = new List<Value>();
            }
        }

        public class PowerDetailsRoot
        {
            public PowerDetails powerDetails { get; set; }
        }

        public class PowerDetails
        {
            /// <summary>The time unit (interval) between values</summary>
            public string timeUnit { get; set; }

            /// <summary>The value's units</summary>
            public string unit { get; set; }

            /// <summary>The power values for each meter</summary>
            public List<Meter> meters { get; set; }
        }

        public class Battery
        {
            public List<StorageData> batteryList { get; set; }
        }

        public class StorageData
        {
            /// <summary>The number of batteries</summary>
            public int batteryCount { get; set; }

            /// <summary>The details of each battery</summary>
            public List<Batteries> batteries { get; set; }
        }

        public class Batteries
        {
            /// <summary>The nameplate value of this battery (max energy available)</summary>
            public double nameplate { get; set; }

            /// <summary>The battery's serial number</summary>
            public string serialNumber { get; set; }

            /// <summary>The number of battery values given</summary>
            public double telemetryCount { get; set; }

            /// <summary>The power values for each meter</summary>
            public List<Telemetries> telemetries { get; set; }
        }

        public class EnergyDetailsList
        {
            //public EnergyDetails energyDetails { get; set; }
            public List<EnergyDetails> energyDetailsList { get; set; }
        }

        public class EnergyDetails
        {
            public string timeUnit { get; set; }
            public string unit { get; set; }
            public List<Meter> meters { get; set; }
        }

        /// <summary>Current connection details, between components</summary>
        public class Connection
        {
            /// <summary>The source connection (e.g. PV)</summary>
            public string from { get; set; }

            /// <summary>The destination connection (e.g. LOAD)</summary>
            public string to { get; set; }
        }

        public class GRID
        {
            public string status { get; set; }
            public double currentPower { get; set; }
        }

        public class LOAD
        {
            public string status { get; set; }
            public double currentPower { get; set; }
        }

        public class PV
        {
            public string status { get; set; }
            public double currentPower { get; set; }
        }

        public class STORAGE
        {
            public string status { get; set; }
            public double currentPower { get; set; }
            public double chargeLevel { get; set; }
            public bool critical { get; set; }
            public double timeLeft { get; set; }
        }

        public class SitePowerFlowList
        {
            public List<SitePowerFlowItem> PowerFlowList = new List<SitePowerFlowItem>();
        }

        public class SitePowerFlowItem
        {
            public string Time;
            public SiteCurrentPowerFlow SitePowerFlow;
        }

        public class SitePowerFlow
        {
            public SiteCurrentPowerFlow siteCurrentPowerFlow { get; set; }
        }

        public class SiteCurrentPowerFlow
        {
            public int updateRefreshRate { get; set; }
            public string unit { get; set; }
            public List<Connection> connections { get; set; }
            public GRID GRID { get; set; }
            public LOAD LOAD { get; set; }
            public PV PV { get; set; }
            public STORAGE STORAGE { get; set; }
        }

        public class List
        {
            public string name { get; set; }
            public string manufacturer { get; set; }
            public string model { get; set; }
            public string serialNumber { get; set; }
            public object kWpDC { get; set; }
        }

        public class ReportersList
        {
            public List<Reporters> reportersList { get; set; }
        }

        public class Reporters
        {
            public int count { get; set; }
            public List<List> list { get; set; }
        }

        public class L1Data
        {
            public double acCurrent { get; set; }
            public double acVoltage { get; set; }
            public double acFrequency { get; set; }
            public double apparentPower { get; set; }
            public double activePower { get; set; }
            public double reactivePower { get; set; }
            public double cosPhi { get; set; }
        }

        public class Telemetry
        {
            public string date { get; set; }
            public double totalActivePower { get; set; }
            public double dcVoltage { get; set; }
            public double groundFaultResistance { get; set; }
            public double powerLimit { get; set; }
            public double totalEnergy { get; set; }
            public double temperature { get; set; }
            public string inverterMode { get; set; }
            public L1Data L1Data { get; set; }
        }

        public class TechDataList 
        {
            public List<TechData> techDataList{ get; set; }
        }

        public class TechData
        {
            public int count { get; set; }
            public List<Telemetry> telemetries { get; set; }
        }

        public class APIVersion
        {
            public Version apiVersion { get; set; }
        }

        public class Version
        {
            public string release { get; set; }
        }

        public class LoginData
        {
            public bool success { get; set; }
            public bool isLicensed { get; set; }
        }
        #endregion SolarEdge sub-classes
    
        #region SolarEdge main classes
        public static string ErrorMessage { get; set; }
        public static bool FileForThisDayComplete { get; set; }

        public static APIVersion GetAPIVersion(string apiKey)
        {
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI($"https://monitoringapi.solaredge.com/version/current?api_key={apiKey}");
            if (!string.IsNullOrEmpty(Data.LastErrorMessage))
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            APIVersion version = JSONUtilities.DeserialiseJSONToClass<APIVersion>(json);
            if (string.IsNullOrEmpty(JSONUtilities.LastErrorMessage)) 
                return version;
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return null;
        }

        public static TechDataList GetInverterTechData(DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            TechDataList returnList = new TechDataList();
            int numDays = (endDate - startDate).Days;

            // First attempt to retrieve all the daily data from archive files
            if (numDays == 1)
            {
                TechDataList inverterTechDataList = GetDetailedSiteTechDataFromFile(startDate, endDate, dataPath);
                if (inverterTechDataList != null && inverterTechDataList.techDataList.Count > 0)
                    return inverterTechDataList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            FileForThisDayComplete = false;
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/equipment/{solarEdgeDetails.SiteID}/{solarEdgeDetails.InverterSerialNum}/data?api_key={solarEdgeDetails.SolarEdgeAPIKey}" + 
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetDataUsingAPI(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            TechData inverterTechData = JSONUtilities.DeserialiseJSONToClass<TechData>(json);
            if (inverterTechData != null)
                FileForThisDayComplete = true;
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            returnList.techDataList.Add(inverterTechData);
            return returnList;
        }

        private static TechDataList GetDetailedSiteTechDataFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            TechDataList returnList = new TechDataList();
            TimeSpan span = endDate - startDate;
            int numDays = (int)span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string subDir = dataPath + @"\" + startDate.ToString("yyyy-MM-dd");
                if (!Directory.Exists(subDir))
                    Directory.CreateDirectory(subDir);
                string fileName = subDir + $"\\Tech {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    TechData inverterTechData = JSONUtilities.DeserialiseJSONToClass<TechData>(json);
                    if (!FileForThisDayComplete)
                        return null;
                    returnList.techDataList.Add(inverterTechData);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }

        public static ReportersList GetSiteEquipment(SolarEdgeDetails solarEdgeDetails)
        {
            string json = Data.GetDataUsingAPI($"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/list?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
            if (!string.IsNullOrEmpty(Data.LastErrorMessage))
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            ReportersList siteEquipment = JSONUtilities.DeserialiseJSONToClass<ReportersList>(json);
            if (string.IsNullOrEmpty(JSONUtilities.LastErrorMessage)) 
                return siteEquipment;
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return null;
        }

        public static SitePowerFlow GetSitePowerFlow(SolarEdgeDetails solarEdgeDetails, bool useAPI = false)
        {
            ErrorMessage = "";
            string json;
            if (!useAPI)
            {
                json = Data.GetDataDirect(
                    "https://monitoring.solaredge.com/solaredge-apigw/api/site/394525/currentPowerFlow.json");
                if (!string.IsNullOrEmpty(Data.LastErrorMessage))
                {
                    ErrorMessage = Data.LastErrorMessage;
                    return null;
                }
            }
            else
            {
                json = Data.GetDataUsingAPI(
                    $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/currentPowerFlow?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
                if (!string.IsNullOrEmpty(Data.LastErrorMessage))
                {
                    ErrorMessage = Data.LastErrorMessage;
                    return null;
                }
            }

            if (json.Contains("null"))
                json = json.Replace("null", "-1");
            SitePowerFlow sitePowerFlow = JSONUtilities.DeserialiseJSONToClass<SitePowerFlow>(json);
            if (string.IsNullOrEmpty(JSONUtilities.LastErrorMessage)) 
                return sitePowerFlow;
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return null;
        }

        public static SiteCurrentPowerFlow ComputeSitePowerFlow(double production, double consumption, double feedIn, double purchased, double storage)
        {
            SiteCurrentPowerFlow sitePowerFlow = new SiteCurrentPowerFlow
            {
                connections = new List<Connection>(),
                unit = "kW",
                PV = new PV {currentPower = production},
                LOAD = new LOAD {currentPower = consumption},
                GRID = new GRID {currentPower = Math.Abs(purchased - feedIn)},
                STORAGE = new STORAGE {currentPower = storage}
            };

            Connection connection;
            if (production > 0.0)
            {
                connection = new Connection
                {
                    @from = "PV",
                    to = "LOAD"
                };
                sitePowerFlow.connections.Add(connection);
            }
            if (purchased > feedIn)
            {
                connection = new Connection
                {
                    @from = "GRID",
                    to = "LOAD"
                };
                sitePowerFlow.connections.Add(connection);
            }
            else
            {
                connection = new Connection
                {
                    @from = "LOAD",
                    to = "GRID"
                };
                sitePowerFlow.connections.Add(connection);
            }

            return sitePowerFlow;
        }

        public static EnergyDetailsList GetDetailedSiteEnergy (DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            EnergyDetailsList detailedSiteEnergyList = new EnergyDetailsList();
            fromFile = true;
            int numDays = (endDate - startDate).Days;

            // First attempt to retrieve all the daily data from archive files
            if (numDays == 1)
            {
                detailedSiteEnergyList = GetDetailedSiteEnergyFromFile(startDate, endDate, dataPath);
                if (detailedSiteEnergyList != null && detailedSiteEnergyList.energyDetailsList?.Count > 0)
                    return detailedSiteEnergyList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            ErrorMessage = "";
            string timeUnit = "QUARTER_OF_AN_HOUR";
            if (numDays > 1)
                timeUnit = "DAY";
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/energyDetails?api_key={solarEdgeDetails.SolarEdgeAPIKey}&timeUnit={timeUnit}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetDataUsingAPI(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            EnergyDetails detailedSiteEnergy = JSONUtilities.DeserialiseJSONToClass<EnergyDetails>(json);
            if (detailedSiteEnergy != null)
                FileForThisDayComplete = true;
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }

            // Convert all watt hours to KWh
            if (detailedSiteEnergyList != null && detailedSiteEnergy != null)
            {
                detailedSiteEnergyList.energyDetailsList?.Add(detailedSiteEnergy);
                if (detailedSiteEnergyList.energyDetailsList != null)
                    foreach (EnergyDetails item in detailedSiteEnergyList.energyDetailsList)
                    {
                        item.unit = "kWh";
                        foreach (Meter meter in item.meters)
                        {
                            foreach (Value value in meter.values)
                            {
                                value.value /= 1000.0;
                            }
                        }
                    }
            }

            return detailedSiteEnergyList;
        }

        private static EnergyDetailsList GetDetailedSiteEnergyFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            EnergyDetailsList returnList = new EnergyDetailsList();
            TimeSpan span = endDate - startDate;
            int numDays = (int)span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string subDir = dataPath + @"\" + startDate.ToString("yyyy-MM-dd");
                if (!Directory.Exists(subDir))
                    Directory.CreateDirectory(subDir);
                string fileName = subDir + $"\\Energy {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    EnergyDetails detailedSiteEnergy = JSONUtilities.DeserialiseJSONToClass<EnergyDetails>(json);
                    if (!FileForThisDayComplete)
                        return null;
                    returnList.energyDetailsList.Add(detailedSiteEnergy);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }

        public static List<PowerDetailsRoot> GetDetailedSitePower(DateTime startDate, DateTime endDate, bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            List<PowerDetailsRoot> powerList = null;

            // First attempt to retrieve all the data from archive files if NOT in live mode
            if (!liveUpdateMode)
            {
                powerList = GetDetailedSitePowerFromFile(startDate, endDate, dataPath, out bool convertedFromOld);
                fromFile = !convertedFromOld;
                if (powerList.Count > 0)
                    return powerList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            ErrorMessage = "";
#if UseAPIForPower
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/powerDetails?api_key={solarEdgeDetails.APIKey}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetData(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            detailedSitePower = JSONUtilities.DeserialiseJSONToClass<DetailedSitePower>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            detailedSitePowerList.Add(detailedSitePower);
#else
            powerList = new List<PowerDetailsRoot>();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            for (int j = 0; j < numDays; j++)
            {
                string csv = "";
                DateTime epoch = new DateTime(1970, 1, 1);
                long startDateJs = startDate.Subtract(epoch).Ticks / 10000;
                DateTime jsEnd = new DateTime(startDate.Year, startDate.Month, startDate.Day, 23, 59, 0);
                long endDateJs = jsEnd.Subtract(epoch).Ticks / 10000;
                string http =
                    $"https://monitoringpublic.solaredge.com/solaredge-web/p/charts/{solarEdgeDetails.SiteID}/chartExport?" +
                    $"st={startDateJs}&et={endDateJs}&fid={solarEdgeDetails.SiteID}&timeUnit=2&pn0=Power&id0=0&t0=0&hasMeters=false";
                Data.LastErrorMessage = "";
                csv = Data.GetDataUsingAPI(http);
                ErrorMessage = Data.LastErrorMessage;
                if (csv == null)
                    return null;
                //csv = File.ReadAllText(@"C:\temp\SolarEdgeDump.csv");

                //**************************************
                //File.WriteAllText(@"C:\temp\SolarEdgeDump.csv", csv);
                //**************************************

                if (JSONUtilities.LastErrorMessage != "")
                {
                    ErrorMessage = JSONUtilities.LastErrorMessage;
                    return null;
                }

                try
                {
                    Dictionary<string, List<Value>> csvMapping = new Dictionary<string, List<Value>>();
                    List<string> headerNames = new List<string>();
                    string[] rows = csv.Split('\n');
                    bool haveHeaders = false;
                    foreach (string row in rows)
                    {
                        if (string.IsNullOrEmpty(row))
                            continue;
                        string[] cols = row.Split(',');
                        if (!haveHeaders)
                        {
                            // [Time,]Battery Charge Level (%),Consumption (W),System Production (W),StoragePower.Power (W),Export (W),Self Consumption (W),Import (W),Solar Production (W)
                            for (int i = 1; i < cols.Length; i++)
                            {
                                string col = cols[i].Trim();
                                csvMapping.Add(col.Trim(), new List<Value>());
                                headerNames.Add(col.Trim());
                            }

                            haveHeaders = true;
                        }
                        else
                        {
                            string valueTime = cols[0];
                            CultureInfo usCulture = CultureInfo.CreateSpecificCulture("en-US");
                            DateTime dt = DateTime.Now;
                            try
                            {
                                dt = DateTime.Parse(valueTime, usCulture);
                            }
                            catch (Exception x)
                            {
                                ErrorMessage = x.Message;
                            }

                            for (int i = 1; i < cols.Length; i++)
                            {
                                double value = Convert.ToDouble(cols[i].Trim().Replace("\"", ""));
                                csvMapping[headerNames[i - 1]]
                                    .Add(new Value(dt.ToString("yyyy-MM-dd HH:mm:ss"), value));
                            }
                        }
                    }

                    // Prepare JSON power details classes
                    //powerList = new List<DetailedSitePower>();
                    //PowerDetails detailedSitePower = new PowerDetails();
                    PowerDetailsRoot powerDetailsRoot = new PowerDetailsRoot
                    {
                        powerDetails = new PowerDetails
                        {
                            timeUnit = "QUARTER_OF_AN_HOUR",
                            unit = "W",
                            meters = new List<Meter>
                            {
                                new Meter("FeedIn"), // Index = 0
                                new Meter("Purchased"), // Index = 1
                                new Meter("Consumption"), // Index = 2
                                new Meter("SelfConsumption"), // Index = 3
                                new Meter("Production"), // Index = 4
                                new Meter("SolarProduction"), // Index = 5
                                new Meter("BatteryCharge"), // Index = 6
                                new Meter("BatteryPercent") // Index = 7
                            }
                        }
                    };

                    // Extract data from map into power details
                    foreach (KeyValuePair<string, List<Value>> keyValuePair in csvMapping)
                    {
                        string key = keyValuePair.Key;
                        List<Value> listOfValues = keyValuePair.Value;
                        switch (key)
                        {
                            case "Export.Power (W)":
                            case "Export (W)":
                                powerDetailsRoot.powerDetails.meters[0].values = listOfValues;
                                break;
                            case "Import.Power (W)":
                            case "Import (W)":
                                powerDetailsRoot.powerDetails.meters[1].values = listOfValues;
                                break;
                            case "Consumption.Power (W)":
                            case "Consumption (W)":
                                powerDetailsRoot.powerDetails.meters[2].values = listOfValues;
                                break;
                            case "SelfConsumption.Power (W)":
                            case "Self Consumption (W)":
                                powerDetailsRoot.powerDetails.meters[3].values = listOfValues;
                                break;
                            case "System Production (W)":
                                powerDetailsRoot.powerDetails.meters[4].values = listOfValues;
                                break;
                            case "SolarProduction.Power (W)":
                            case "Solar Production (W)":
                                powerDetailsRoot.powerDetails.meters[5].values = listOfValues;
                                break;
                            case "StoragePower.Power (W)":
                                powerDetailsRoot.powerDetails.meters[6].values = listOfValues;
                                break;
                            case "Battery Charge Level (%)":
                            case "StorageEnergyLevel.Power (%)":
                            case "Storage Energy Level (%)":
                                powerDetailsRoot.powerDetails.meters[7].values = listOfValues;
                                break;
                            case "Production.Power (W)":
                            case "Production (W)":
                                break;
                            default:
                                ErrorMessage = "Unknown Mapping Title: " + key;
                                break;
                        }
                    }
                    powerList.Add(powerDetailsRoot);
                }
                catch (Exception ex)
                {
                    ErrorMessage = ex.Message;
                    return null;
                }
                startDate += TimeSpan.FromDays(1);
            }
#endif
            return powerList;
        }

        public static List<PowerDetailsRoot> GetDetailedSitePowerFromFile(DateTime startDate, DateTime endDate, string dataPath, out bool convertedFromOld)
        {
            List<PowerDetailsRoot> detailedSitePower = new List<PowerDetailsRoot>();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            convertedFromOld = false;
            for (int i = 0; i < numDays; i++)
            {
                string subDir = dataPath + @"\" + startDate.ToString("yyyy-MM-dd");
                if (!Directory.Exists(subDir))
                    Directory.CreateDirectory(subDir);
                string fileName = subDir + $"\\Power {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    PowerDetailsRoot dailyDetails = JSONUtilities.DeserialiseJSONToClass<PowerDetailsRoot>(json);
                    if (JSONUtilities.LastErrorMessage != "")
                    {
                        ErrorMessage = JSONUtilities.LastErrorMessage;
                        return null;
                    }
                    if (dailyDetails != null)
                        detailedSitePower.Add(dailyDetails);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return detailedSitePower;
        }

        public static PowerValues GetSitePower(DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails)
        {
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/power?api_key={solarEdgeDetails.SolarEdgeAPIKey}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetDataUsingAPI(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            PowerValues sitePower = JSONUtilities.DeserialiseJSONToClass<PowerValues>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            return sitePower;
        }

        public static Battery GetDetailedBatteryPower(DateTime startDate, DateTime endDate, bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;

            // First attempt to retrieve all the data from archive files if NOT in live mode
            if (!liveUpdateMode)
            {
                Battery detailedBatteryPowerList = GetDetailedBatteryPowerFromFile(startDate, endDate, dataPath);
                if (detailedBatteryPowerList != null && detailedBatteryPowerList.batteryList.Count > 0)
                    return detailedBatteryPowerList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/storageData?api_key={solarEdgeDetails.SolarEdgeAPIKey}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetDataUsingAPI(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            Battery detailedBatteryPower = JSONUtilities.DeserialiseJSONToClass<Battery>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            return detailedBatteryPower;
        }

        private static Battery GetDetailedBatteryPowerFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            Battery returnList = new Battery();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string subDir = dataPath + @"\" + startDate.ToString("yyyy-MM-dd");
                if (!Directory.Exists(subDir))
                    Directory.CreateDirectory(subDir);
                string fileName = subDir + $"\\Battery {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    StorageData detailedBatteryPower = JSONUtilities.DeserialiseJSONToClass<StorageData>(json);
                    returnList.batteryList.Add(detailedBatteryPower);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }

        public static SiteDetails GetSiteDetails(SolarEdgeDetails solarEdgeDetails)
        {
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI($"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/details?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            SiteDetails siteDetails = JSONUtilities.DeserialiseJSONToClass<SiteDetails>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }

            return siteDetails;
        }

        //public static string GetDetailedSitePower(DateTime startDate, DateTime endDate,
        //    bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        //{
        //    string csv = "";
        //    fromFile = false;
        //    ErrorMessage = "";
        //    DateTime epoch = new DateTime(1970, 1, 1);
        //    TimeSpan span = endDate - startDate;
        //    int numDays = (int) span.TotalDays;
        //    for (int i = 0; i < numDays; i++)
        //    {
        //        DateTime today = startDate + TimeSpan.FromDays(i);
        //        long startDateJs = today.Subtract(epoch).Ticks / 10000;
        //        DateTime jsEnd = new DateTime(today.Year, today.Month, today.Day, 23, 59, 0);
        //        long endDateJs = jsEnd.Subtract(epoch).Ticks / 10000;
        //        string http = $"https://monitoringpublic.solaredge.com/solaredge-web/p/charts/{solarEdgeDetails.SiteID}/chartExport?" + 
        //                      $"st={startDateJs}&et={endDateJs}&fid={solarEdgeDetails.SiteID}&timeUnit=2&pn0=Power&id0=0&t0=0&hasMeters=false";
        //        Data.LastErrorMessage = "";
        //        csv = Data.GetData(http);
        //        ErrorMessage = Data.LastErrorMessage;
        //        if (csv == null)
        //            return null;
        //        if (JSONUtilities.LastErrorMessage != "")
        //        {
        //            ErrorMessage = JSONUtilities.LastErrorMessage;
        //            return null;
        //        }
        //    }

        //    return csv;
        //}

        //private static List<PowerDetails> GetDetailedSiteDataFromFile(DateTime startDate, DateTime endDate,
        //    string dataPath)
        //{
        //    List<DetailedDailySiteData> returnList = new List<DetailedDailySiteData>();
        //    TimeSpan span = endDate - startDate;
        //    int numDays = (int) span.TotalDays;
        //    for (int i = 0; i < numDays; i++)
        //    {
        //        DateTime today = startDate + TimeSpan.FromDays(i);
        //        string subDir = dataPath + @"\" + startDate.ToString("yyyy-MM-dd");
        //        if (!Directory.Exists(subDir))
        //            Directory.CreateDirectory(subDir);
        //        string fileName = subDir + $"\\DailySiteData {today.Day}-{today.Month}-{today.Year}.json";
        //        if (File.Exists(fileName))
        //        {
        //            string json = File.ReadAllText(fileName);
        //            DetailedDailySiteData detailedSitePower =
        //                JSONUtilities.DeserialiseJSONToClass<DetailedDailySiteData>(json);
        //            returnList.Add(detailedSitePower);
        //        }
        //    }

        //    return returnList;
        //}
        #endregion SolarEdge main classes
    }

    /// <summary>The SolCast class</summary>
    public class SolCast
    {
        #region SolCast sub-classes
        // SolCast - Tech details
        public class SolCastDetails
        {
            public string SolCastAPIKey { get; set; } = "Vg1lLkcUAXbSrC1HOWyxo4XKbYpOOfNv";
            public double Latitude { get; set; } = -32.064;
            public double Longitude { get; set; } = 115.963;
            public int PanelTilt1 { get; set; } = 25;
            public int PanelAzimuth1 { get; set; } = 360;
            public int NumPanels1 { get; set; } = 9;
            public int PanelTilt2 { get; set; } = 25;
            public int PanelAzimuth2 { get; set; } = 90;
            public int NumPanels2 { get; set; } = 10;
            public int TimeZone { get; set; } = 8;
            public int StationHeightMetres { get; set; } = 20;
        }

        // SolCast - Estimated Actuals/Latest
        public class EstimatedActual
        {
            public DateTime period_end { get; set; }
            public string period { get; set; }
            public double pv_estimate { get; set; }
        }

        public class SolCastEstimatedActualList
        {
            public List<EstimatedActual> estimated_actuals { get; set; }
        }

        // SolCast - Forecasts
        public class Forecast
        {
            public DateTime period_end { get; set; }
            public string period { get; set; }
            public double pv_estimate { get; set; }
            public int cloudPercent { get; set; }
        }

        public class SolCastForecastList
        {
            public List<Forecast> forecasts { get; set; }
        }
        #endregion SolCast sub-classes

        #region SolCast main classes
        public static string ErrorMessage { get; set; }

        public static SolCastEstimatedActualList GetSolCastEstimatedActualList(string apiKey)
        {
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI($"https://api.solcast.com.au/rooftop_sites/abf3-6594-6c45-e7e0/estimated_actuals?format=json&api_key={apiKey}");
            if (json == null)
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            SolCastEstimatedActualList solCastEstimatedActualList = JSONUtilities.DeserialiseJSONToClass<SolCastEstimatedActualList>(json);
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return JSONUtilities.LastErrorMessage != "" ? null : solCastEstimatedActualList;
        }

        public static SolCastEstimatedActualList GetSolCastEstimatedActualList(double latitude, double longitude, double capacity, double tilt, double azimuth, DateTime installDate, string apiKey)
        {
            double reqAzimuth = -azimuth;
            if (azimuth > 180)
                reqAzimuth += 360;
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI("https://api.solcast.com.au:443/pv_power/estimated_actuals/latest?" + 
                                               $"latitude={latitude:##.###}&longitude={longitude:##.###}&capacity={capacity:##}" + 
                                               $"&tilt={tilt:##.#}&azimuth={reqAzimuth:0}&install_date={installDate:yyyyMMdd}&api_key={apiKey}&format=json");
            if (json == null)
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            SolCastEstimatedActualList solCastEstimatedActualList = JSONUtilities.DeserialiseJSONToClass<SolCastEstimatedActualList>(json);
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return JSONUtilities.LastErrorMessage != "" ? null : solCastEstimatedActualList;
        }

        public static SolCastForecastList GetSolCastForecast(string apiKey)
        {
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI($"https://api.solcast.com.au/rooftop_sites/abf3-6594-6c45-e7e0/forecasts?format=json&api_key={apiKey}");
            if (json == null)
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            SolCastForecastList solCastForecastList = JSONUtilities.DeserialiseJSONToClass<SolCastForecastList>(json);
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return JSONUtilities.LastErrorMessage != "" ? null : solCastForecastList;
        }

        public static SolCastForecastList GetSolCastForecast(double latitude, double longitude, double capacity, double tilt, double azimuth, DateTime installDate, string apiKey)
        {
            double reqAzimuth = -azimuth;
            if (azimuth > 180)
                reqAzimuth += 360;
            ErrorMessage = "";
            string json = Data.GetDataUsingAPI("https://api.solcast.com.au:443/pv_power/forecasts?" + 
                                               $"latitude={latitude:##.###}&longitude={longitude:##.###}&capacity={capacity:##}" + 
                                               $"&tilt={tilt:##.#}&azimuth={reqAzimuth:0}&install_date={installDate:yyyyMMdd}&api_key={apiKey}&format=json");
            if (json == null)
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            SolCastForecastList solCastForecastList = JSONUtilities.DeserialiseJSONToClass<SolCastForecastList>(json);
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return JSONUtilities.LastErrorMessage != "" ? null : solCastForecastList;
        }
        #endregion SolCast main classes
    }

    /// <summary>The UPS class</summary>
    public class UPS
    {
        public class UpsUtility
        {
            public string state { get; set; }
            public bool stateWarning { get; set; }
            public string voltage { get; set; }
            public object frequency { get; set; }
            public object voltages { get; set; }
            public object currents { get; set; }
            public object frequencies { get; set; }
            public object powerFactors { get; set; }
        }

        public class UpsBypass
        {
            public string state { get; set; }
            public bool stateWarning { get; set; }
            public object voltage { get; set; }
            public object current { get; set; }
            public object frequency { get; set; }
            public object voltages { get; set; }
            public object currents { get; set; }
            public object frequencies { get; set; }
            public object powerFactors { get; set; }
        }

        public class UpsOutput
        {
            public string state { get; set; }
            public bool stateWarning { get; set; }
            public string voltage { get; set; }
            public object frequency { get; set; }
            public int load { get; set; }
            public int watt { get; set; }
            public object va { get; set; }
            public object current { get; set; }
            public bool outputLoadWarning { get; set; }
            public object outlet1 { get; set; }
            public object outlet2 { get; set; }
            public object activePower { get; set; }
            public object apparentPower { get; set; }
            public object reactivePower { get; set; }
            public object voltages { get; set; }
            public object currents { get; set; }
            public object frequencies { get; set; }
            public object powerFactors { get; set; }
            public object loads { get; set; }
            public object activePowers { get; set; }
            public object apparentPowers { get; set; }
            public object reactivePowers { get; set; }
            public object emergencyOff { get; set; }
            public object batteryExhausted { get; set; }
        }

        public class UpsBattery
        {
            public string state { get; set; }
            public bool stateWarning { get; set; }
            public string voltage { get; set; }
            public int capacity { get; set; }
            public int runtimeFormat { get; set; }
            public bool modularUpsRuntimeZero { get; set; }
            public bool runtimeFormatWarning { get; set; }
            public int runtimeHour { get; set; }
            public int runtimeMinute { get; set; }
            public object chargetimeFormat { get; set; }
            public object chargetimeHour { get; set; }
            public object chargetimeMinute { get; set; }
            public object temperatureCelsius { get; set; }
            public object highVoltage { get; set; }
            public object lowVoltage { get; set; }
            public object highCurrent { get; set; }
            public object lowCurrent { get; set; }
        }

        public class UpsSystem
        {
            public string state { get; set; }
            public bool stateWarning { get; set; }
            public object temperatureCelsius { get; set; }
            public object temperatureFahrenheit { get; set; }
            public object maintenanceBreak { get; set; }
            public object systemFaultDueBypass { get; set; }
            public object systemFaultDueBypassFan { get; set; }
            public string originalHardwareFaultCode { get; set; }
        }

        public class Status
        {
            public bool communicationAvaiable { get; set; }
            public bool onlyPhaseArch { get; set; }
            public UpsUtility utility { get; set; }
            public UpsBypass bypass { get; set; }
            public UpsOutput output { get; set; }
            public UpsBattery battery { get; set; }
            public UpsSystem upsSystem { get; set; }
            public object modules { get; set; }
            public int deviceId { get; set; }
        }

        public class UPSBatteryStatus
        {
            public Status status { get; set; }
        }

        public class UPSData
        {
            public DateTime time { get; set; }
            public int load { get; set; }
            public int batCap { get; set; }
            public int remainTime { get; set; }
            public float voltIn { get; set; }
            public float voltOut { get; set; }
            public bool statusInGood { get; set; }
            public bool statusOutGood { get; set; }
        }

        public class UPSConfig
        {
            public string UPSServer { get; set; } = "http://10.1.1.13:3052/agent/ppbe.js/init_status.js?s=1530239977045";
        }
    }
}
