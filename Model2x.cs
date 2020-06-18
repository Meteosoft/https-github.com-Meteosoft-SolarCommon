#define TESTING
//#define UseAPIForPower

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
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
    #region Sub Classes
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

        public Value() {}
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

        public Meter() {}
        public Meter(string meterType)
        {
            type = meterType;
            values = new List<Value>();
        }
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

    public class TechData
    {
        public int count { get; set; }
        public List<Telemetry> telemetries { get; set; }
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
    }

    public class SolCastForecastList
    {
        public List<Forecast> forecasts { get; set; }
    }
    #endregion Sub Classes

    #region Main SolCast Classes
    public class SolCastEstimatedActualsLatest
    {
        public static string ErrorMessage { get; set; }
        public static SolCastEstimatedActualList GetSolCastEstimatedActualList(double latitude, double longitude, double capacity, double tilt, double azimuth, DateTime installDate, string apiKey)
        {
            double reqAzimuth = -azimuth;
            if (azimuth > 180)
                reqAzimuth += 360;
            ErrorMessage = "";
            string json = Data.GetData("https://api.solcast.com.au:443/pv_power/estimated_actuals/latest?" + 
                                       $"latitude={latitude:##.###}&longitude={longitude:##.###}&capacity={capacity:##}" + 
                                       $"&tilt={tilt:##.#}&azimuth={reqAzimuth:0}&install_date={date:yyyyMMdd}&api_key={apiKey}&format=json");
            if (json == null)
            {
                ErrorMessage = Data.LastErrorMessage;
                return null;
            }
            SolCastEstimatedActualList solCastEstimatedActualList = JSONUtilities.DeserialiseJSONToClass<SolCastEstimatedActualList>(json);
            ErrorMessage = JSONUtilities.LastErrorMessage;
            return JSONUtilities.LastErrorMessage != "" ? null : solCastEstimatedActualList;
        }
    }

    public class SolCastForecast
    {
        public static string ErrorMessage { get; set; }
        public static SolCastForecastList GetSolCastForecast(double latitude, double longitude, double capacity, double tilt, double azimuth, DateTime installDate, string apiKey)
        {
            double reqAzimuth = -azimuth;
            if (azimuth > 180)
                reqAzimuth += 360;
            ErrorMessage = "";
            string json = Data.GetData("https://api.solcast.com.au:443/pv_power/forecasts?" + 
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
    }
    #endregion

    #region Main SolarEdge Classes
    public class LoginToSolarEdge
    {
        // ReSharper disable once StringLiteralTypo
        public static string m_recaptcha = @"03AL4dnxr5mkYrG9GKUPTe86UVW67y4y6qJwAcTYRQHK64QwA6Cbl5M_kZ8rOiK7T_6U6NWtZIk68Mofas0cfnXw5f8jlqnECDAqhjmlVjpWA7excNM8H6icIil09q3khcFAVVlM7z4GVRbl_Sy4a9VcF2LS6AwlbX-jwDh-fvySqfNxgwE7np67BCG4oae9uQUsKw51uFjhPv5yxuab33k5CIan1n_0BwAeUvpI_HjgtjKpICZ5xlfcikkTZrpbKoRWdYCY521Pk4HWjCQS6hB9sJA7wtIJsUkJ5_qfMVwxLcw9HWdLgycow";
        public static string ErrorMessage { get; set; }

        /// <summary>
        /// Log into the web site
        /// </summary>
        /// <param name="userName">The user name</param>
        /// <param name="password">The password</param>
        /// <returns>Whether successful, the user type and a secret token to be used for other requests</returns>
        public static bool Login(string userName, string password)
        {
            // Reset error messages
            ErrorMessage = "";
            Data.LastErrorMessage = "";

            // Format the request for data acquisition
            string parameters = "cmd=login";
            //parameters += "&demo=false";
            //parameters += "&g-recaptcha-response=" + m_recaptcha;
            //parameters += "&username=" + userName.Replace("@", "%40");
            //parameters += "&password=" + password;

            //// Login and get the response as a JSON structure
            //string json = Data.PostData("https://monitoring.solaredge.com/solaredge-web/p/", "submitLogin", parameters);
            //ErrorMessage = Data.LastErrorMessage;
            //if (Data.LastErrorMessage != "")
            //{
            //    ErrorMessage = Data.LastErrorMessage;
            //    return false;
            //}
            //if (json == null)
            //    return false;

            //// Deserialise the structure into a class
            //LoginData loginData = JSONUtilities.DeserialiseJSONToClass<LoginData>(json);
            //if (JSONUtilities.LastErrorMessage != "")
            //{
            //    ErrorMessage = JSONUtilities.LastErrorMessage;
            //    return false;
            //}

            //// Check for any errors from the web site
            //if (!loginData.success)
            //    return false;

            parameters = "username=" + userName.Replace("@", "%40");
            parameters += "&password=" + password;

            // Login
            string json = Data.PostData("https://monitoring.solaredge.com/solaredge-web/p/", "login", parameters);
            ErrorMessage = Data.LastErrorMessage;
            if (Data.LastErrorMessage != "")
            {
                ErrorMessage = Data.LastErrorMessage;
                return false;
            }
            return true;
        }
    }

    public class APIVersion
    {
        public Version version { get; set; }

        public static APIVersion GetAPIVersion(string apiKey)
        {
            string json = Data.GetData($"https://monitoringapi.solaredge.com/version/current?api_key={apiKey}");
            APIVersion APIVersion = JSONUtilities.DeserialiseJSONToClass<APIVersion>(json);
            if (JSONUtilities.LastErrorMessage != "")
                return null;
            return APIVersion;
        }
    }

    public class InverterTechData
    {
        public TechData data { get; set; }
        public bool fileForThisDayComplete { get; set; }
        public static string ErrorMessage { get; set; }

        public static List<InverterTechData> GetInverterTechData(DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            List<InverterTechData> returnList = new List<InverterTechData>();
            int numDays = (endDate - startDate).Days;

            // First attempt to retrieve all the daily data from archive files
            if (numDays == 1)
            {
                List<InverterTechData> inverterTechDataList = GetDetailedSiteTechDataFromFile(startDate, endDate, dataPath);
                if (inverterTechDataList != null && inverterTechDataList.Count > 0)
                    return inverterTechDataList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/equipment/{solarEdgeDetails.SiteID}/{solarEdgeDetails.InverterSerialNum}/data?api_key=NDX1KPEAJ853R2CYJVQXXUBYLIEKG1C2" + 
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetData(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            InverterTechData inverterTechData = JSONUtilities.DeserialiseJSONToClass<InverterTechData>(json);
            if (inverterTechData != null)
                inverterTechData.fileForThisDayComplete = true;
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            returnList.Add(inverterTechData);
            return returnList;
        }

        private static List<InverterTechData> GetDetailedSiteTechDataFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            List<InverterTechData> returnList = new List<InverterTechData>();
            TimeSpan span = endDate - startDate;
            int numDays = (int)span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string fileName = dataPath + $"\\Tech {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    InverterTechData inverterTechData = JSONUtilities.DeserialiseJSONToClass<InverterTechData>(json);
                    if (!inverterTechData.fileForThisDayComplete)
                        return null;
                    returnList.Add(inverterTechData);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }
    }

    public class SiteEquipment
    {
        public Reporters reporters { get; set; }

        public static SiteEquipment GetSiteEquipment(SolarEdgeDetails solarEdgeDetails)
        {
            string json = Data.GetData($"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/list?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
            SiteEquipment siteEquipment = JSONUtilities.DeserialiseJSONToClass<SiteEquipment>(json);
            if (JSONUtilities.LastErrorMessage != "")
                return null;
            return siteEquipment;
        }
    }

    public class SitePowerFlow
    {
        public SiteCurrentPowerFlow siteCurrentPowerFlow { get; set; }
        public static string ErrorMessage { get; set; }

        public static SitePowerFlow GetSitePowerFlow(SolarEdgeDetails solarEdgeDetails)
        {
            //string json = Data.GetData($"https://monitoring.solaredge.com/solaredge-apigw/api/site/{solarEdgeDetails.SiteID}/currentPowerFlow.json");
            string json = Data.GetData($"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/currentPowerFlow?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            if (json.Contains("null"))
                json = json.Replace("null", "-1");
            SitePowerFlow sitePowerFlow = JSONUtilities.DeserialiseJSONToClass<SitePowerFlow>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }

            return sitePowerFlow;
        }

        public static SitePowerFlow ComputeSitePowerFlow(double production, double consumption, double feedIn, double purchased, double storage)
        {
            SitePowerFlow sitePowerFlow = new SitePowerFlow
            {
                siteCurrentPowerFlow = new SiteCurrentPowerFlow
                {
                    connections = new List<Connection>(),
                    unit = "kW",
                    PV = new PV {currentPower = production}
                }
            };

            sitePowerFlow.siteCurrentPowerFlow.LOAD = new LOAD {currentPower = consumption};
            sitePowerFlow.siteCurrentPowerFlow.GRID = new GRID {currentPower = Math.Abs(purchased - feedIn)};
            sitePowerFlow.siteCurrentPowerFlow.STORAGE = new STORAGE {currentPower = storage};

            Connection connection;
            if (production > 0.0)
            {
                connection = new Connection
                {
                    from = "PV",
                    to = "LOAD"
                };
                sitePowerFlow.siteCurrentPowerFlow.connections.Add(connection);
            }
            if (purchased > feedIn)
            {
                connection = new Connection
                {
                    from = "GRID",
                    to = "LOAD"
                };
                sitePowerFlow.siteCurrentPowerFlow.connections.Add(connection);
            }
            else
            {
                connection = new Connection
                {
                    from = "LOAD",
                    to = "GRID"
                };
                sitePowerFlow.siteCurrentPowerFlow.connections.Add(connection);
            }

            return sitePowerFlow;
        }
    }

    public class DetailedSiteEnergy
    {
        public EnergyDetails energyDetails { get; set; }
        public bool fileForThisDayComplete { get; set; }
        public static string ErrorMessage { get; set; }

        public static List<DetailedSiteEnergy> GetDetailedSiteEnergy (DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            List<DetailedSiteEnergy> returnList = new List<DetailedSiteEnergy>();
            int numDays = (endDate - startDate).Days;

            // First attempt to retrieve all the daily data from archive files
            if (numDays == 1)
            {
                List<DetailedSiteEnergy> detailedSiteEnergyList = GetDetailedSiteEnergyFromFile(startDate, endDate, dataPath);
                if (detailedSiteEnergyList != null && detailedSiteEnergyList.Count > 0)
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
            string json = Data.GetData(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            DetailedSiteEnergy detailedSiteEnergy = JSONUtilities.DeserialiseJSONToClass<DetailedSiteEnergy>(json);
            if (detailedSiteEnergy != null)
                detailedSiteEnergy.fileForThisDayComplete = true;
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }

            // Convert all watt hours to KWh
            if (detailedSiteEnergy != null)
            {
                detailedSiteEnergy.energyDetails.unit = "kWh";
                foreach (Meter meter in detailedSiteEnergy.energyDetails.meters)
                {
                    foreach (Value value in meter.values)
                    {
                        value.value /= 1000.0;
                    }
                }
            }
            returnList.Add(detailedSiteEnergy);
            return returnList;
        }

        private static List<DetailedSiteEnergy> GetDetailedSiteEnergyFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            List<DetailedSiteEnergy> returnList = new List<DetailedSiteEnergy>();
            TimeSpan span = endDate - startDate;
            int numDays = (int)span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string fileName = dataPath + $"\\Energy {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    DetailedSiteEnergy detailedSiteEnergy = JSONUtilities.DeserialiseJSONToClass<DetailedSiteEnergy>(json);
                    if (!detailedSiteEnergy.fileForThisDayComplete)
                        return null;
                    returnList.Add(detailedSiteEnergy);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }
    }

    public class DetailedSitePower
    {
        public PowerDetails powerDetails { get; set; }
        public static string ErrorMessage { get; set; }

        public static List<DetailedSitePower> GetDetailedSitePower(DateTime startDate, DateTime endDate, bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            List<DetailedSitePower> powerList = null;

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
            powerList = new List<DetailedSitePower>();
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
                csv = Data.GetData(http);
                ErrorMessage = Data.LastErrorMessage;
                if (csv == null)
                    return null;
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
                    DetailedSitePower detailedSitePower = new DetailedSitePower();
                    PowerDetails powerDetails = new PowerDetails
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
                    };
                    detailedSitePower.powerDetails = powerDetails;

                    // Extract data from map into power details
                    foreach (KeyValuePair<string, List<Value>> keyValuePair in csvMapping)
                    {
                        string key = keyValuePair.Key;
                        List<Value> listOfValues = keyValuePair.Value;
                        switch (key)
                        {
                            case "Export (W)":
                                powerDetails.meters[0].values = listOfValues;
                                break;
                            case "Import (W)":
                                powerDetails.meters[1].values = listOfValues;
                                break;
                            case "Consumption (W)":
                                powerDetails.meters[2].values = listOfValues;
                                break;
                            case "Self Consumption (W)":
                                powerDetails.meters[3].values = listOfValues;
                                break;
                            case "System Production (W)":
                                powerDetails.meters[4].values = listOfValues;
                                break;
                            case "Solar Production (W)":
                                powerDetails.meters[5].values = listOfValues;
                                break;
                            case "StoragePower.Power (W)":
                                powerDetails.meters[6].values = listOfValues;
                                break;
                            case "Battery Charge Level (%)":
                                powerDetails.meters[7].values = listOfValues;
                                break;
                            default:
                                break;
                        }
                    }
                    powerList.Add(detailedSitePower);
                }
                catch (Exception ex)
                {
                    DetailedDailySiteData.ErrorMessage = ex.Message;
                    return null;
                }
                startDate += TimeSpan.FromDays(1);
            }
#endif
            return powerList;
        }

        private static List<DetailedSitePower> GetDetailedSitePowerFromFile(DateTime startDate, DateTime endDate, string dataPath, out bool convertedFromOld)
        {
            List<DetailedSitePower> detailedSitePower = new List<DetailedSitePower>();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            convertedFromOld = false;
            for (int i = 0; i < numDays; i++)
            {
                string fileName = dataPath + $"\\Power {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    DetailedSitePower dailyDetails = JSONUtilities.DeserialiseJSONToClass<DetailedSitePower>(json);
                    if (dailyDetails == null)
                    {
                        List<DetailedSitePower> dailyDetailsList = JSONUtilities.DeserialiseJSONToClass<List<DetailedSitePower>>(json);
                        if (dailyDetailsList != null)
                        {
                            dailyDetails = dailyDetailsList[0];
                            convertedFromOld = true;
                        }
                    }
                    if (dailyDetails != null)
                        detailedSitePower.Add(dailyDetails);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return detailedSitePower;
        }
    }

    public class SitePower
    {
        public Power power { get; set; }
        public static string ErrorMessage { get; set; }

        public static SitePower GetSitePower(DateTime startDate, DateTime endDate, SolarEdgeDetails solarEdgeDetails)
        {
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/power?api_key={solarEdgeDetails.SolarEdgeAPIKey}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetData(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            SitePower sitePower = JSONUtilities.DeserialiseJSONToClass<SitePower>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            return sitePower;
        }
    }

    public class DetailedBatteryPower
    {
        public StorageData storageData { get; set; }
        public static string ErrorMessage { get; set; }

        public static List<DetailedBatteryPower> GetDetailedBatteryPower(DateTime startDate, DateTime endDate, bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            fromFile = true;
            List<DetailedBatteryPower> returnList = new List<DetailedBatteryPower>();

            // First attempt to retrieve all the data from archive files if NOT in live mode
            if (!liveUpdateMode)
            {
                List<DetailedBatteryPower> detailedBatteryPowerList = GetDetailedBatteryPowerFromFile(startDate, endDate, dataPath);
                if (detailedBatteryPowerList != null && detailedBatteryPowerList.Count > 0)
                    return detailedBatteryPowerList;
            }

            // Otherwise request the data from SolarEdge
            fromFile = false;
            ErrorMessage = "";
            string http = $"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/storageData?api_key={solarEdgeDetails.SolarEdgeAPIKey}" +
                $"&startTime={startDate:yyyy-MM-dd}%20{startDate:HH:mm:ss}&endTime={endDate:yyyy-MM-dd}%20{endDate:HH:mm:ss}";
            Data.LastErrorMessage = "";
            string json = Data.GetData(http);
            ErrorMessage = Data.LastErrorMessage;
            if (json == null)
                return null;
            DetailedBatteryPower detailedBatteryPower = JSONUtilities.DeserialiseJSONToClass<DetailedBatteryPower>(json);
            if (JSONUtilities.LastErrorMessage != "")
            {
                ErrorMessage = JSONUtilities.LastErrorMessage;
                return null;
            }
            returnList.Add(detailedBatteryPower);
            return returnList;
        }

        private static List<DetailedBatteryPower> GetDetailedBatteryPowerFromFile(DateTime startDate, DateTime endDate, string dataPath)
        {
            List<DetailedBatteryPower> returnList = new List<DetailedBatteryPower>();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                string fileName = dataPath + $"\\Battery {startDate.Day}-{startDate.Month}-{startDate.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    DetailedBatteryPower detailedBatteryPower = JSONUtilities.DeserialiseJSONToClass<DetailedBatteryPower>(json);
                    returnList.Add(detailedBatteryPower);
                }
                startDate += TimeSpan.FromDays(1);
            }
            return returnList;
        }
    }

    public class SiteDetails
    {
        public Details details { get; set; }
        public static string ErrorMessage { get; set; }

        public static SiteDetails GetSiteDetails(SolarEdgeDetails solarEdgeDetails)
        {
            string json = Data.GetData($"https://monitoringapi.solaredge.com/site/{solarEdgeDetails.SiteID}/details?api_key={solarEdgeDetails.SolarEdgeAPIKey}");
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
    }

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
    }

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
        public string UPSServer = "http://10.1.1.3:3052/agent/ppbe.js/init_status.js?s=1530239977045";
        public string DataPath { get; set; } = @"C:\Solar Inverter";
        public string NephiSysIP { get; set; } = "10.1.1.3";
        public string CloudDataPath { get; set; } = @"N:\NephiSys\Archive\";
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
            public double PV;
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
            public double PV;
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

    public class DetailedDailySiteData
    {
        public static string ErrorMessage { get; set; }

        public static string GetDetailedSitePower(DateTime startDate, DateTime endDate,
            bool liveUpdateMode, SolarEdgeDetails solarEdgeDetails, string dataPath, out bool fromFile)
        {
            string csv = "";
            fromFile = false;
            ErrorMessage = "";
            DateTime epoch = new DateTime(1970, 1, 1);
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                DateTime today = startDate + TimeSpan.FromDays(i);
                long startDateJs = today.Subtract(epoch).Ticks / 10000;
                DateTime jsEnd = new DateTime(today.Year, today.Month, today.Day, 23, 59, 0);
                long endDateJs = jsEnd.Subtract(epoch).Ticks / 10000;
                string http = $"https://monitoringpublic.solaredge.com/solaredge-web/p/charts/{solarEdgeDetails.SiteID}/chartExport?" + 
                              $"st={startDateJs}&et={endDateJs}&fid={solarEdgeDetails.SiteID}&timeUnit=2&pn0=Power&id0=0&t0=0&hasMeters=false";
                Data.LastErrorMessage = "";
                csv = Data.GetData(http);
                ErrorMessage = Data.LastErrorMessage;
                if (csv == null)
                    return null;
                if (JSONUtilities.LastErrorMessage != "")
                {
                    ErrorMessage = JSONUtilities.LastErrorMessage;
                    return null;
                }
            }

            return csv;
        }

        private static List<DetailedDailySiteData> GetDetailedSiteDataFromFile(DateTime startDate, DateTime endDate,
            string dataPath)
        {
            List<DetailedDailySiteData> returnList = new List<DetailedDailySiteData>();
            TimeSpan span = endDate - startDate;
            int numDays = (int) span.TotalDays;
            for (int i = 0; i < numDays; i++)
            {
                DateTime today = startDate + TimeSpan.FromDays(i);
                string fileName = dataPath + $"\\DailySiteData {today.Day}-{today.Month}-{today.Year}.json";
                if (File.Exists(fileName))
                {
                    string json = File.ReadAllText(fileName);
                    DetailedDailySiteData detailedSitePower =
                        JSONUtilities.DeserialiseJSONToClass<DetailedDailySiteData>(json);
                    returnList.Add(detailedSitePower);
                }
            }

            return returnList;
        }
    }

    #endregion Main Classes

   #region UPS Classes
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
        public string UPSServer { get; set; } = "http://10.1.1.3:3052/agent/ppbe.js/init_status.js?s=1530239977045";
    }
    #endregion
}
