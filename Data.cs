using System;
using System.IO;
using System.Net;
using System.Text;

namespace Meteosoft.SolarCommon
{
    public static class Data
    {
        public static string LastErrorMessage { get; set; } = "";
        public static string Cookie { get; set; } = "";

        public static string GetDataDirect(string http)
        {
            LastErrorMessage = "";
            UriBuilder uriBuilder = new UriBuilder(http);
            CookieContainer cookieContainer = new CookieContainer();
            cookieContainer.SetCookies(uriBuilder.Uri, Cookie);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(http);
            request.Timeout = 10000;
            request.AllowAutoRedirect = true;
            request.Referer = "https://monitoring.solaredge.com/solaredge-web/p/site/394525/";
            request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134";
            request.CookieContainer = cookieContainer;
            request.Host = "monitoring.solaredge.com";
            request.Accept = "*/*";
            request.KeepAlive = true;

            try
            {
                HttpWebResponse response = (HttpWebResponse) request.GetResponse();
                Stream data = response.GetResponseStream();
                if (data != null)
                {
                    StreamReader reader = new StreamReader(data);
                    string result = reader.ReadToEnd();
                    reader.Close();
                    data.Close();
                    response.Close();
                    return result;
                }
            }
            catch (Exception e)
            {
                LastErrorMessage = e.Message + "[" + http + "]";
            }
            return null;
        }

        public static string GetDataUsingAPI(string http)
        {
            LastErrorMessage = "";
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(http);
            request.Timeout = 10000;
            request.AllowAutoRedirect = true;

            try
            {
                HttpWebResponse response = (HttpWebResponse) request.GetResponse();
                Stream data = response.GetResponseStream();
                if (data != null)
                {
                    StreamReader reader = new StreamReader(data);
                    string result = reader.ReadToEnd();
                    reader.Close();
                    data.Close();
                    response.Close();
                    return result;
                }
            }
            catch (Exception e)
            {
                LastErrorMessage = e.Message + "[" + http + "]";
            }
            return null;
        }

        public static string PostData(string ipAddress, string method, string parameters)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ipAddress + method); // e.g. http://api.alphaess.com/ras/v2
            byte[] data = Encoding.ASCII.GetBytes(parameters);

            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = data.Length;

            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(data, 0, data.Length);
            }

            string responseString = "";
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    responseString =
                        new StreamReader(response.GetResponseStream() ?? throw new InvalidOperationException()).ReadToEnd();
                }
            }
            catch (Exception e)
            {
                LastErrorMessage = e.Message;
            }

            return responseString;
        }
    }
}
