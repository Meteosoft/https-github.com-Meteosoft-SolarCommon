using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace Meteosoft.SolarCommon
{
    public static class Data
    {
        public static string LastErrorMessage { get; set; } = "";
        //private static readonly HttpClient m_client = new HttpClient();

        public static string GetData(string http)
        {
            X509Certificate cert = X509Certificate.CreateFromCertFile(@"C:\Users\rob\Documents\Personal\PV System\solaredgecom.crt");
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(http);
            request.ClientCertificates.Add(cert);
            ServicePointManager.ServerCertificateValidationCallback = ServerCertificateValidationCallback;
            //request.ClientCertificates.Add(new X509Certificate2(@"C:\Users\rob\Documents\Personal\PV System\solaredgecom.crt", "Tr3dg0ld"));

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
                LastErrorMessage = e.Message;
            }
            return null;
        }

        private static bool ServerCertificateValidationCallback(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslpolicyerrors)
        {
            return false;
        }


        
        
        //public static string PostData(string ipAddress, string method, string parameters)
        //{
        //    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(ipAddress + "/" + method); // e.g. http://api.alphaess.com/ras/v2
        //    byte[] data = Encoding.ASCII.GetBytes(parameters);

        //    request.Method = "POST";
        //    request.ContentType = "application/x-www-form-urlencoded";
        //    request.ContentLength = data.Length;

        //    using (Stream stream = request.GetRequestStream())
        //    {
        //        stream.Write(data, 0, data.Length);
        //    }

        //    string responseString = "";
        //    try
        //    {
        //        using (HttpWebResponse response = (HttpWebResponse) request.GetResponse())
        //        {
        //            responseString =
        //                new StreamReader(response.GetResponseStream() ?? throw new InvalidOperationException()).ReadToEnd();
        //        }
        //    }
        //    catch (Exception e)
        //    {
        //        LastErrorMessage = e.Message;
        //    }

        //    return responseString;
        //}
    }
}
