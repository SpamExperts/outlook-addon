using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.IO;
using System.Xml;

namespace reportmail
{
    class Program
    {
        static void ParseSettings(string settings_file, out string url, out string req_type, out string field, out string user, out string pass, out string response)
        {
            XmlDocument docXML = new XmlDocument();
            docXML.Load(settings_file);

            XmlNode nodeDoc = docXML.DocumentElement;
            XmlNode nodeURL = nodeDoc.FirstChild;
            XmlNode nodeReqType = nodeURL.NextSibling;
            XmlNode nodeField = nodeReqType.NextSibling;
            XmlNode nodeUser = nodeField.NextSibling;
            XmlNode nodePass = nodeUser.NextSibling;
            XmlNode nodeResponse = nodePass.NextSibling;

            url = nodeURL.InnerText;
            req_type = nodeReqType.InnerText;
            field = nodeField.InnerText;
            user = nodeUser.InnerText;
            pass = nodePass.InnerText;
            response = nodeResponse.InnerText;
        }

        static void Main(string[] args)
        {
            string dirCurr = AppDomain.CurrentDomain.BaseDirectory;

            try
            {
                string url, reqType, field, user, pass, response;
                ParseSettings(Path.Combine(dirCurr, "settings.xml"), 
                    out url, out reqType, out field, out user, out pass, out response);
                string data = field + "=";

#if DEBUG
                Console.WriteLine(
                    "Settings( URL: {0}, RequestType: {1}, FieldName: {2}, Username: {3}, Password: {4}, Response: {5} )",
                     url, reqType, field, user, pass, response);
                Console.ReadLine();
#endif

                string sEml = "";
                string strRes = "";
                using ( TextReader tr = new StreamReader(args[0]) )
                {
                    sEml = tr.ReadToEnd();
                }
                data += System.Web.HttpUtility.UrlEncode(sEml);

                //Our postvars
                byte[] buffer = Encoding.ASCII.GetBytes(data);
                //Initialisation, we use localhost, change if appliable
                HttpWebRequest WebReq = (HttpWebRequest)WebRequest.Create(url);
                //Our method is post, otherwise the buffer (postvars) would be useless
                WebReq.Method = reqType;
                //Set credentials
                if (user.Length > 0 && pass.Length > 0)
                    WebReq.Credentials = new NetworkCredential(user, pass);
                //We use form contentType, for the postvars.
                WebReq.ContentType = "application/x-www-form-urlencoded";
                //The length of the buffer (postvars) is used as contentlength.
                WebReq.ContentLength = buffer.Length;
                //We open a stream for writing the postvars
                Stream PostData = WebReq.GetRequestStream();

                //Now we write, and afterwards, we close.
                PostData.Write(buffer, 0, buffer.Length);
                PostData.Close();

                HttpWebResponse WebResp = (HttpWebResponse)WebReq.GetResponse();

                Stream res = WebResp.GetResponseStream();
                using ( TextReader tr = new StreamReader(res) )
                {
                    strRes = tr.ReadLine();
                }

                // write the response to the file
                using ( BinaryWriter bw = new BinaryWriter(File.Open(args[0], FileMode.Truncate)) )
                {
#if DEBUG
                    Console.WriteLine("ResponseStream: {0}, ResponseSettings: {1}", strRes, response);
#endif

                    if (strRes.StartsWith(response))
                        bw.Write(respOK);
                    else
                        bw.Write(respFail);
                }

#if DEBUG
                Console.WriteLine("Response: {0}", strRes);
#endif
            }
            catch (Exception ex)
            {
                using (TextWriter logW = new StreamWriter(Path.Combine(dirCurr, "error_log.txt"), true))
                {
                    StringBuilder sb = new StringBuilder("[");
                    sb.Append(DateTime.Now.ToString());
                    sb.Append("] Error: ");
                    sb.Append(ex.Message);

                    logW.WriteLine(sb.ToString());
                }

#if DEBUG
                Console.WriteLine(ex.Message);
#endif
            }
#if DEBUG
            Console.ReadLine();
#endif
        }

        private static readonly int respOK = 1;
        private static readonly int respFail = 0;
    }
}
