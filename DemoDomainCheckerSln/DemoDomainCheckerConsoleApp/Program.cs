using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace DemoDomainCheckerConsoleApp
{
    class Program
    {
        private static string O365Password { get; set; } = Properties.Settings.Default.O365Password;
        private static string O365Url { get; set; } = Properties.Settings.Default.O365Url;
        private static string O365Username { get; set; } = Properties.Settings.Default.O365Username;
        private const string USER_AGENT = "Mozilla/5.0 (compatible, MSIE 11, Windows NT 6.3; Trident/7.0;  rv:11.0)";

        static ClientContext GetContext()
        {
            SecureString securePassword = new SecureString();
            foreach (char c in O365Password)
            {
                securePassword.AppendChar(c);
            }
            SharePointOnlineCredentials credentials = new SharePointOnlineCredentials(O365Username, securePassword);
            ClientContext ctx = new ClientContext(O365Url);
            ctx.Credentials = credentials;
            return ctx;
        }
        static void Main(string[] args)
        {
            Console.WriteLine("Application Started");
            Uri filename = new Uri(Properties.Settings.Default.SourceFileUri);
            string server = filename.AbsoluteUri.Replace(filename.AbsolutePath, "");
            string serverrelative = filename.AbsolutePath;
            Microsoft.SharePoint.Client.ClientContext clientContext = GetContext();
            Microsoft.SharePoint.Client.FileInformation f = Microsoft.SharePoint.Client.File.OpenBinaryDirect(clientContext, serverrelative);
            var tempPath = System.IO.Path.GetTempPath();
            string filePath = Path.Combine(tempPath, "Empresas Por Tamaño Temp.xlsx");
            using (var fileStream = new FileStream(filePath, FileMode.Create))
            {
                //check http://stackoverflow.com/questions/2624333/how-do-i-read-data-from-a-spreadsheet-using-the-openxml-format-sdk
                //check https://www.helloitsliam.com/2016/01/28/c-console-application-and-office-365-2/?utm_content=bufferb4e9c&utm_medium=social&utm_source=twitter.com&utm_campaign=buffer
                //check http://sharepoint.stackexchange.com/questions/62087/how-to-get-a-file-using-sharepoint-client-object-model-with-only-an-absolute-url
                f.Stream.CopyTo(fileStream);
            }
            ReadFile(filePath);
            CheckWebsitesStatus();
            Console.WriteLine("Application Finised");
            Console.WriteLine("Press any key to quit");
            Console.ReadKey();
        }

        private static void CheckWebsitesStatus()
        {
            System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            foreach (var singleClientInfo in objClientsInfo)
            {
                try
                {
                    watch.Start();
                    HttpWebRequest objWebRequest = (HttpWebRequest)WebRequest.Create(singleClientInfo.Url);
                    objWebRequest.UserAgent = USER_AGENT;
                    HttpWebResponse objHttpWebResponse = (HttpWebResponse)objWebRequest.GetResponse();
                    watch.Stop();
                    Console.WriteLine(string.Format("Checked Url: {0}. Response: {1}. Description: {2}. Time:{3}", singleClientInfo.Url, objHttpWebResponse.StatusCode,
                        objHttpWebResponse.StatusDescription, watch.Elapsed.ToString()));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    if (watch.IsRunning)
                        watch.Stop();
                    watch.Reset();
                }
            }
        }

        public class ClientInfo
        {
            public string Domain { get; set; }
            public Uri Url { get; set; }
            public string Email { get; set; }
        }
        private static List<ClientInfo> objClientsInfo = new List<ClientInfo>();
        private static void ReadFile(string excelFilePath)
        {
            using (CsvHelper.CsvReader csv = new CsvHelper.CsvReader(new CsvHelper.Excel.ExcelParser(excelFilePath)))
            {
                while (csv.Read())
                {
                    string url = csv.GetField("Sitio Web");
                    string correoPrincipal = csv.GetField("Correo principal");
                    if (!string.IsNullOrWhiteSpace(url) && url.ToLower() != "n/a")
                    {
                        try
                        {
                            if (url.ToLower().IndexOf("http") == -1)
                                url = string.Format("http://{0}", url);
                            Uri uri = new Uri(url);
                            string domain = uri.Host;
                            objClientsInfo.Add(
                                new ClientInfo()
                                {
                                    Domain = domain,
                                    Url = uri,
                                    Email = correoPrincipal
                                }
                                );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.ToString());
                        }
                    }
                }
            }
        }
    }
}
